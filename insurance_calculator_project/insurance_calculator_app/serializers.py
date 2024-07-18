from datetime import datetime

from rest_framework import serializers

from insurance_calculator_app.utils import to_snake_case

INSURANCE_TYPE_CHOICES = [
    'чистое дожитие',
    'страхование жизни на срок',
    'чисто накопительное страхование',
    'пожизненное страхование',
]

PAYMENT_FREQUENCY_CHOICES = [
    'единовременно',
    'ежегодно',
    'ежемесячно'
]

GENDER_CHOICES = [
    'мужской',
    'женский'
]


class CalculatorInputSerializer(serializers.Serializer):
    insurance_type = serializers.ChoiceField(choices=INSURANCE_TYPE_CHOICES)
    insurance_premium_frequency = serializers.ChoiceField(choices=PAYMENT_FREQUENCY_CHOICES)
    birth_date = serializers.DateField()
    insurance_start_date = serializers.DateField()
    insurance_end_date = serializers.DateField()
    gender = serializers.ChoiceField(choices=GENDER_CHOICES)
    insurance_premium_rate = serializers.FloatField()
    insurance_premium_supplement = serializers.FloatField(min_value=0)
    # insurance_premium = serializers.FloatField()
    insurance_premium = serializers.FloatField(default=None, allow_null=True)
    insurance_sum = serializers.FloatField(default=None, allow_null=True)

    def validate(self, data):
        """
        Check calculator input.
        """

        if data['birth_date'] > datetime.today().date():
            raise serializers.ValidationError('Birth date can\'t be later than current moment.')

        if data['birth_date'] > data['insurance_start_date']:
            raise serializers.ValidationError('Birth date can\'t be later than insurance start date.')

        if data['insurance_start_date'] >= data['insurance_end_date']:
            raise serializers.ValidationError('Insurance end date must be later than insurance start date.')

        if data['insurance_premium_supplement'] >= 1:
            raise serializers.ValidationError('Insurance premium supplement must be less than 1.')

        if data['insurance_premium'] is not None and data['insurance_premium'] <= 0:
            raise serializers.ValidationError('Insurance premium must be greater than 0.')

        # if data['insurance_premium'] <= 0:
        #     raise serializers.ValidationError('Insurance premium must be greater than 0.')

        if data['insurance_premium'] is not None and data['insurance_premium'] <= 0:
            raise serializers.ValidationError('Insurance premium must be greater than 0.')

        if data['insurance_sum'] is not None and data['insurance_sum'] <= 0:
            raise serializers.ValidationError('Insurance sum must be greater than 0.')

        return data

    def to_internal_value(self, data):
        # Convert data fields that came from react frontend from camel case to snake case format expected by serializer
        converted_data = {to_snake_case(k): v for k, v in data.items()}
        return super().to_internal_value(converted_data)
