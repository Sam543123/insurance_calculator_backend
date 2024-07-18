from datetime import datetime

from rest_framework import serializers

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

        return data
