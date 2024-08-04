from abc import ABC
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


class BaseCalculatorInputSerializer(serializers.Serializer):
    insurance_type = serializers.ChoiceField(choices=INSURANCE_TYPE_CHOICES)
    insurance_premium_frequency = serializers.ChoiceField(choices=PAYMENT_FREQUENCY_CHOICES)
    gender = serializers.ChoiceField(choices=GENDER_CHOICES)
    insurance_premium_rate = serializers.FloatField(min_value=0)
    insurance_loading = serializers.FloatField()

    def validate(self, data):
        """
        Check calculator input.
        """

        if data['insurance_loading'] < 0 or data['insurance_loading'] >= 1:
            raise serializers.ValidationError(
                'Insurance loading must be greater than or equal to 0 and less than 1.')

        return data

    def to_internal_value(self, data):
        # Convert data fields that come from react frontend from camel case to snake case format expected by serializer
        converted_data = {to_snake_case(k): v for k, v in data.items()}
        return super().to_internal_value(converted_data)


class TariffCalculatorInputSerializer(BaseCalculatorInputSerializer):
    insurance_start_age = serializers.IntegerField(min_value=0)
    insurance_end_age = serializers.IntegerField(min_value=0)
    maximum_insurance_period = serializers.IntegerField()

    def validate(self, data):
        """
        Check calculator input.
        """

        super().validate(data)

        if (data['insurance_start_age'] and data['insurance_end_age']) and data['insurance_start_age'] > data['insurance_end_age']:
            raise serializers.ValidationError('Age of insurance start can\'t be greater than age of insurance end.')

        if data['maximum_insurance_period'] <= 0:
            raise serializers.ValidationError('Maximum insurance period must be greater than 0.')

        return data


class IntermediateCalculatorInputSerializer(BaseCalculatorInputSerializer):
    birth_date = serializers.DateField(default=None)
    insurance_start_date = serializers.DateField(default=None)
    insurance_period = serializers.IntegerField(default=None)

    def validate(self, data):
        """
        Check calculator input.
        """

        super().validate(data)

        if data['birth_date'] and data['birth_date'] > datetime.today().date():
            raise serializers.ValidationError('Birth date can\'t be later than current moment.')

        if (data['birth_date'] and data['insurance_start_date']) and data['birth_date'] > data['insurance_start_date']:
            raise serializers.ValidationError('Birth date can\'t be later than insurance start date.')

        if data['insurance_period'] is not None and data['insurance_period'] <= 0:
            raise serializers.ValidationError('Insurance period must be greater than 0.')

        return data


class PremiumCalculatorInputSerializer(IntermediateCalculatorInputSerializer):
    insurance_sum = serializers.FloatField()

    def validate(self, data):
        """
        Check calculator input.
        """

        super().validate(data)

        if data['insurance_sum'] <= 0:
            raise serializers.ValidationError('Insurance sum must be greater than 0.')

        return data


class SumCalculatorInputSerializer(IntermediateCalculatorInputSerializer):
    insurance_premium = serializers.FloatField()

    def validate(self, data):
        """
        Check calculator input.
        """

        super().validate(data)

        if data['insurance_premium'] <= 0:
            raise serializers.ValidationError('Insurance premium must be greater than 0.')

        return data


class ReserveCalculatorInputSerializer(IntermediateCalculatorInputSerializer):
    insurance_premium = serializers.FloatField(default=None)
    insurance_sum = serializers.FloatField(default=None)
    reserve_calculation_period = serializers.IntegerField()

    def validate(self, data):
        """
        Check calculator input.
        """

        super().validate(data)

        if data['insurance_premium'] is not None and data['insurance_premium'] <= 0:
            raise serializers.ValidationError('Insurance premium must be greater than 0.')

        if data['insurance_sum'] is not None and data['insurance_sum'] <= 0:
            raise serializers.ValidationError('Insurance sum must be greater than 0.')

        if data['insurance_premium'] is None and data['insurance_sum'] is None:
            raise serializers.ValidationError('Insurance premium or insurance sum must be specified.')

        if data['reserve_calculation_period'] == 0:
            raise serializers.ValidationError(
                'Time from start of insurance to insurance reserve calculation must be greater than 0.')

        if data['reserve_calculation_period'] >= data['insurance_period']:
            raise serializers.ValidationError(
                'Time from start of insurance to insurance reserve calculation must be less than insurance period.')

        return data
