import datetime as dt

from django.test import SimpleTestCase

from insurance_calculator_app.serializers import (PremiumCalculatorInputSerializer, BaseCalculatorInputSerializer,
                                                  IntermediateCalculatorInputSerializer, SumCalculatorInputSerializer,
                                                  ReserveCalculatorInputSerializer, TariffCalculatorInputSerializer)
from insurance_calculator_app.utils import get_default_request_base_data, get_default_request_data, \
    get_default_tariffs_request_data


class CalculatorSerializerTest(SimpleTestCase):

    def check_valid_data_processing(self, data, serializer_type=BaseCalculatorInputSerializer):
        with self.subTest(**data):
            serializer = serializer_type(data=data)
            self.assertTrue(serializer.is_valid())

    def check_invalid_data_processing(self, data, expected_errors, serializer_type=BaseCalculatorInputSerializer):
        with self.subTest(**data):
            serializer = serializer_type(data=data)
            serializer.is_valid()
            self.assertEqual(expected_errors, serializer.errors)

    def check_optional_field_processing(self, field, base_data, serializer_type=BaseCalculatorInputSerializer,
                                       insurance_type='cumulative insurance'):
        # test validating fields that are optional for some insurance types
        expected_errors = {
            'non_field_errors': [f'"{field}" field is required for all insurance types except {insurance_type}']}
        data = {**base_data, field: None}
        self.check_invalid_data_processing(data, expected_errors, serializer_type)

    def test_base_calculator_serializer(self):
        base_data = get_default_request_base_data()
        data = {**base_data}
        self.check_valid_data_processing(data)

        # test field conversion from camel case to pascal case
        data = {**get_default_request_base_data(False)}
        self.check_valid_data_processing(data)

        expected_errors = {
            'non_field_errors': ['Insurance loading must be greater than or equal to 0 and less than 1.']}
        data = {**base_data, 'insurance_loading': -0.1}
        self.check_invalid_data_processing(data, expected_errors)
        data['insurance_loading'] = 1
        self.check_invalid_data_processing(data, expected_errors)

        self.check_optional_field_processing('gender', base_data)

    def test_intermediate_calculator_serializer(self):
        base_data = get_default_request_data()
        data = {**base_data}
        self.check_valid_data_processing(data, IntermediateCalculatorInputSerializer)

        self.check_optional_field_processing('birth_date', base_data, IntermediateCalculatorInputSerializer)
        self.check_optional_field_processing('insurance_start_date', base_data, IntermediateCalculatorInputSerializer)
        self.check_optional_field_processing('insurance_period', base_data, IntermediateCalculatorInputSerializer,
                                            'whole life insurance')
        expected_errors = {
            'non_field_errors': ['Birth date can\'t be later than current moment.']}
        data = {**base_data, 'birth_date': (dt.datetime.today() + dt.timedelta(days=1)).date().strftime('%Y-%m-%d')}
        self.check_invalid_data_processing(data, expected_errors, IntermediateCalculatorInputSerializer)
        expected_errors = {
            'non_field_errors': ['Birth date can\'t be later than insurance start date.']}
        data['birth_date'] = '2024-08-25'
        self.check_invalid_data_processing(data, expected_errors, IntermediateCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Insurance period must be greater than 0.']}
        data = {**base_data, 'insurance_period': 0}
        self.check_invalid_data_processing(data, expected_errors, IntermediateCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Age of insured person at the end of insurance period can\'t be greater than 101.']}
        data = {**base_data, 'birth_date': '2000-01-01', 'insurance_start_date': '2021-01-02', 'insurance_period': 960}
        self.check_invalid_data_processing(data, expected_errors, IntermediateCalculatorInputSerializer)

    def test_premium_calculator_serializer(self):
        base_data = get_default_request_data()
        data = {**base_data, 'insurance_sum': 10000}
        self.check_valid_data_processing(data, PremiumCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Insurance sum must be greater than 0.']}
        data = {**base_data, 'insurance_sum': 0}
        self.check_invalid_data_processing(data, expected_errors, PremiumCalculatorInputSerializer)

    def test_sum_calculator_serializer(self):
        base_data = get_default_request_data()
        data = {**base_data, 'insurance_premium': 100}
        self.check_valid_data_processing(data, SumCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Insurance premium must be greater than 0.']}
        data = {**base_data, 'insurance_premium': 0}
        self.check_invalid_data_processing(data, expected_errors, SumCalculatorInputSerializer)

    def test_reserve_calculator_serializer(self):
        base_data = {**get_default_request_data(), 'reserve_calculation_period': 38}
        data = {**base_data, 'insurance_sum': 10000, 'insurance_loading': None}
        self.check_valid_data_processing(data, ReserveCalculatorInputSerializer)
        data = {**base_data, 'insurance_premium': 100}
        self.check_valid_data_processing(data, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Either "insurance_premium" or "insurance_sum" must be provided.']}
        data = {**base_data}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                '"insurance_loading" field is required for reserve calculation using insurance premium.']}
        data = {**base_data, 'insurance_premium': 100, 'insurance_loading': None}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Insurance sum must be greater than 0.']}
        data = {**base_data, 'insurance_sum': 0}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': ['Insurance premium must be greater than 0.']}
        data = {**base_data, 'insurance_premium': 0}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                'Time from start of insurance to insurance reserve calculation must be greater than 0.']}
        data = {**base_data, 'reserve_calculation_period': 0, 'insurance_sum': 10000}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                'Time from start of insurance to insurance reserve calculation must be less than insurance period.']}
        data = {**base_data, 'reserve_calculation_period': 51, 'insurance_period': 50, 'insurance_sum': 10000}
        self.check_invalid_data_processing(data, expected_errors, ReserveCalculatorInputSerializer)

    def test_tariffs_calculator_serializer(self):
        base_data = get_default_tariffs_request_data()
        data = {**base_data}
        self.check_valid_data_processing(data, TariffCalculatorInputSerializer)

        self.check_optional_field_processing('minimum_insurance_start_age', base_data, TariffCalculatorInputSerializer)
        self.check_optional_field_processing('maximum_insurance_start_age', base_data, TariffCalculatorInputSerializer)
        self.check_optional_field_processing('maximum_insurance_period', base_data, TariffCalculatorInputSerializer,
                                            'whole life insurance')

        expected_errors = {
            'non_field_errors': [
                'Minimum age of insurance start can\'t be greater than maximum age of insurance start.']}
        data = {**base_data, 'minimum_insurance_start_age': 21, 'maximum_insurance_start_age': 20}
        self.check_invalid_data_processing(data, expected_errors, TariffCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                'Maximum age of insurance start can\'t be greater than 100.']}
        data = {**base_data, 'maximum_insurance_start_age': 101}
        self.check_invalid_data_processing(data, expected_errors, TariffCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                'Maximum insurance period must be greater than 0.']}
        data = {**base_data, 'maximum_insurance_period': 0}
        self.check_invalid_data_processing(data, expected_errors, TariffCalculatorInputSerializer)

        expected_errors = {
            'non_field_errors': [
                'Sum of maximum insurance start age and maximum insurance period can\'t be greater than 101 year.']}
        data = {**base_data, 'maximum_insurance_start_age': 92, 'maximum_insurance_period': 120}
        self.check_invalid_data_processing(data, expected_errors, TariffCalculatorInputSerializer)
