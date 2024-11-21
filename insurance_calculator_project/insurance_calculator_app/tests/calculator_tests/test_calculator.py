import datetime as dt
import itertools
from pathlib import Path

import openpyxl
from django.test import TestCase

from insurance_calculator_app.insurance_calculator import InsuranceCalculator
from insurance_calculator_app.serializers import INSURANCE_TYPE_CHOICES, PAYMENT_FREQUENCY_CHOICES
from insurance_calculator_app.tests.test_utils import compare_excel_files


class CalculatorTest(TestCase):

    @staticmethod
    def get_default_base_common_params():
        params = {'insurance_type': 'term life insurance',
                  'insurance_premium_frequency': 'annually',
                  'technical_interest_rate': 0.05,
                  'insurance_loading': 0.2, 'gender': 'male'}
        return params

    @staticmethod
    def get_default_base_params():
        base_common_params = CalculatorTest.get_default_base_common_params()
        params = {**base_common_params, 'birth_date': dt.date(1995, 10, 21),
                  'insurance_start_date': dt.date(2024, 8, 24),
                  'insurance_period': 69}
        return params

    @staticmethod
    def get_default_base_tariffs_params():
        base_common_params = CalculatorTest.get_default_base_common_params()
        params = {**base_common_params,
                  'minimum_insurance_start_age': 19,
                  'maximum_insurance_start_age': 28,
                  'maximum_insurance_period': 72}
        return params

    @staticmethod
    def update_params(params):
        # update parameters for calculation method depending on insurance type
        new_params = {**params}
        insurance_type = params['insurance_type']
        if insurance_type == 'cumulative insurance':
            new_params['gender'] = None
            new_params['birth_date'] = None
            new_params['insurance_start_date'] = None
        if insurance_type == 'whole life insurance':
            new_params['insurance_period'] = None
        return new_params

    @staticmethod
    def update_tariffs_params(params):
        # update parameters for tariffs calculation depending on insurance type
        new_params = {**params}
        insurance_type = params['insurance_type']
        if insurance_type == 'cumulative insurance':
            new_params['gender'] = None
            new_params['minimum_insurance_start_age'] = None
            new_params['maximum_insurance_start_age'] = None
        if insurance_type == 'whole life insurance':
            new_params['maximum_insurance_period'] = None
        return new_params

    @staticmethod
    def get_insurance_premium(insurance_premium_frequency):
        # get insurance premium for testing depending on insurance premium frequency
        if insurance_premium_frequency == 'simultaneously':
            insurance_premium = 1000
        elif insurance_premium_frequency == 'annually':
            insurance_premium = 100
        else:
            insurance_premium = 10
        return insurance_premium

    def check_result(self, params, expected_result, calculation_method_name='calculate_premium'):
        updated_params = self.update_params(params)
        with self.subTest(**updated_params):
            calculation_method = getattr(InsuranceCalculator, calculation_method_name)
            result = calculation_method(**updated_params)
            self.assertEqual(result, expected_result)

    def check_tariffs_result(self, params, expected_tariffs_file_path):
        updated_params = self.update_tariffs_params(params)
        with self.subTest(**updated_params):
            result = InsuranceCalculator.calculate_tariffs(**updated_params)
            self.assertTrue(compare_excel_files(result, expected_tariffs_file_path))

    def check_translation_tariffs_table(self, params, expected_results_dict):
        updated_params = self.update_tariffs_params(params)
        with self.subTest(**updated_params):
            result = InsuranceCalculator.calculate_tariffs(**updated_params)
            workbook = openpyxl.load_workbook(result, data_only=True)
            work_sheet = workbook.active
            for cell, expected_result in expected_results_dict.items():
                self.assertEqual(work_sheet[cell].value, expected_result)

    def test_insurance_premium_calculation(self):
        insurance_sum = 10000
        base_params = {**self.get_default_base_params(), 'insurance_sum': insurance_sum}
        # test premium calculation for different combinations of insurance type and premium frequency
        expected_results = (9339.22019, 1759.54811, 155.6392,
                            117.71554, 22.1781, 1.96174,
                            2140.23723, 122.36514, 10.47412,
                            9442.16409, 1771.68692, 156.61589)
        expected_results_iterator = iter(expected_results)
        params = {**base_params}
        for insurance_type, insurance_premium_frequency in itertools.product(INSURANCE_TYPE_CHOICES,
                                                                             PAYMENT_FREQUENCY_CHOICES):
            params['insurance_type'] = insurance_type
            params['insurance_premium_frequency'] = insurance_premium_frequency
            self.check_result(params, next(expected_results_iterator))

        # test premium calculation for zero insurance premium rate with different insurance types
        expected_results = (2069.59742, 22.84189, 299.1907, 2083.33333)
        expected_results_iterator = iter(expected_results)
        params = {**base_params, 'technical_interest_rate': 0}
        for insurance_type in INSURANCE_TYPE_CHOICES:
            params['insurance_type'] = insurance_type
            self.check_result(params, next(expected_results_iterator))

        # test premium calculation for zero insurance loading
        expected_result = 17.74248
        params = {**base_params, 'insurance_loading': 0}
        self.check_result(params, expected_result)

        # test premium calculation for female gender
        expected_result = 6.74024
        params = {**base_params, 'gender': 'female'}
        self.check_result(params, expected_result)

    def test_insurance_sum_calculation(self):
        base_params = {**self.get_default_base_params()}
        # test insurance sum calculation for different combinations of insurance type and premium frequency
        expected_results = (1070.75321, 568.32774, 642.51165,
                            84950.54821, 45089.5243, 50975.06718,
                            4672.37924, 8172.2619, 9547.34533,
                            1059.07924, 564.43381, 638.50481)
        expected_results_iterator = iter(expected_results)
        params = {**base_params}
        for insurance_type, insurance_premium_frequency in itertools.product(INSURANCE_TYPE_CHOICES,
                                                                             PAYMENT_FREQUENCY_CHOICES):
            params['insurance_type'] = insurance_type
            params['insurance_premium_frequency'] = insurance_premium_frequency
            params['insurance_premium'] = self.get_insurance_premium(insurance_premium_frequency)
            self.check_result(params, next(expected_results_iterator), 'calculate_insurance_sum')

        # test insurance sum calculation for zero insurance premium rate with different insurance types
        expected_results = (483.18576, 43779.20484, 3342.34984, 480)
        expected_results_iterator = iter(expected_results)
        params = {**base_params, 'technical_interest_rate': 0}
        for insurance_type in INSURANCE_TYPE_CHOICES:
            params['insurance_type'] = insurance_type
            params['insurance_premium'] = self.get_insurance_premium(params['insurance_premium_frequency'])
            self.check_result(params, next(expected_results_iterator), 'calculate_insurance_sum')

        # test insurance sum calculation for zero insurance loading
        expected_result = 56361.90538
        params = {**base_params, 'insurance_premium': 100, 'insurance_loading': 0}
        self.check_result(params, expected_result, 'calculate_insurance_sum')

        # test insurance sum calculation for female gender
        expected_result = 148362.56499
        params = {**base_params, 'insurance_premium': 100, 'gender': 'female'}
        self.check_result(params, expected_result, 'calculate_insurance_sum')

    def test_reserve_calculation(self):
        insurance_sum = 10000
        reserve_calculation_period = 38
        base_params = {**self.get_default_base_params(), 'reserve_calculation_period': reserve_calculation_period,
                       'insurance_premium': None, 'insurance_sum': None, 'insurance_loading': None}
        # test reserve calculation for different combinations of insurance type and premium frequency
        expected_results = (8766.51593, 6134.96807, 5141.89194,
                            52.39326, 17.8477, 6.70694,
                            1950.71844, 302.71153, 287.76192,
                            8815.78474, 6158.83232, 5159.0958)
        expected_results_iterator = iter(expected_results)
        params = {**base_params, 'insurance_sum': insurance_sum}
        for insurance_type, insurance_premium_frequency in itertools.product(INSURANCE_TYPE_CHOICES,
                                                                             PAYMENT_FREQUENCY_CHOICES):
            params['insurance_type'] = insurance_type
            params['insurance_premium_frequency'] = insurance_premium_frequency
            self.check_result(params, next(expected_results_iterator), 'calculate_reserve')

        # test reserve calculation for zero insurance premium rate with different insurance types
        expected_results = (6641.87767, 19.44067, 904.86721, 6666.66667)
        expected_results_iterator = iter(expected_results)
        params = {**base_params, 'insurance_sum': insurance_sum, 'technical_interest_rate': 0}
        for insurance_type in INSURANCE_TYPE_CHOICES:
            params['insurance_type'] = insurance_type
            self.check_result(params, next(expected_results_iterator), 'calculate_reserve')

        # test reserve calculation for female gender
        expected_result = 4.41027
        params = {**base_params, 'insurance_sum': insurance_sum, 'gender': 'female'}
        self.check_result(params, expected_result, 'calculate_reserve')

        # test reserve calculation using insurance premium instead of insurance sum
        expected_result = 80.47445
        params = {**base_params, 'insurance_premium': 100, 'insurance_loading': 0.2}
        self.check_result(params, expected_result, 'calculate_reserve')

        # test reserve calculation using insurance premium with zero insurance loading
        expected_result = 100.59306
        params = {**base_params, 'insurance_premium': 100, 'insurance_loading': 0}
        self.check_result(params, expected_result, 'calculate_reserve')

    def test_tariffs_calculation(self):
        base_expected_tariffs_folder_path = Path(__file__).parents[0] / 'expected_tariffs_files'
        params = self.get_default_base_tariffs_params()
        # test tariffs calculation for different combinations of insurance type and premium frequency
        expected_tariffs_folder_path = base_expected_tariffs_folder_path / 'different_cases'
        for insurance_type, insurance_premium_frequency in itertools.product(INSURANCE_TYPE_CHOICES,
                                                                             PAYMENT_FREQUENCY_CHOICES):
            params['insurance_type'] = insurance_type
            params['insurance_premium_frequency'] = insurance_premium_frequency
            expected_tariffs_file_path = expected_tariffs_folder_path / f'{insurance_type.replace(" ", "_")}_{insurance_premium_frequency}_tariffs.xlsx'
            self.check_tariffs_result(params, expected_tariffs_file_path)

        # test tariffs calculation for zero insurance premium rate with different insurance types
        expected_tariffs_folder_path = base_expected_tariffs_folder_path / 'zero_premium_rate_case'
        params = {**self.get_default_base_tariffs_params(), 'technical_interest_rate': 0}
        for insurance_type in INSURANCE_TYPE_CHOICES:
            params['insurance_type'] = insurance_type
            expected_tariffs_file_path = expected_tariffs_folder_path / f'{insurance_type.replace(" ", "_")}_tariffs.xlsx'
            self.check_tariffs_result(params, expected_tariffs_file_path)

        # test tariffs calculation for zero insurance loading
        expected_tariffs_folder_path = base_expected_tariffs_folder_path / 'zero_insurance_loading_case'
        params = {**self.get_default_base_tariffs_params(), 'insurance_loading': 0}
        expected_tariffs_file_path = expected_tariffs_folder_path / 'tariffs.xlsx'
        self.check_tariffs_result(params, expected_tariffs_file_path)

        # test tariffs calculation for female gender
        expected_tariffs_folder_path = base_expected_tariffs_folder_path / 'female_gender_case'
        params = {**self.get_default_base_tariffs_params(), 'gender': 'female'}
        expected_tariffs_file_path = expected_tariffs_folder_path / 'tariffs.xlsx'
        self.check_tariffs_result(params, expected_tariffs_file_path)

        # test tariffs calculation for cumulative insurance with non-whole number of years in maximum insurance period
        expected_tariffs_folder_path = base_expected_tariffs_folder_path / 'fractional_period_cumulative_insurance_case'
        params = {**self.get_default_base_tariffs_params(), 'insurance_type': 'cumulative insurance',
                  'maximum_insurance_period': 69}
        expected_tariffs_file_path = expected_tariffs_folder_path / 'tariffs.xlsx'
        self.check_tariffs_result(params, expected_tariffs_file_path)

        # test creating tariffs table for russian language
        params = {**self.get_default_base_tariffs_params(), 'response_language_code': 'ru'}
        # dictionary that matches cell of tariffs table with expected value
        expected_results_dict = {
            'A1': 'Таблица тарифов в % с технической процентной ставкой 5.00% и нагрузкой 20.00%',
            'A2': 'Тип страхования: страхование жизни на срок',
            'A3': 'Периодичность уплаты страхового взноса: ежегодно',
            'A4': 'Пол застрахованного: мужской',
            'A5': 'Возраст застрахованного',
            'B5': 'Период страхования (лет)'
        }
        self.check_translation_tariffs_table(params, expected_results_dict)

        params = {**self.get_default_base_tariffs_params(), 'insurance_type': 'cumulative insurance',
                  'insurance_premium_frequency': 'simultaneously', 'response_language_code': 'ru'}
        expected_results_dict = {
            'A2': 'Тип страхования: чисто накопительное страхование',
            'A3': 'Периодичность уплаты страхового взноса: единовременно',
            'A4': 'Период страхования (лет, месяцев)',
            'A5': 'Год',
            'B5': 'Месяц'
        }
        self.check_translation_tariffs_table(params, expected_results_dict)

        params = {**self.get_default_base_tariffs_params(),
                  'insurance_type': 'whole life insurance',
                  'insurance_premium_frequency': 'monthly',
                  'gender': 'female', 'response_language_code': 'ru'}
        expected_results_dict = {
            'A2': 'Тип страхования: пожизненное страхование',
            'A3': 'Периодичность уплаты страхового взноса: ежемесячно',
            'A4': 'Пол застрахованного: женский',
            'B5': 'Тариф'
        }
        self.check_translation_tariffs_table(params, expected_results_dict)

        params = {
            **self.get_default_base_tariffs_params(),
            'insurance_type': 'pure endowment',
            'response_language_code': 'ru'
        }
        expected_results_dict = {'A2': 'Тип страхования: чистое дожитие'}
        self.check_translation_tariffs_table(params, expected_results_dict)
