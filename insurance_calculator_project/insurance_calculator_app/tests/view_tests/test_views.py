from io import BytesIO
from pathlib import Path

from django.urls import reverse
from rest_framework.test import APITestCase
from rest_framework import status

from insurance_calculator_app.tests.test_utils import get_default_request_data, \
    compare_excel_files, get_default_tariffs_request_data


class CalculatorAPITest(APITestCase):

    def check_valid_request_processing(self, url, params, expected_result):
        with self.subTest(**params):
            response = self.client.post(url, params)
            self.assertEqual(response.status_code, status.HTTP_200_OK)
            self.assertEqual(response.data, {'result': expected_result})

    def check_invalid_request_processing(self, url, params, expected_errors):
        with self.subTest(**params):
            response = self.client.post(url, params)
            self.assertEqual(response.status_code, status.HTTP_400_BAD_REQUEST)
            self.assertEqual(response.data, {'errors': expected_errors})

    def test_insurance_premium_route(self):
        url = reverse('insurance-premium')
        base_params = get_default_request_data(False)
        params = {**base_params, 'insuranceSum': 10000}
        expected_result = 22.1781
        self.check_valid_request_processing(url, params, expected_result)
        params['insuranceSum'] = 0
        expected_errors = {
            'non_field_errors': ['Insurance sum must be greater than 0.']}
        self.check_invalid_request_processing(url, params, expected_errors)

    def test_insurance_sum_route(self):
        url = reverse('insurance-sum')
        base_params = get_default_request_data(False)
        params = {**base_params, 'insurancePremium': 100}
        expected_result = 45089.5243
        self.check_valid_request_processing(url, params, expected_result)
        params['insurancePremium'] = 0
        expected_errors = {
            'non_field_errors': ['Insurance premium must be greater than 0.']}
        self.check_invalid_request_processing(url, params, expected_errors)

    def test_reserve_route(self):
        url = reverse('reserve')
        base_params = get_default_request_data(False)
        params = {**base_params, 'insurancePremium': 100, 'reserveCalculationPeriod': 38}
        expected_result = 80.47445
        self.check_valid_request_processing(url, params, expected_result)
        params['reserveCalculationPeriod'] = 0
        expected_errors = {
            'non_field_errors': [
                'Time from start of insurance to insurance reserve calculation must be greater than 0.']}
        self.check_invalid_request_processing(url, params, expected_errors)

    def test_tariffs_route(self):
        url = reverse('tariffs')
        params = get_default_tariffs_request_data(False)
        expected_tariffs_folder_path = Path(__file__).parents[0] / 'expected_tariffs_files'
        with self.subTest(**params):
            response = self.client.post(url, params)
            self.assertEqual(response.status_code, status.HTTP_200_OK)
            expected_tariffs_file_path = expected_tariffs_folder_path / 'tariffs.xlsx'
            self.assertTrue(compare_excel_files(BytesIO(response.content), expected_tariffs_file_path))

        params['maximumInsurancePeriod'] = 0
        expected_errors = {
            'non_field_errors': [
                'Maximum insurance period must be greater than 0.']}
        self.check_invalid_request_processing(url, params, expected_errors)
