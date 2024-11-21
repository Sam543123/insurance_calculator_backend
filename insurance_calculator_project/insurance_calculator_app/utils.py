import re
from typing import Any

from drf_spectacular.utils import OpenApiResponse, OpenApiExample


def to_snake_case(string):
    changed_string = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', string)
    result = re.sub('([a-z0-9])([A-Z])', r'\1_\2', changed_string).lower()
    return result


def format_number(n, precision=5):
    return round(n, precision)


def data_fields_to_snake_case(data):
    new_data = {to_snake_case(k): v for k, v in data.items()}
    return new_data


def get_default_request_base_data(snake_case=True):
    data = {'insuranceType': 'term life insurance',
            'insurancePremiumFrequency': 'annually',
            'technicalInterestRate': 0.05,
            'insuranceLoading': 0.2, 'gender': 'male'}
    if snake_case:
        return data_fields_to_snake_case(data)
    return data


def get_default_request_data(snake_case=True):
    base_common_data = get_default_request_base_data()
    data = {**base_common_data, 'birthDate': "1995-10-21",
            'insuranceStartDate': "2024-08-24",
            'insurancePeriod': 69}
    if snake_case:
        return data_fields_to_snake_case(data)
    return data


def get_default_tariffs_request_data(snake_case=True):
    base_common_data = get_default_request_base_data()
    data = {**base_common_data,
            'minimumInsuranceStartAge': 19,
            'maximumInsuranceStartAge': 28,
            'maximumInsurancePeriod': 72}
    if snake_case:
        return data_fields_to_snake_case(data)
    return data


def get_default_errors():
    errors = {
        "errors": {
            "insurance_premium_frequency": [
                "This field is required."
            ],
            "technical_interest_rate": [
                "This field is required."
            ]
        }
    }
    return errors


def get_response_documentation(successful_result):
    responses = {
                   200: OpenApiResponse(response=Any, examples=[
                       OpenApiExample(
                           'Successful response example',
                           value={"result": successful_result},
                           status_codes=[200]
                       )
                   ]),
                   400: OpenApiResponse(response=Any, examples=[
                       OpenApiExample(
                           'Error response example',
                           value=get_default_errors(),
                           status_codes=[400]
                       )
                   ])
               }
    return responses
