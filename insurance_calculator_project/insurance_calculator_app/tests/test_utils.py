import openpyxl

from insurance_calculator_app.utils import to_snake_case


def data_fields_to_snake_case(data):
    new_data = {to_snake_case(k): v for k, v in data.items()}
    return new_data


def get_default_request_base_data(snake_case=True):
    data = {'insuranceType': 'term life insurance',
            'insurancePremiumFrequency': 'annually',
            'insurancePremiumRate': 0.05,
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


def compare_excel_files(file_like_obj1, file_like_obj2):
    """Check if two excel files are equal"""
    workbook1 = openpyxl.load_workbook(file_like_obj1, data_only=True)
    work_sheet1 = workbook1.active
    workbook2 = openpyxl.load_workbook(file_like_obj2, data_only=True)
    work_sheet2 = workbook2.active
    if work_sheet1.max_row != work_sheet2.max_row or work_sheet1.max_column != work_sheet2.max_column:
        return False
    for row1, row2 in zip(work_sheet1.iter_rows(values_only=True), work_sheet2.iter_rows(values_only=True)):
        if row1 != row2:
            return False
    return True
