from io import BytesIO
from math import ceil, log
from typing import Optional

from dateutil.relativedelta import relativedelta
from openpyxl.styles import Alignment, NamedStyle, Font, Side, Border
import datetime as dt

from openpyxl.workbook import Workbook

from insurance_calculator_app.models import LifeTable
from insurance_calculator_app.utils import format_number


class InsuranceCalculator:

    # Maximum age of insured person in months
    MAXIMUM_INSURANCE_AGE_MONTHS = 1212

    @staticmethod
    def __parse_common_params(insurance_type: str, gender: Optional[str], birth_date: dt.datetime,
                              insurance_start_date: dt.datetime, insurance_period: Optional[int]):
        """
        Parse params for calculating insurance sum, insurance premium and reserve
        """
        parsed_params = {'lx': None, 'dx': None}
        delta = relativedelta(insurance_start_date, birth_date)
        # age of the insured person at the time of insurance start in months
        start_age = delta.months + 12 * delta.years
        processed_insurance_period = insurance_period
        if insurance_type == "whole life insurance":
            # It is supposed that the maximum age of the insured person is 101 years
            end_age = InsuranceCalculator.MAXIMUM_INSURANCE_AGE_MONTHS
            processed_insurance_period = end_age - start_age
        if insurance_type != "cumulative insurance":
            # age of the insured person  at the time of insurance start and end in years
            # to get necessary records from life table
            lower_age_border = start_age // 12
            higher_age_border = (start_age + processed_insurance_period) // 12 + 1
            life_table_values = InsuranceCalculator.__get_life_table_records(gender, lower_age_border,
                                                                             higher_age_border)
            parsed_params.update(life_table_values)
        parsed_params.update({'insurance_period': processed_insurance_period, 'start_age': start_age})
        return parsed_params

    @staticmethod
    def __parse_tariffs_params(insurance_type, gender: Optional[str],
                               maximum_insurance_start_age: Optional[int],
                               minimum_insurance_start_age: Optional[int],
                               maximum_insurance_period: Optional[int]):
        """
        Parse params for calculating tariffs
        """
        parsed_params = {'lx': None, 'dx': None}
        processed_maximum_insurance_period = maximum_insurance_period
        if insurance_type == "whole life insurance":
            # It is supposed that the maximum age of the insured person is 101 years
            end_age = InsuranceCalculator.MAXIMUM_INSURANCE_AGE_MONTHS
            processed_maximum_insurance_period = end_age - maximum_insurance_start_age
        if insurance_type != "cumulative insurance":
            # age of the insured person  at the time of insurance start and end in years
            # to get necessary records from life table
            lower_age_border = maximum_insurance_start_age
            higher_age_border = minimum_insurance_start_age + processed_maximum_insurance_period // 12
            life_table_values = InsuranceCalculator.__get_life_table_records(gender, lower_age_border,
                                                                             higher_age_border)
            parsed_params.update(life_table_values)
        parsed_params['maximum_insurance_period'] = processed_maximum_insurance_period
        return parsed_params

    @staticmethod
    def __get_life_table_records(gender: str, lower_age_border: int, higher_age_border: int):
        """
        Get necessary values from life table
        """

        # number of people survived to particular age
        lx = {}
        # number of people died at particular age
        dx = {}
        age_field = 'age'
        if gender == "male":
            lx_field, dx_field = 'men_survived_to_age', 'men_died_at_age'
        else:
            lx_field, dx_field = 'women_survived_to_age', 'women_died_at_age'
        filter_conditions = {f'{age_field}__gte': lower_age_border, f'{age_field}__lte': higher_age_border}
        life_table_records = LifeTable.objects.filter(**filter_conditions).values(age_field, lx_field, dx_field)
        for r in life_table_records:
            lx[r[age_field]] = r[lx_field]
            dx[r[age_field]] = r[dx_field]
        life_table_values = {'lx': lx, 'dx': dx}
        return life_table_values

    @staticmethod
    def calculate_premium(insurance_type: str, insurance_premium_frequency: str, gender: Optional[str],
                          insurance_premium_rate: float, insurance_loading: float, birth_date: Optional[dt.datetime],
                          insurance_start_date: Optional[dt.datetime], insurance_period: Optional[int],
                          insurance_sum: float):
        """
        Calculate insurance premium using raw parameters
        """
        parsed_params = InsuranceCalculator.__parse_common_params(insurance_type, gender,
                                                                  birth_date,
                                                                  insurance_start_date,
                                                                  insurance_period)
        params = {
            'insurance_type': insurance_type, 'insurance_premium_frequency': insurance_premium_frequency,
            'insurance_premium_rate': insurance_premium_rate, 'insurance_loading': insurance_loading,
            'insurance_sum': insurance_sum, **parsed_params
        }

        result = InsuranceCalculator.__calculate_premium(**params)
        return format_number(result)

    @staticmethod
    def __calculate_premium(insurance_type: str, insurance_premium_frequency: str,
                            insurance_premium_rate: float, insurance_loading: float, insurance_period: int,
                            insurance_sum: float, lx: Optional[dict[int, float]], dx: Optional[dict[int, float]],
                            start_age: Optional[int] = None):
        """
        Calculate insurance premium using parsed parameters
        """
        sum_annuity = InsuranceCalculator.__calculate_insurance_sum_annuity(insurance_type, insurance_premium_rate,
                                                                            insurance_period, lx, dx, start_age)
        premium_annuity = InsuranceCalculator.__calculate_premium_annuity(insurance_type, insurance_premium_frequency,
                                                                          insurance_premium_rate, insurance_period, lx,
                                                                          dx, start_age)
        result = (insurance_sum * sum_annuity) / (premium_annuity * (1 - insurance_loading))
        return result

    @staticmethod
    def calculate_insurance_sum(insurance_type: str, insurance_premium_frequency: str, gender: Optional[str],
                                insurance_premium_rate: float, insurance_loading: float,
                                birth_date: Optional[dt.datetime],
                                insurance_start_date: Optional[dt.datetime], insurance_period: Optional[int],
                                insurance_premium: float):
        """
        Calculate insurance sum using raw parameters
        """
        parsed_params = InsuranceCalculator.__parse_common_params(insurance_type, gender,
                                                                  birth_date,
                                                                  insurance_start_date,
                                                                  insurance_period)
        params = {
            'insurance_type': insurance_type, 'insurance_premium_frequency': insurance_premium_frequency,
            'insurance_premium_rate': insurance_premium_rate, 'insurance_loading': insurance_loading,
            'insurance_premium': insurance_premium, **parsed_params
        }

        result = InsuranceCalculator.__calculate_insurance_sum(**params)
        return format_number(result)

    @staticmethod
    def __calculate_insurance_sum(insurance_type: str, insurance_premium_frequency: str,
                                  insurance_premium_rate: float, insurance_loading: float, insurance_period: int,
                                  insurance_premium: float, lx: Optional[dict[int, float]],
                                  dx: Optional[dict[int, float]],
                                  start_age: Optional[int] = None):
        """
        Calculate insurance sum using parsed parameters
        """
        sum_annuity = InsuranceCalculator.__calculate_insurance_sum_annuity(insurance_type, insurance_premium_rate,
                                                                            insurance_period, lx, dx, start_age)
        premium_annuity = InsuranceCalculator.__calculate_premium_annuity(insurance_type, insurance_premium_frequency,
                                                                          insurance_premium_rate, insurance_period, lx,
                                                                          dx, start_age)
        result = (insurance_premium * premium_annuity * (1 - insurance_loading)) / sum_annuity
        return result

    @staticmethod
    def calculate_reserve(insurance_type: str, insurance_premium_frequency: str, gender: Optional[str],
                          insurance_premium_rate: float, insurance_loading: Optional[float],
                          birth_date: Optional[dt.datetime],  insurance_start_date: Optional[dt.datetime],
                          insurance_period: Optional[int], insurance_sum: Optional[float],
                          insurance_premium: Optional[float], reserve_calculation_period: int):
        """
        Calculate reserve using raw parameters
        """
        parsed_params = InsuranceCalculator.__parse_common_params(insurance_type, gender,
                                                                  birth_date,
                                                                  insurance_start_date,
                                                                  insurance_period)
        params = {
            'insurance_type': insurance_type, 'insurance_premium_frequency': insurance_premium_frequency,
            'insurance_premium_rate': insurance_premium_rate, 'insurance_loading': insurance_loading,
            'insurance_premium': insurance_premium, 'insurance_sum': insurance_sum,
            'reserve_calculation_period': reserve_calculation_period, **parsed_params
        }
        result = InsuranceCalculator.__calculate_reserve(**params)
        return format_number(result)

    @staticmethod
    def __calculate_reserve(insurance_type: str, insurance_premium_frequency: str,
                            insurance_premium_rate: float, insurance_loading: Optional[float],
                            insurance_sum: Optional[float], insurance_premium: Optional[float],
                            insurance_period: int, reserve_calculation_period: int, lx: Optional[dict[int, float]],
                            dx: Optional[dict[int, float]], start_age: int):
        """
        Calculate reserve using parsed parameters
        """
        processed_insurance_loading = insurance_loading
        v = 1 / (1 + insurance_premium_rate)
        if insurance_type != "cumulative insurance":
            # number of people survived to time of reserve calculation
            l_t = InsuranceCalculator.__fractional_lx(start_age + reserve_calculation_period, lx, dx)
        # if insurance_premium is None then insurance_sum is used to calculate reserve
        if insurance_premium is None:
            # if reserve is calculated using insurance sum then
            # insurance_loading doesn't affect reserve calculation, and
            # we can assign to it an arbitrary value, 0 for example
            processed_insurance_loading = 0 if insurance_loading is None else insurance_loading
            insurance_premium = InsuranceCalculator.__calculate_premium(insurance_type, insurance_premium_frequency,
                                                                        insurance_premium_rate,
                                                                        processed_insurance_loading,
                                                                        insurance_period, insurance_sum, lx, dx,
                                                                        start_age)
        else:
            insurance_sum = InsuranceCalculator.__calculate_insurance_sum(insurance_type,
                                                                          insurance_premium_frequency,
                                                                          insurance_premium_rate,
                                                                          processed_insurance_loading,
                                                                          insurance_period, insurance_premium,
                                                                          lx, dx, start_age)

        #  calculate reserve for pure endowment and cumulative insurance as difference between
        #  past payments of insurance premium and insurance sum
        if insurance_type in {"pure endowment", "cumulative insurance"}:
            premium_annuity = InsuranceCalculator.__calculate_premium_annuity(insurance_type,
                                                                              insurance_premium_frequency,
                                                                              insurance_premium_rate,
                                                                              reserve_calculation_period, lx,
                                                                              dx, start_age)
            if insurance_type == "pure endowment":
                reserve = insurance_premium * (1 - processed_insurance_loading) * premium_annuity / (
                            l_t * v ** (reserve_calculation_period / 12))
            else:
                reserve = insurance_premium * (1 - processed_insurance_loading) * premium_annuity / (
                            v ** (reserve_calculation_period / 12))
        # calculate reserve for term life insurance and whole life insurance as difference between
        # future payments of insurance sum and insurance premium
        else:
            premium_annuity = InsuranceCalculator.__calculate_premium_annuity(insurance_type,
                                                                              insurance_premium_frequency,
                                                                              insurance_premium_rate,
                                                                              insurance_period, lx,
                                                                              dx, start_age,
                                                                              skip_period=reserve_calculation_period)
            sum_annuity = InsuranceCalculator.__calculate_insurance_sum_annuity(insurance_type, insurance_premium_rate,
                                                                                insurance_period, lx, dx, start_age,
                                                                                skip_period=reserve_calculation_period)
            reserve = (insurance_sum * sum_annuity - insurance_premium * (
                        1 - processed_insurance_loading) * premium_annuity) / l_t
        return reserve

    @staticmethod
    def __calculate_premium_annuity(insurance_type: str, insurance_premium_frequency: str,
                                    insurance_premium_rate: float, insurance_period: int,
                                    lx: Optional[dict[int, float]], dx: Optional[dict[int, float]],
                                    start_age: Optional[int] = None, skip_period: int = 0):
        """
        Calculate insurance premium annuity cost assuming that cost of one payment is equal to 1
        """
        annuity_cost = 0
        v = 1 / (1 + insurance_premium_rate)
        if insurance_premium_frequency == "simultaneously" and skip_period == 0:
            if insurance_type != "cumulative insurance":
                annuity_cost = InsuranceCalculator.__fractional_lx(start_age, lx, dx)
            else:
                annuity_cost = 1
        elif insurance_premium_frequency == "annually":
            for j, year in enumerate(range(ceil(skip_period / 12), ceil(insurance_period / 12))):
                #  multiplier to discount annuity cost at start time
                discount_multiplier = v ** j
                if insurance_type != "cumulative insurance":
                    l_frac = InsuranceCalculator.__fractional_lx(start_age + 12 * year, lx, dx)
                    annuity_cost += discount_multiplier * l_frac
                else:
                    annuity_cost += discount_multiplier
        elif insurance_premium_frequency == "monthly":
            for j, month in enumerate(range(skip_period, insurance_period)):
                discount_multiplier = v ** (j / 12)
                if insurance_type != "cumulative insurance":
                    l_frac = InsuranceCalculator.__fractional_lx(start_age + month, lx, dx)
                    annuity_cost += discount_multiplier * l_frac
                else:
                    annuity_cost += discount_multiplier

        return annuity_cost

    @staticmethod
    def __calculate_insurance_sum_annuity(insurance_type: str, insurance_premium_rate: float, insurance_period: int,
                                          lx: Optional[dict[int, float]], dx: Optional[dict[int, float]],
                                          start_age: Optional[int] = None, skip_period: int = 0):
        """
        Calculate insurance sum annuity cost assuming that cost of one payment is equal to 1
        """
        annuity_cost = 0
        v = 1 / (1 + insurance_premium_rate)
        # use alias for frequently used parameters
        i = insurance_premium_rate
        if insurance_type == "cumulative insurance":
            annuity_cost = v ** ((insurance_period - skip_period) / 12)
        elif insurance_type == "pure endowment":
            l_end = InsuranceCalculator.__fractional_lx(start_age + insurance_period, lx, dx)
            annuity_cost = v ** ((insurance_period - skip_period) / 12) * l_end
        else:
            for year in range((insurance_period - skip_period) // 12):
                fractional_dx = InsuranceCalculator.__fractional_dx(start_age + skip_period + 12 * year, lx, dx)
                annuity_cost += v ** (year + 1) * fractional_dx
            if i != 0:
                annuity_cost *= i / log(1 + i)
            # Use derived formula to estimate the remaining cost of the annuity
            # if the annuity period is fractional
            if (insurance_period - skip_period) % 12 != 0:
                end_age = start_age + skip_period + 12 * ((insurance_period - skip_period) // 12)
                d_end = InsuranceCalculator.__fractional_dx(end_age, lx, dx)
                if i != 0:
                    annuity_cost += v ** ((insurance_period - skip_period) // 12 + 1) * d_end * (
                            (1 + i) - (1 + i) ** (1 - ((insurance_period - skip_period) % 12) / 12)) / log(
                        1 + i)
                else:
                    annuity_cost += v ** ((insurance_period - skip_period) // 12 + 1) * d_end * (
                            (insurance_period - skip_period) % 12) / 12

        return annuity_cost

    @staticmethod
    def __fractional_lx(age: int, lx: Optional[dict[int, float]], dx: Optional[dict[int, float]]):
        """Evaluate number of people survived to fractional age"""
        return lx[age // 12] - dx[age // 12] * (age % 12) / 12

    @staticmethod
    def __fractional_dx(age: int, lx: Optional[dict[int, float]], dx: Optional[dict[int, float]]):
        """Evaluate number of people died at fractional age"""
        return InsuranceCalculator.__fractional_lx(age, lx, dx) - InsuranceCalculator.__fractional_lx(age + 12, lx, dx)

    @staticmethod
    def calculate_tariffs(insurance_type: str, insurance_premium_frequency: str, gender: Optional[str],
                          insurance_premium_rate: float, insurance_loading: float,
                          minimum_insurance_start_age: Optional[int], maximum_insurance_start_age: Optional[int],
                          maximum_insurance_period: int):
        """
        Create tariffs table using raw parameters
        """

        parsed_params = InsuranceCalculator.__parse_tariffs_params(insurance_type, gender,
                                                                   minimum_insurance_start_age,
                                                                   maximum_insurance_start_age,
                                                                   maximum_insurance_period)

        params = {
            'insurance_type': insurance_type, 'insurance_premium_frequency': insurance_premium_frequency,
            'gender': gender, 'insurance_premium_rate': insurance_premium_rate,
            'insurance_loading': insurance_loading, 'minimum_start_age': minimum_insurance_start_age,
            'maximum_start_age': maximum_insurance_start_age, **parsed_params
        }

        return InsuranceCalculator.__calculate_tariffs(**params)

    @staticmethod
    def __calculate_tariffs(insurance_type: str, insurance_premium_frequency: str, gender: Optional[str],
                            insurance_premium_rate: float, insurance_loading: float,
                            minimum_start_age: Optional[int], maximum_start_age: Optional[int],
                            maximum_insurance_period: Optional[int], lx: Optional[dict[int, float]],
                            dx: Optional[dict[int, float]]):
        """
        Create tariffs table using parsed parameters
        """
        if insurance_type != 'whole life insurance':
            tariffs_table_column_count = maximum_insurance_period // 12
        else:
            tariffs_table_column_count = 1
        # Set tariffs_insurance_sum to 100 instead of 1 to get tariffs values in %
        tariffs_insurance_sum = 100
        tariffs_table = []
        if insurance_type != 'cumulative insurance':
            for x in range(minimum_start_age, maximum_start_age + 1):
                start_age = 12 * x
                tariffs_table.append([])
                for n in range(1, tariffs_table_column_count + 1):
                    if insurance_type == 'whole life insurance':
                        insurance_period = InsuranceCalculator.MAXIMUM_INSURANCE_AGE_MONTHS - 12 * x
                        # Calculate tariff as insurance premium based on insurance sum equal to 1
                        tariff = format_number(InsuranceCalculator.__calculate_premium(insurance_type,
                                                                                       insurance_premium_frequency,
                                                                                       insurance_premium_rate,
                                                                                       insurance_loading,
                                                                                       insurance_period,
                                                                                       tariffs_insurance_sum,
                                                                                       lx, dx,
                                                                                       start_age))
                    elif insurance_type == 'whole life insurance' or x + n <= 101:
                        insurance_period = 12 * n
                        tariff = format_number(InsuranceCalculator.__calculate_premium(insurance_type,
                                                                                       insurance_premium_frequency,
                                                                                       insurance_premium_rate,
                                                                                       insurance_loading,
                                                                                       insurance_period,
                                                                                       tariffs_insurance_sum,
                                                                                       lx, dx,
                                                                                       start_age))
                    else:
                        tariff = '-'
                    tariffs_table[-1].append(tariff)

        else:
            for n in range(maximum_insurance_period // 12 + 1):
                tariffs_table.append([])
                for m in range(12):
                    insurance_period = 12 * n + m
                    if m + n != 0 and 12 * n + m <= maximum_insurance_period:
                        tariff = format_number(InsuranceCalculator.__calculate_premium(insurance_type,
                                                                                       insurance_premium_frequency,
                                                                                       insurance_premium_rate,
                                                                                       insurance_loading,
                                                                                       insurance_period,
                                                                                       tariffs_insurance_sum, lx,
                                                                                       dx))
                    else:
                        tariff = '-'
                    tariffs_table[-1].append(tariff)

        # return tariffs table file
        file = InsuranceCalculator.__create_tariffs_table(insurance_type, insurance_premium_frequency,
                                                          gender, insurance_premium_rate, insurance_loading,
                                                          tariffs_table, minimum_start_age)
        return file

    @staticmethod
    def __create_tariffs_table(insurance_type: str, insurance_premium_frequency: str,
                               gender: Optional[str],  insurance_premium_rate: float,
                               insurance_loading: float, tariffs_table: list[list[float]],
                               start_age: Optional[int]):
        """
        Create xlsx file with tariffs table
        """
        row_number = len(tariffs_table)
        column_number = len(tariffs_table[0])
        tariffs_page = Workbook()
        work_sheet = tariffs_page.active
        work_sheet.title = 'Tariffs'

        border_side = Side(style='thin', color='000000')
        border = Border(left=border_side, top=border_side, right=border_side, bottom=border_side)
        header_font = Font(name='Times New Roman', size=14)
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        table_text_style = NamedStyle(name='tableText', font=Font(name='Times New Roman', size=12),
                                      alignment=Alignment(horizontal='center', vertical='center'),
                                      border=border)
        table_header_style = NamedStyle(name='tableHeader', font=header_font,
                                        alignment=center_alignment,
                                        border=border)

        header_style = NamedStyle(name='header', font=header_font, alignment=center_alignment)

        tariffs_page.add_named_style(header_style)
        tariffs_page.add_named_style(table_header_style)
        tariffs_page.add_named_style(table_text_style)

        # fill table info rows
        for i in range(1, 5):
            work_sheet[f'A{i}'].style = 'header'
            work_sheet.merge_cells(start_row=i, start_column=1, end_row=i, end_column=column_number + 1)

        work_sheet['A5'].style = 'tableHeader'
        work_sheet['B5'].style = 'tableHeader'

        work_sheet['A1'].font = Font(name='Times New Roman', size=14, bold=True)
        work_sheet['A1'] = (f'Tariffs table in % with insurance premium rate {(100 * insurance_premium_rate):.2f}% '
                            f'and loading {(100 * insurance_loading):.2f}%')
        work_sheet['A2'] = f'Insurance type: {insurance_type}'
        work_sheet['A3'] = f'Payment frequency: {insurance_premium_frequency}'

        work_sheet.row_dimensions[1].height = 50
        work_sheet.row_dimensions[2].height = 40
        work_sheet.row_dimensions[3].height = 40
        work_sheet.row_dimensions[4].height = 40

        # fill table headers
        skipped_rows = 6
        if insurance_type != 'cumulative insurance':
            work_sheet.column_dimensions['A'].width = 20
            work_sheet['A4'] = f'Gender of insured person: {gender}'
            work_sheet['A5'] = 'Age of insured person'
            if insurance_type != 'whole life insurance':
                work_sheet.row_dimensions[5].height = 30
                work_sheet.row_dimensions[6].height = 20
                work_sheet.merge_cells(start_row=5, start_column=1, end_row=6, end_column=1)
                work_sheet.merge_cells(start_row=5, start_column=2, end_row=5, end_column=column_number + 1)
                work_sheet['B5'] = 'Insurance period (years)'
                for j in range(1, column_number + 1):
                    cell = work_sheet.cell(row=6, column=j + 1)
                    cell.style = 'tableText'
                    cell.value = j
            else:
                work_sheet.row_dimensions[5].height = 50
                work_sheet.column_dimensions['B'].width = 20
                work_sheet['B5'] = 'Tariff'
                skipped_rows = 5

            start_first_column_value = start_age

        else:
            work_sheet.row_dimensions[5].height = 30
            work_sheet.merge_cells(start_row=5, start_column=1, end_row=6, end_column=1)
            work_sheet.merge_cells(start_row=5, start_column=2, end_row=5, end_column=column_number + 1)
            work_sheet['A4'] = 'Insurance period (years, months)'
            work_sheet['A5'] = 'Year'
            work_sheet['B5'] = 'Month'

            for j in range(2, column_number + 2):
                cell = work_sheet.cell(row=6, column=j)
                cell.style = 'tableText'
                cell.value = j - 2
            start_first_column_value = 0

        # fill table content
        for i in range(1, row_number + 1):
            cell = work_sheet.cell(row=i + skipped_rows, column=1)
            cell.style = 'tableText'
            cell.value = start_first_column_value + (i - 1)
            for j in range(2, column_number + 2):
                cell = work_sheet.cell(row=i + skipped_rows, column=j)
                cell.style = 'tableText'
                cell.value = tariffs_table[i - 1][j - 2]

        # Saving xlsx file in memory
        output = BytesIO()
        tariffs_page.save(output)
        output.seek(0)
        return output
