from io import BytesIO
from math import ceil, log

import openpyxl
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Alignment, NamedStyle, Font, Side, Border


class InsuranceCalculator:

    def __init__(self, life_table_path):
        self.dx_m = []
        self.dx_f = []
        self.lx_m = []
        self.lx_f = []
        self.qx_m = []
        self.qx_f = []

        life_table = openpyxl.load_workbook(life_table_path)

        work_sheet = life_table.active
        for i in range(4, 106):
            for j in range(1, 8):
                cell = work_sheet.cell(row=i, column=j)
                if j == 2:
                    self.lx_m.append(cell.value)
                elif j == 3:
                    self.dx_m.append(cell.value)
                elif j == 4:
                    self.qx_m.append(cell.value)
                elif j == 5:
                    self.lx_f.append(cell.value)
                elif j == 6:
                    self.dx_f.append(cell.value)
                elif j == 7:
                    self.qx_f.append(cell.value)

    def create_tarifs_table(self, tarifs_matrix,
                            type_of_insuranse, time_of_payment, gender, first_age, income, f):
        a = len(tarifs_matrix)
        b = len(tarifs_matrix[0])
        tarifs_table = openpyxl.Workbook()
        tarifs_table.worksheets[0].title = 'Тарифы'
        work_sheet = tarifs_table['Тарифы']

        ns = NamedStyle(name='tableText')
        ns.font = Font(name='Times New Roman', size=12)
        ns.alignment = Alignment(horizontal='center', vertical='center')
        border = Side(style='thin', color='000000')
        ns.border = Border(left=border, top=border, right=border, bottom=border)

        ns2 = NamedStyle(name='hText')
        ns2.font = Font(name='Times New Roman', size=14)
        ns2.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        tarifs_table.add_named_style(ns)
        tarifs_table.add_named_style(ns2)

        for i in range(1, 7):
            for j in range(1, b + 2):
                cell = work_sheet.cell(row=i, column=j)
                if i < 5:
                    cell.style = 'hText'
                else:
                    cell.style = 'tableText'

        work_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=b + 1)
        work_sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=b + 1)
        work_sheet.merge_cells(start_row=3, start_column=1, end_row=3, end_column=b + 1)
        work_sheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=b + 1)
        work_sheet.merge_cells(start_column=1, start_row=5, end_column=1, end_row=6)

        hfont = Font(name='Times New Roman', size=14, bold=True)
        h2font = Font(name='Times New Roman', size=14)

        work_sheet['A1'].font = hfont
        work_sheet['A1'] = 'Таблица величин тарифов в %% со ставкой доходности %0.2f %%  и нагрузкой %0.2f %%' % (
            100 * income, 100 * f)
        work_sheet['A2'] = 'Тип страхования: %s' % type_of_insuranse
        work_sheet['A3'] = 'Периодичность уплаты взносов: %s' % time_of_payment

        if type_of_insuranse != 'чисто накопительное страхование':
            work_sheet.column_dimensions['A'].width = 21
            if b == 1:
                work_sheet.column_dimensions['B'].width = 20
            work_sheet.row_dimensions[1].height = 50
            work_sheet.row_dimensions[2].height = 40
            work_sheet.row_dimensions[3].height = 40
            work_sheet.row_dimensions[4].height = 20
            work_sheet.row_dimensions[5].height = 50
            if type_of_insuranse != 'пожизненное страхование':
                work_sheet.row_dimensions[6].height = 20
            work_sheet['A5'].alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            work_sheet['B5'].alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            work_sheet['B5'].font = h2font
            work_sheet['A4'] = 'Пол застрахованного: %s' % gender
            work_sheet['A5'].font = h2font
            work_sheet['A5'] = 'Возраст застрахованного'
            if type_of_insuranse != 'пожизненное страхование':
                work_sheet.merge_cells(start_row=5, start_column=2, end_row=5, end_column=b + 1)
                work_sheet['B5'] = 'Период страхования (лет)'
                for j in range(1, b + 1):
                    cell = work_sheet.cell(row=6, column=j + 1)
                    cell.style = 'tableText'
                    cell.value = j
            else:
                work_sheet.row_dimensions[5].height = 30
                work_sheet.merge_cells(start_column=2, start_row=5, end_column=2, end_row=6)
                work_sheet['B5'] = 'Величина тарифа'

            for i in range(1, a + 1):
                cell = work_sheet.cell(row=i + 6, column=1)
                cell.style = 'tableText'
                cell.value = first_age + (i - 1)
                for j in range(2, b + 2):
                    cell = work_sheet.cell(row=i + 6, column=j)
                    cell.style = 'tableText'
                    cell.value = tarifs_matrix[i - 1][j - 2]

        else:
            work_sheet.merge_cells(start_row=5, start_column=2, end_row=5, end_column=b + 1)
            work_sheet['A4'] = 'Период страхования (лет, месяцев)'
            work_sheet['A5'].font = h2font
            work_sheet['A5'] = 'Год'
            work_sheet['B5'].font = h2font
            work_sheet['B5'] = 'Месяц'

            for j in range(2, b + 2):
                cell = work_sheet.cell(row=6, column=j)
                cell.style = 'tableText'
                cell.value = j - 2

            for i in range(1, a + 1):
                cell = work_sheet.cell(row=i + 6, column=1)
                cell.style = 'tableText'
                cell.value = i - 1
                for j in range(2, b + 2):
                    cell = work_sheet.cell(row=i + 6, column=j)
                    cell.style = 'tableText'
                    cell.value = tarifs_matrix[i - 1][j - 2]

        # Saving Excel file in memory
        output = BytesIO()
        tarifs_table.save(output)
        output.seek(0)
        return output

    def calculate(self, params):
        excluded_fields = {'birth_date', 'insurance_start_date'}
        new_params = {k: v for k, v in params.items() if k not in excluded_fields}

        delta = relativedelta(params['insurance_start_date'], params['birth_date'])
        start_age = delta.months + delta.years * 12
        end_age = start_age + params['insurance_period']

        new_params['insurance_start_age'] = start_age
        new_params['insurance_end_age'] = end_age
        return self.__calculate(new_params)

    def __calculate(self, params):
        insurance_type = params['insurance_type']
        insurance_premium_frequency = params['insurance_premium_frequency']
        gender = params['gender']
        # insurance_premium_rate
        i = params['insurance_premium_rate']
        # insurance_loading
        f = params['insurance_loading']
        # TODO improve
        # "or -1" is hotfix for calculating reserves
        insurance_premium = params.get('insurance_premium', -1) or -1
        insurance_sum = params.get('insurance_sum', -1) or -1
        insurance_period = params['insurance_period']

        start_age = params['insurance_start_age']
        end_age = params['insurance_end_age']

        result = -1
        lx = []
        dx = []
        if gender == "мужской":
            lx = self.lx_m
            dx = self.dx_m
            qx = self.qx_m
        elif gender == "женский":
            lx = self.lx_f
            dx = self.dx_f
            qx = self.qx_f

        l_start = 0
        l_end = 0
        if insurance_type != 'чисто накопительное страхование':
            l_start = lx[start_age // 12] - dx[start_age // 12] * \
                      (start_age % 12) / 12
            l_end = lx[end_age // 12] - dx[end_age // 12] * \
                    (end_age % 12) / 12
        n = insurance_period // 12
        m = insurance_period % 12
        if m + n == 0:
            return -1
        if insurance_type == "чистое дожитие":
            if insurance_premium_frequency == "единовременно":
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / l_start
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / l_start) ** (-1)
                    result *= insurance_premium
            elif insurance_premium_frequency == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    l_fraq = lx[start_age // 12 + year] - dx[start_age // 12 + year] * \
                             (start_age % 12) / 12
                    v = (1 + i) ** (-year)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= insurance_premium
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    l_fraq = lx[(start_age + month_num) // 12] - dx[(start_age + month_num) // 12] * \
                             ((start_age + month_num) % 12) / 12
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= insurance_premium
        elif insurance_type == "страхование жизни на срок" or insurance_type == "пожизненное страхование":
            sum_compensations = 0
            l1 = l2 = l_start
            for year in range(1, n + 1):
                l2 = lx[start_age // 12 + year] - dx[start_age // 12 + year] * \
                     (start_age % 12) / 12
                d_fraq = l1 - l2
                sum_compensations += (1 + i) ** (-year) * d_fraq
                l1 = l2
            if i != 0:
                sum_compensations *= i / log(1 + i)
            if m != 0:
                l2 = lx[start_age // 12 + n + 1] - dx[start_age // 12 + n + 1] * \
                     (start_age % 12) / 12
                d_end = l1 - l2
                if i != 0:
                    sum_compensations += (1 + i) ** (-(n + 1)) * d_end \
                                         * ((1 + i) - (1 + i) ** (1 - m / 12)) / log(1 + i)
                else:
                    sum_compensations += (1 + i) ** (-(n + 1)) * d_end * m / 12
            if insurance_premium_frequency == "единовременно":
                if insurance_premium == -1:
                    result = sum_compensations / l_start
                    result *= insurance_sum
                else:
                    result = (sum_compensations / l_start) ** (-1)
                    result *= insurance_premium
            elif insurance_premium_frequency == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    l_fraq = lx[start_age // 12 + year] - dx[start_age // 12 + year] * \
                             (start_age % 12) / 12
                    v = (1 + i) ** (-year)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = sum_compensations / annuitet_cost
                    result *= insurance_sum
                else:
                    result = (sum_compensations / annuitet_cost) ** (-1)
                    result *= insurance_premium
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    l_fraq = lx[(start_age + month_num) // 12] - dx[(start_age + month_num) // 12] * \
                             ((start_age + month_num) % 12) / 12
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = sum_compensations / annuitet_cost
                    result *= insurance_sum
                else:
                    result = (sum_compensations / annuitet_cost) ** (-1)
                    result *= insurance_premium
        elif insurance_type == "чисто накопительное страхование":
            if insurance_premium_frequency == "единовременно":
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12))
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12))) ** (-1)
                    result *= insurance_premium
            elif insurance_premium_frequency == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    v = (1 + i) ** (-year)
                    annuitet_cost += v
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12)) / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= insurance_premium
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += v
                if insurance_premium == -1:
                    result = (1 + i) ** (-(n + m / 12)) / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + i) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= insurance_premium

        if insurance_sum != -1:
            result /= (1 - f)
        elif insurance_premium != -1:
            result *= (1 - f)
        return result

    def calculate_reserve(self, params):

        insurance_type = params['insurance_type']
        insurance_premium_frequancy = params['insurance_premium_frequency']
        birth_date = params['birth_date']
        insurance_start_date = params['insurance_start_date']
        gender = params['gender']
        # insurance_premium_rate
        i = params['insurance_premium_rate']
        # insurance_loading
        f = params['insurance_loading']
        # insurance_premium
        A = params['insurance_premium']
        # insurance_sum
        B = params['insurance_sum']
        if A is None:
            A = self.calculate(params)
        #     if is None
        else:
            B = self.calculate(params)
        insurance_period = params['insurance_period']
        reserve_period = params['reserve_calculation_period']

        # TODO check it later
        delta = relativedelta(insurance_start_date, birth_date)
        age_start = delta.months + delta.years * 12
        age_end = age_start + insurance_period

        V = -1
        lx = []
        dx = []
        qx = []
        if gender == "мужской":
            lx = self.lx_m
            dx = self.dx_m
            qx = self.qx_m
        elif gender == "женский":
            lx = self.lx_f
            dx = self.dx_f
            qx = self.qx_f

        n = insurance_period // 12
        m = insurance_period % 12

        p = reserve_period // 12
        q = reserve_period % 12

        l_start = l_end = l_t = 0
        if insurance_type != 'чисто накопительное страхование':
            l_start = lx[age_start // 12] - dx[age_start // 12] * \
                      (age_start % 12) / 12
            l_end = lx[age_end // 12] - dx[age_end // 12] * \
                    (age_end % 12) / 12
            l_t = lx[(age_start + reserve_period) // 12] - dx[(age_start + reserve_period) // 12] * \
                  ((age_start + reserve_period) % 12) / 12

        if insurance_type == "чистое дожитие":
            if insurance_premium_frequancy == "единовременно":
                V = A * (1 - f) * l_start * (1 + i) ** (reserve_period / 12) / l_t
            elif insurance_premium_frequancy == "ежегодно":
                annuitet_cost = 0
                for year in range(p + ceil(q / 12)):
                    l_fraq = lx[age_start // 12 + year] - dx[age_start // 12 + year] * \
                             (age_start % 12) / 12
                    v = (1 + i) ** (-year)
                    annuitet_cost += l_fraq * v
                V = A * (1 - f) * annuitet_cost * (1 + i) ** (reserve_period / 12) / l_t
            else:
                annuitet_cost = 0
                for month_num in range(reserve_period):
                    l_fraq = lx[(age_start + month_num) // 12] - dx[(age_start + month_num) // 12] * \
                             ((age_start + month_num) % 12) / 12
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                V = A * (1 - f) * annuitet_cost * (1 + i) ** (reserve_period / 12) / l_t
        elif insurance_type == "чисто накопительное страхование":
            if insurance_premium_frequancy == "единовременно":
                V = A * (1 - f) * (1 + i) ** (reserve_period / 12)
            elif insurance_premium_frequancy == "ежегодно":
                annuitet_cost = 0
                for year in range(p + ceil(q / 12)):
                    v = (1 + i) ** (-year)
                    annuitet_cost += v
                V = A * (1 - f) * annuitet_cost * (1 + i) ** (reserve_period / 12)
            else:
                annuitet_cost = 0
                for month_num in range(reserve_period):
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += v
                V = A * (1 - f) * annuitet_cost * (1 + i) ** (reserve_period / 12)
        elif insurance_type in ("страхование жизни на срок", "пожизненное страхование"):
            sum_compensations = 0
            l1 = l2 = l_t
            for year in range(p + 1, p + 1 + (insurance_period - reserve_period) // 12):
                l2 = lx[(age_start + q) // 12 + year] - dx[(age_start + q) // 12 + year] * \
                     ((age_start + q) % 12) / 12
                d_fraq = l1 - l2
                sum_compensations += (1 + i) ** (p - year) * d_fraq
                l1 = l2
            if i != 0:
                sum_compensations *= i / log(1 + i)
            if (insurance_period - reserve_period) % 12 != 0:
                l2 = lx[(age_start + q) // 12 + p + 1 + (insurance_period - reserve_period) // 12] - dx[
                    (age_start + q) // 12 + p + 1 + (insurance_period - reserve_period) // 12] * \
                     ((age_start + q) % 12) / 12
                d_end = l1 - l2
                if i != 0:
                    sum_compensations += (1 + i) ** (-(insurance_period - reserve_period) // 12 - 1) * d_end \
                                         * ((1 + i) - (1 + i) ** (
                                1 - ((insurance_period - reserve_period) % 12) / 12)) / log(
                        1 + i)
                else:
                    sum_compensations += (1 + i) ** (-(insurance_period - reserve_period) // 12 - 1) * d_end * (
                            (insurance_period - reserve_period) % 12) / 12
            if insurance_premium_frequancy == "единовременно":
                V = B * sum_compensations / l_t
            elif insurance_premium_frequancy == "ежегодно":
                annuitet_cost = 0
                for year in range(p + ceil(q / 12), n + ceil(m / 12)):
                    l_fraq = lx[age_start // 12 + year] - dx[age_start // 12 + year] * \
                             (age_start % 12) / 12
                    v = (1 + i) ** (p + ceil(q / 12) - year)
                    annuitet_cost += l_fraq * v
                V = (B * sum_compensations - A * (1 - f) * annuitet_cost) / l_t
            else:
                annuitet_cost = 0
                for month_num in range(reserve_period, insurance_period):
                    l_fraq = lx[(age_start + month_num) // 12] - dx[(age_start + month_num) // 12] * \
                             ((age_start + month_num) % 12) / 12
                    v = (1 + i) ** (reserve_period / 12 - month_num / 12)
                    annuitet_cost += l_fraq * v
                V = (B * sum_compensations - A * (1 - f) * annuitet_cost) / l_t
        return V

    def calculate_tariffs(self, params):

        insurance_type = params['insurance_type']
        insurance_premium_frequency = params['insurance_premium_frequency']
        gender = params['gender']
        # insurance_premium_rate
        i = params['insurance_premium_rate']
        # insurance_loading
        f = params['insurance_loading']
        start_age = params['insurance_start_age']
        end_age = params['insurance_end_age']
        # maximum insurance period in years
        # TODO change calculate_tariffs function for life insurance
        if insurance_type == 'пожизненное страхование':
            maximum_period_years = 1
        else:
            maximum_period = params['maximum_insurance_period']
            maximum_period_years = maximum_period // 12

        base_new_params_dict = {
            'insurance_type': insurance_type,
            'insurance_premium_frequency': insurance_premium_frequency,
            'gender': gender,
            'insurance_premium_rate': i,
            'insurance_loading': f
        }

        tarifs_table = []
        if insurance_type != 'чисто накопительное страхование':
            for x in range(start_age, end_age + 1):
                tarifs_table.append([])
                for n in range(1, maximum_period_years + 1):
                    if x + n <= 101:
                        if insurance_type != 'пожизненное страхование':
                            new_params_dict = {
                                'insurance_start_age': 12 * x,
                                'insurance_end_age': 12 * x + 12 * n,
                                'insurance_period': 12 * n,
                                'insurance_premium': -1,
                                'insurance_sum': 100,
                                **base_new_params_dict
                            }
                            tarifs_table[-1].append(self.__calculate(new_params_dict))
                        else:
                            new_params_dict = {
                                'insurance_start_age': 12 * x,
                                'insurance_end_age': 1212,
                                'insurance_period': 1212 - 12 * x,
                                'insurance_premium': -1,
                                'insurance_sum': 100,
                                **base_new_params_dict
                            }
                            tarifs_table[-1].append(self.__calculate(new_params_dict))
                    else:
                        tarifs_table[-1].append('-')

        else:
            for n in range(maximum_period // 12 + 1):
                tarifs_table.append([])
                for m in range(12):
                    new_params_dict = {
                        #  ?
                        'insurance_start_age': -1,
                        # ?
                        'insurance_end_age': -1,
                        'insurance_period': 12 * n + m,
                        'insurance_premium': -1,
                        'insurance_sum': 100,
                        **base_new_params_dict
                    }
                    if 12 * n + m <= maximum_period:
                        tarifs_table[-1].append(self.__calculate(new_params_dict))
                    else:
                        tarifs_table[-1].append('-')
            if tarifs_table[0][0] == -1:
                tarifs_table[0][0] = 100

        # TODO create excel file in memory and send to frontend
        file = self.create_tarifs_table(tarifs_table,
                                 insurance_type, insurance_premium_frequency, gender, start_age, i, f)
        return file
