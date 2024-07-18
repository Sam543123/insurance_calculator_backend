from math import ceil, log

import openpyxl
from dateutil.relativedelta import relativedelta


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

    def calculate_insurance_premium(self, typ, variant, age_start, age_end, period,
                  gender, i, f, compensation):
        return self.calculate(typ, variant, age_start, age_end, period,
                         gender, i, f, -1, compensation)

    def calculate_insurance_sum(self, typ, variant, age_start, age_end, period,
                                    gender, i, f, payment):
        return self.calculate(typ, variant, age_start, age_end, period,
                  gender, i, f, payment, -1)

    def calculate(self, params):

        insurance_type = params['insurance_type']
        insurance_premium_frequency = params['insurance_premium_frequency']
        birth_date = params['birth_date']
        insurance_start_date = params['insurance_start_date']
        insurance_end_date = params['insurance_end_date']
        gender = params['gender']
        insurance_premium_rate = params['insurance_premium_rate']
        insurance_premium_supplement = params['insurance_premium_supplement']
        insurance_premium = params['insurance_premium'] if params['insurance_premium'] is not None else -1
        insurance_sum = params['insurance_sum']  if params['insurance_sum'] is not None else -1

        # TODO check it later
        delta = relativedelta(insurance_start_date, birth_date)
        insurance_start_age = delta.months + delta.years * 12
        delta2 = relativedelta(insurance_end_date, insurance_start_date)
        insurance_period = delta2.months + delta2.years * 12 + round(delta2.days / 31)
        insurance_end_age = insurance_start_age + insurance_period

        # delta3 = relativedelta(end_insurance_date, birth_date)
        # age_end_insurance_months = delta3.months + delta3.years * 12

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
            l_start = lx[insurance_start_age // 12] - dx[insurance_start_age // 12] * \
                      (insurance_start_age % 12) / 12
            l_end = lx[insurance_end_age // 12] - dx[insurance_end_age // 12] * \
                    (insurance_end_age % 12) / 12
        n = insurance_period // 12
        m = insurance_period % 12
        if m + n == 0:
            return -1
        if insurance_type == "чистое дожитие":
            if insurance_premium_frequency == "единовременно":
                if insurance_premium == -1:
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / l_start
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / l_start) ** (-1)
                    result *= insurance_premium
            elif insurance_premium_frequency == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    l_fraq = lx[insurance_start_age // 12 + year] - dx[insurance_start_age // 12 + year] * \
                             (insurance_start_age % 12) / 12
                    v = (1 + insurance_premium_rate) ** (-year)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= insurance_premium
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    l_fraq = lx[(insurance_start_age + month_num) // 12] - dx[(insurance_start_age + month_num) // 12] * \
                             ((insurance_start_age + month_num) % 12) / 12
                    v = (1 + insurance_premium_rate) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                if insurance_premium == -1:
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= insurance_premium
        elif insurance_type == "страхование жизни на срок" or insurance_type == "пожизненное страхование":
            sum_compensations = 0
            l1 = l2 = l_start
            for year in range(1, n + 1):
                l2 = lx[insurance_start_age // 12 + year] - dx[insurance_start_age // 12 + year] * \
                     (insurance_start_age % 12) / 12
                d_fraq = l1 - l2
                sum_compensations += (1 + insurance_premium_rate) ** (-year) * d_fraq
                l1 = l2
            if insurance_premium_rate != 0:
                sum_compensations *= insurance_premium_rate / log(1 + insurance_premium_rate)
            if m != 0:
                l2 = lx[insurance_start_age // 12 + n + 1] - dx[insurance_start_age // 12 + n + 1] * \
                     (insurance_start_age % 12) / 12
                d_end = l1 - l2
                if insurance_premium_rate != 0:
                    sum_compensations += (1 + insurance_premium_rate) ** (-(n + 1)) * d_end \
                                         * ((1 + insurance_premium_rate) - (1 + insurance_premium_rate) ** (1 - m / 12)) / log(1 + insurance_premium_rate)
                else:
                    sum_compensations += (1 + insurance_premium_rate) ** (-(n + 1)) * d_end * m / 12
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
                    l_fraq = lx[insurance_start_age // 12 + year] - dx[insurance_start_age // 12 + year] * \
                             (insurance_start_age % 12) / 12
                    v = (1 + insurance_premium_rate) ** (-year)
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
                    l_fraq = lx[(insurance_start_age + month_num) // 12] - dx[(insurance_start_age + month_num) // 12] * \
                             ((insurance_start_age + month_num) % 12) / 12
                    v = (1 + insurance_premium_rate) ** (-month_num / 12)
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
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12))
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12))) ** (-1)
                    result *= insurance_premium
            elif insurance_premium_frequency == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    v = (1 + insurance_premium_rate) ** (-year)
                    annuitet_cost += v
                if insurance_premium == -1:
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12)) / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= insurance_premium
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    v = (1 + insurance_premium_rate) ** (-month_num / 12)
                    annuitet_cost += v
                if insurance_premium == -1:
                    result = (1 + insurance_premium_rate) ** (-(n + m / 12)) / annuitet_cost
                    result *= insurance_sum
                else:
                    result = ((1 + insurance_premium_rate) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= insurance_premium

        if insurance_sum != -1:
            result /= (1 - insurance_premium_supplement)
        elif insurance_premium != -1:
            result *= (1 - insurance_premium_supplement)
        return result