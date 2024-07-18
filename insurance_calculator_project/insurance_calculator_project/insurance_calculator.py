from math import ceil, log

import openpyxl


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

    def calculate(self, typ, variant, age_start, age_end, period,
                  gender, i, f, payment, compensation):
        result = -1
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

        l_start = 0
        l_end = 0
        if typ != 'чисто накопительное страхование':
            l_start = lx[age_start // 12] - dx[age_start // 12] * \
                      (age_start % 12) / 12
            l_end = lx[age_end // 12] - dx[age_end // 12] * \
                    (age_end % 12) / 12
        n = period // 12
        m = period % 12
        if m + n == 0:
            return -1
        if typ == "чистое дожитие":
            if variant == "единовременно":
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / l_start
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / l_start) ** (-1)
                    result *= payment
            elif variant == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    l_fraq = lx[age_start // 12 + year] - dx[age_start // 12 + year] * \
                             (age_start % 12) / 12
                    v = (1 + i) ** (-year)
                    annuitet_cost += l_fraq * v
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= payment
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    l_fraq = lx[(age_start + month_num) // 12] - dx[(age_start + month_num) // 12] * \
                             ((age_start + month_num) % 12) / 12
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12)) * l_end / annuitet_cost) ** (-1)
                    result *= payment
        elif typ == "страхование жизни на срок" or typ == "пожизненное страхование":
            sum_compensations = 0
            l1 = l2 = l_start
            for year in range(1, n + 1):
                l2 = lx[age_start // 12 + year] - dx[age_start // 12 + year] * \
                     (age_start % 12) / 12
                d_fraq = l1 - l2
                sum_compensations += (1 + i) ** (-year) * d_fraq
                l1 = l2
            if i != 0:
                sum_compensations *= i / log(1 + i)
            if m != 0:
                l2 = lx[age_start // 12 + n + 1] - dx[age_start // 12 + n + 1] * \
                     (age_start % 12) / 12
                d_end = l1 - l2
                if i != 0:
                    sum_compensations += (1 + i) ** (-(n + 1)) * d_end \
                                         * ((1 + i) - (1 + i) ** (1 - m / 12)) / log(1 + i)
                else:
                    sum_compensations += (1 + i) ** (-(n + 1)) * d_end * m / 12
            if variant == "единовременно":
                if payment == -1:
                    result = sum_compensations / l_start
                    result *= compensation
                else:
                    result = (sum_compensations / l_start) ** (-1)
                    result *= payment
            elif variant == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    l_fraq = lx[age_start // 12 + year] - dx[age_start // 12 + year] * \
                             (age_start % 12) / 12
                    v = (1 + i) ** (-year)
                    annuitet_cost += l_fraq * v
                if payment == -1:
                    result = sum_compensations / annuitet_cost
                    result *= compensation
                else:
                    result = (sum_compensations / annuitet_cost) ** (-1)
                    result *= payment
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    l_fraq = lx[(age_start + month_num) // 12] - dx[(age_start + month_num) // 12] * \
                             ((age_start + month_num) % 12) / 12
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += l_fraq * v
                if payment == -1:
                    result = sum_compensations / annuitet_cost
                    result *= compensation
                else:
                    result = (sum_compensations / annuitet_cost) ** (-1)
                    result *= payment
        elif typ == "чисто накопительное страхование":
            if variant == "единовременно":
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12))
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12))) ** (-1)
                    result *= payment
            elif variant == "ежегодно":
                annuitet_cost = 0
                for year in range(n + ceil(m / 12)):
                    v = (1 + i) ** (-year)
                    annuitet_cost += v
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12)) / annuitet_cost
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= payment
            else:
                r = 12 * n + m
                annuitet_cost = 0
                for month_num in range(r):
                    v = (1 + i) ** (-month_num / 12)
                    annuitet_cost += v
                if payment == -1:
                    result = (1 + i) ** (-(n + m / 12)) / annuitet_cost
                    result *= compensation
                else:
                    result = ((1 + i) ** (-(n + m / 12)) / annuitet_cost) ** (-1)
                    result *= payment

        if compensation != -1:
            result /= (1 - f)
        elif payment != -1:
            result *= (1 - f)
        return result