from django.shortcuts import render
from rest_framework.decorators import api_view
from rest_framework.response import Response

from insurance_calculator_project.insurance_calculator import InsuranceCalculator


# Create your views here.

@api_view(['GET'])
def calculate_insurance_sum(request, parameters):
    calculator = InsuranceCalculator('data_files/life_table.xlsx')
    result = calculator.calculate_insurance_sum()
    return Response(result)
