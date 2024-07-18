from django.shortcuts import render
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from insurance_calculator_project.insurance_calculator import InsuranceCalculator
from insurance_calculator_project.serializers import CalculatorInputSerializer


# Create your views here.

@api_view(['POST'])
def calculate_insurance_sum(request):
    serializer = CalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator('data_files/life_table.xlsx')
        result = calculator.calculate_insurance_sum(**serializer.validated_data)
        return Response(result)
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
