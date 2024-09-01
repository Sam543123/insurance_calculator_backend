import os

from django.http import HttpResponse
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from insurance_calculator_app.insurance_calculator import InsuranceCalculator
from insurance_calculator_app.serializers import SumCalculatorInputSerializer, PremiumCalculatorInputSerializer, \
    ReserveCalculatorInputSerializer, TariffCalculatorInputSerializer


@api_view(['POST'])
def calculate_insurance_premium(request):
    serializer = PremiumCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator()
        result = calculator.calculate_premium(**serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_insurance_sum(request):
    serializer = SumCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator()
        result = calculator.calculate_insurance_sum(**serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_reserve(request):
    serializer = ReserveCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator()
        result = calculator.calculate_reserve(**serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_tariffs(request):
    serializer = TariffCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator()
        tariffs_table_file = calculator.calculate_tariffs(**serializer.validated_data)
        response = HttpResponse(tariffs_table_file,
                                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8')
        return response
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
