import asyncio
import os
import threading
import time

from django.http import HttpResponse
from rest_framework import status
from rest_framework.decorators import api_view
from adrf.decorators import api_view as async_api_view
from rest_framework.response import Response

from insurance_calculator_app.insurance_calculator import InsuranceCalculator
from insurance_calculator_app.serializers import SumCalculatorInputSerializer, PremiumCalculatorInputSerializer, \
    ReserveCalculatorInputSerializer, TariffCalculatorInputSerializer


@api_view(['POST'])
def calculate_insurance_premium(request):
    serializer = PremiumCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator(os.path.join(os.path.dirname(__file__), 'data_files', 'life_table.xlsx'))
        result = calculator.calculate(serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_insurance_sum(request):
    serializer = SumCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator(os.path.join(os.path.dirname(__file__), 'data_files', 'life_table.xlsx'))
        result = calculator.calculate(serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_reserve(request):
    serializer = ReserveCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator(os.path.join(os.path.dirname(__file__), 'data_files', 'life_table.xlsx'))
        result = calculator.calculate_reserve(serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def calculate_tariffs(request):
    serializer = TariffCalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator(os.path.join(os.path.dirname(__file__), 'data_files', 'life_table.xlsx'))
        file = calculator.calculate_tariffs(serializer.validated_data)
        response = HttpResponse(file, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8')
        return response
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['GET'])
def test_io_bound_route(request):
    current_thread_id = threading.get_ident()
    time.sleep(20)
    return Response({"message": "test_io_bound_route", "thread_id": current_thread_id, "process_id": os.getpid()})


@api_view(['GET'])
def test_cpu_bound_route(request):
    current_thread_id = threading.get_ident()
    start_time = time.time()

    result = 0
    for _ in range(90000000):  # it approximately takes 10 seconds to process one request
        result += 1

    # Проверяем, сколько времени прошло
    elapsed_time = time.time() - start_time


    print(f"Completed in {elapsed_time:.2f} seconds")
    return Response({"message": "test_cpu_bound_route", "thread_id": current_thread_id, "process_id": os.getpid()})


# It is possible to send requests to this route only if app is run on uvicorn or other asgi server
@async_api_view(['GET'])
async def test_async_io_bound_route(request):
    current_thread_id = threading.get_ident()
    await asyncio.sleep(20)
    return Response({"message": "test_async_io_bound_route", "thread_id": current_thread_id, "process_id": os.getpid()})
