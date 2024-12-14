from typing import Any

from django.http import HttpResponse
from drf_spectacular.types import OpenApiTypes
from drf_spectacular.utils import extend_schema, OpenApiResponse, OpenApiExample
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from insurance_calculator_app.insurance_calculator import InsuranceCalculator
from insurance_calculator_app.serializers import SumCalculatorInputSerializer, PremiumCalculatorInputSerializer, \
    ReserveCalculatorInputSerializer, TariffCalculatorInputSerializer
from insurance_calculator_app.utils import get_default_request_data, get_default_errors, get_response_documentation, \
    get_default_tariffs_request_data


@extend_schema(request=PremiumCalculatorInputSerializer,
               responses=get_response_documentation(221.781),
               examples=[
                   OpenApiExample(
                       'Request Example',
                       value={
                           **get_default_request_data(),
                           'insurance_sum': 10000
                       },
                       request_only=True,
                   )]
               )
@api_view(['POST'])
def calculate_insurance_premium(request):
    serializer = PremiumCalculatorInputSerializer(data=request.data)
    # validate calculator input
    if serializer.is_valid():
        result = InsuranceCalculator.calculate_premium(**serializer.validated_data)
        return Response({'result': result})
    return Response({'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@extend_schema(request=SumCalculatorInputSerializer,
               responses=get_response_documentation(4508.95243),
               examples=[
                   OpenApiExample(
                       'Request Example',
                       value={
                           **get_default_request_data(),
                           'insurance_premium': 10
                       },
                       request_only=True,
                   )]
               )
@api_view(['POST'])
def calculate_insurance_sum(request):
    serializer = SumCalculatorInputSerializer(data=request.data)
    # validate calculator input
    if serializer.is_valid():
        result = InsuranceCalculator.calculate_insurance_sum(**serializer.validated_data)
        return Response({'result': result})
    return Response({'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@extend_schema(request=ReserveCalculatorInputSerializer,
               responses=get_response_documentation(8.04744),
               examples=[
                   OpenApiExample(
                       'Example of request with insurance premium',
                       value={
                           **get_default_request_data(),
                           'insurance_premium': 10,
                           'reserve_calculation_period': 38
                       },
                       request_only=True,
                   ),
                   OpenApiExample(
                       'Example of request with insurance sum',
                       value={
                           **get_default_request_data(),
                           'insurance_loading': None,
                           'insurance_sum': 4508.95243,
                           'reserve_calculation_period': 38
                       },
                       request_only=True,
                   )
               ]
               )
@api_view(['POST'])
def calculate_reserve(request):
    serializer = ReserveCalculatorInputSerializer(data=request.data)
    # validate calculator input
    if serializer.is_valid():
        result = InsuranceCalculator.calculate_reserve(**serializer.validated_data)
        return Response({'result': result})
    return Response({'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@extend_schema(request=TariffCalculatorInputSerializer,
               responses={
                   200: OpenApiResponse(response=OpenApiTypes.BINARY,
                                        examples=[
                                            OpenApiExample(
                                                'Successful response start part example',
                                                value=r"b'PK\x03\x04\x14\x00\x00\x00\x08\x00\xc7\x96pYF\xc7MH'",
                                                status_codes=[200]
                                            )
                                        ]
                                        ),
                   400: OpenApiResponse(response=Any, examples=[
                       OpenApiExample(
                           'Error response example',
                           value=get_default_errors(),
                           status_codes=[400]
                       )
                   ])
               },
               examples=[
                   OpenApiExample(
                       'Request Example',
                       value=get_default_tariffs_request_data(),
                       request_only=True,
                   )]
               )
@api_view(['POST'])
def calculate_tariffs(request):
    serializer = TariffCalculatorInputSerializer(data=request.data)
    # validate calculator input
    if serializer.is_valid():
        tariffs_table_file = InsuranceCalculator.calculate_tariffs(**serializer.validated_data)
        # return xlsx file with tariffs table
        response = HttpResponse(tariffs_table_file,
                                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8')
        response['Content-Disposition'] = 'attachment; filename="tariffs.xlsx"'
        return response
    return Response({'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)
