import os

from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from insurance_calculator_app.insurance_calculator import InsuranceCalculator
from insurance_calculator_app.serializers import CalculatorInputSerializer


# Create your views here.

@api_view(['POST'])
def calculate_insurance_sum(request):
    serializer = CalculatorInputSerializer(data=request.data)
    if serializer.is_valid():
        calculator = InsuranceCalculator(os.path.join(os.path.dirname(__file__), 'data_files', 'life_table.xlsx'))
        result = calculator.calculate(serializer.validated_data)
        return Response({'result': result})
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
