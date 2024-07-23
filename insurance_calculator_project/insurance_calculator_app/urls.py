from django.urls import path
from rest_framework.urlpatterns import format_suffix_patterns

from insurance_calculator_app.views import calculate_insurance_sum, calculate_insurance_premium, calculate_tariffs, \
    calculate_reserve

urlpatterns = [
    path('insurance_premium/', calculate_insurance_premium),
    path('insurance_sum/', calculate_insurance_sum),
    path('tariffs/', calculate_tariffs),
    path('reserve/', calculate_reserve),
]