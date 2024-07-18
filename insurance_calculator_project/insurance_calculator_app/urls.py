from django.urls import path
from rest_framework.urlpatterns import format_suffix_patterns

from insurance_calculator_app.views import calculate_insurance_sum

urlpatterns = [
    path('insurance_sum/', calculate_insurance_sum),
]