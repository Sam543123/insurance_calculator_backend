from django.urls import path

from insurance_calculator_app.views import calculate_insurance_sum, calculate_insurance_premium, calculate_tariffs, \
    calculate_reserve

urlpatterns = [
    path('insurance_premium/', calculate_insurance_premium, name='insurance-premium'),
    path('insurance_sum/', calculate_insurance_sum, name='insurance-sum'),
    path('tariffs/', calculate_tariffs, name='tariffs'),
    path('reserve/', calculate_reserve, name='reserve'),
]