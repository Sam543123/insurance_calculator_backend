FROM python:3.10-slim
WORKDIR /app
COPY requirements.txt .
COPY insurance_calculator_project insurance_calculator_project/
ENV PYTHONPATH "${PYTHONPATH}:/app/insurance_calculator_project"
ENV PYTHONUNBUFFERED=1

RUN pip install --no-cache-dir --upgrade -r requirements.txt
RUN python -m manage makemigrations
RUN python -m manage migrate

CMD ["python", "-m", "manage", "runserver", "0.0.0.0:8000", "--noreload"]