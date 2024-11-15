@echo off
docker build -f Dockerfile -t calculator_backend_image ../
docker rm -f calculator_backend 2>nul || exit /b 0
docker run -d --name calculator_backend -p 8000:8000 -v db_data:/app/insurance_calculator_project/db/ calculator_backend_image