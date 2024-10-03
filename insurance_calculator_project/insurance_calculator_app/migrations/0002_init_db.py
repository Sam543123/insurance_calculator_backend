# Generated by Django 5.0.6 on 2024-08-22 11:26
import csv
import os
from pathlib import Path

from django.db import migrations


def load_initial_data(apps, schema_editor):
    life_table_model = apps.get_model('insurance_calculator_app', 'LifeTable')
    life_table_file = Path(__file__).parents[1] / 'data_files' / 'life_table.csv'
    with open(life_table_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        records = [life_table_model(**row) for row in reader]
    life_table_model.objects.bulk_create(records)


class Migration(migrations.Migration):

    dependencies = [
        ('insurance_calculator_app', '0001_initial'),
    ]

    operations = [
        migrations.RunPython(load_initial_data),
    ]
