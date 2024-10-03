from django.db import models


class LifeTable(models.Model):
    age = models.IntegerField(primary_key=True)
    men_survived_to_age = models.IntegerField()
    men_died_at_age = models.IntegerField()
    women_survived_to_age = models.IntegerField()
    women_died_at_age = models.IntegerField()
