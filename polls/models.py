from django.db import models
# ユーザー認証
from django.contrib.auth.models import User

# Create your models here.

class Polls(models.Model):
    syain_cd = models.IntegerField(
        verbose_name='',
        blank=True,
        null=True,
        default=0,
        # validators=[validators.MinValueValidator(0),
        #            validators.MaxValueValidator(100)]
    )
    syozoku = models.TextField(
        blank=True,
    )
    user_name = models.CharField(
        max_length=150,
    )
    password = models.CharField(
        max_length=150,
    )
    authority = models.CharField(
        max_length=150,
    )
    created_datetime = models.DateTimeField(
        auto_now_add=True
    )
    updated_datetime = models.DateTimeField(
        auto_now=True
    )
    
    def __str__(self):
        return self.syain_cd