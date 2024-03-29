# Generated by Django 5.0.1 on 2024-01-17 06:25

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='Account',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('syain_no', models.IntegerField(blank=True, null=True, verbose_name='社員番号')),
                ('syain_name', models.CharField(blank=True, max_length=20, null=True, verbose_name='社員名')),
                ('syozoku', models.CharField(blank=True, max_length=10, null=True, verbose_name='所属')),
                ('authority', models.CharField(blank=True, max_length=50, null=True, verbose_name='権限')),
                ('created_datetime', models.DateTimeField(auto_now_add=True, null=True, verbose_name='登録日時')),
                ('regist_name', models.CharField(blank=True, max_length=20, null=True, verbose_name='登録者名')),
                ('updated_datetime', models.DateTimeField(auto_now_add=True, null=True, verbose_name='登録日時')),
                ('update_name', models.CharField(blank=True, max_length=20, null=True, verbose_name='更新者名')),
                ('user', models.OneToOneField(on_delete=django.db.models.deletion.CASCADE, to=settings.AUTH_USER_MODEL)),
            ],
        ),
    ]
