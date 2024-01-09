from django.contrib import admin
from .models import Polls   # 追加

# Register your models here.
class PollsAdmin(admin.ModelAdmin):
    list_display = ('id', 'syain_cd', 'created_datetime', 'updated_datetime')
    list_display_links = ('id','syain_cd')
    
admin.site.register(Polls, PollsAdmin)