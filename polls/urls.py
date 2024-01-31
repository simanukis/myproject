from django.urls import path
from . import views

# app_name = 'polls'

urlpatterns = [    
    # path('', views.index, name='index'),
    path('',views.Login,name='login'),
    path('logout',views.Logout,name='logout'),
    path('register', views.AccountRegistration.as_view(), name='register'),
    path('home',views.home,name='home'),
    
    # 以下を追記(views.callのcall_write_data()にデータを送信できるようにする)
    # path('ajax/', views.call_write_data, name='call_write_data'),
]
