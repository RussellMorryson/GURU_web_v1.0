from django.contrib import admin
from django.urls import path
from .views import login_user, regist_user, logout_user, program

urlpatterns = [
    path('', login_user, name='login_user'),
    path('registration/', regist_user, name='registration'),
    path('', logout_user, name='logout_user'),
    path('program', program, name='program'),
]