import os
from django.shortcuts import redirect, render
from .forms import LoginUserForm, CreateUserForm, UploadFile
from django.contrib.auth import authenticate, login, logout
from django.core.files.storage import FileSystemStorage
from django.contrib.auth.hashers import make_password
from django.contrib.auth.models import User
from scripts import ABO_GRR

def login_user(request):
    if request.method == 'POST':
        form = LoginUserForm(request.POST)
        if form.is_valid():
            #form.save()
            cd = form.cleaned_data
            user = authenticate(request, username=cd['username'], password=cd['password'])
            if user and user.is_active:
                login(request, user)
                form = UploadFile()
                context = {
                    'username': user.get_username,
                    'email': user.get_email_field_name.__str__()
                }
                return render(request, 'users/program.html', {'form': form}, context)
            form = LoginUserForm()
            return render(request, 'users/login.html', {'form': form})
    else:
        form = LoginUserForm()
    return render(request, 'users/login.html', {'form': form})

def logout_user(request):
    logout(request)
    form = LoginUserForm()
    return render(request, 'users/login.html', {'form': form})

def regist_user(request):
    if request.method == 'POST':
        form = CreateUserForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data['username']
            email = form.cleaned_data['email']
            password = make_password(form.cleaned_data['password'])
            user = User(username=username, email=email, password=password)

            #user = form.save()

            #user.cleaned_data['password'] = make_password(form.cleaned_data['password'])
            login(request, user, backend='django.contrib.auth.backends.ModelBackend')
            form = LoginUserForm()
            return render(request, 'users/login.html', {'form': form})
    else:
        form = CreateUserForm()
    return render(request, 'users/registration.html', {'form': form})

def program(request):
    if request.method == 'POST':
        form = UploadFile(request.POST, request.FILES)
        if form.is_valid():
            file = form.cleaned_data['file']

            fs = FileSystemStorage()
            fs.save(file.name, file)

            result: str = ABO_GRR.analysisOfAccountingStatements('uploads/' + file.name)
            context = {
                'result': result
            }
            os.remove('uploads/' + file.name)
            return render(request, 'users/result.html', context)
    else:
        form = UploadFile()
    return render(request, 'users/program.html', {'form': form})