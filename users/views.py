import os
from django.shortcuts import redirect, render
from .forms import LoginUserForm, CreateUserForm, UploadFile
from django.contrib.auth import authenticate, login, logout
from django.core.files.storage import FileSystemStorage
from scripts import ABO_GRR

username = ''

def login_user(request):
    global username
    if request.method == 'POST':
        form = LoginUserForm(request.POST)
        if form.is_valid():
            cd = form.cleaned_data
            user = authenticate(request, username=cd['username'], password=cd['password'])            
            if user and user.is_active:
                login(request, user)
                pro = redirect('program')                
                username = str(user.get_username())
                pro.set_cookie('username', username)
                return pro
    else:
        form = LoginUserForm()
    return render(request, 'users/login.html', {'form': form})

def logout_user(request):
    global username
    username = ''
    logout(request)
    form = LoginUserForm()
    return render(request, 'users/login.html', {'form': form})

def regist_user(request):
    global username
    if request.method == 'POST':
        form = CreateUserForm(request.POST)
        if form.is_valid():
            user = form.save()
            user.set_password(user.password)
            user.save()
            login(request, user, backend='django.contrib.auth.backends.ModelBackend')
            form = LoginUserForm()
            return render(request, 'users/login.html', {'form': form})
    else:
        form = CreateUserForm()
    return render(request, 'users/registration.html', {'form': form, 'username': username})

def program(request):    
    global username
    if request.method == 'POST':        
        form = UploadFile(request.POST, request.FILES) # 
        if form.is_valid():            
            file = form.cleaned_data['file']

            fs = FileSystemStorage()
            fs.save(file.name, file)

            result: str = ABO_GRR.analysisOfAccountingStatements('uploads/' + file.name)
            context = {'username': username, 'result': result}
            os.remove('uploads/' + file.name)
            return render(request, 'users/result.html', context)
    else:
        form = UploadFile()
    return render(request, 'users/program.html', {'form': form, 'username': username})