"""
Django settings for guru project.

Generated by 'django-admin startproject' using Django 5.0.3.

For more information on this file, see
https://docs.djangoproject.com/en/5.0/topics/settings/

For the full list of settings and their values, see
https://docs.djangoproject.com/en/5.0/ref/settings/
"""

import os
from pathlib import Path

# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent


# Quick-start development settings - unsuitable for production
# See https://docs.djangoproject.com/en/5.0/howto/deployment/checklist/

# SECURITY WARNING: keep the secret key used in production secret!
#SECRET_KEY = "django-insecure-p$94$+t5_o^p24a0&6i%&a1wy!)5x@z(_15((db%-qdm)1wq5*"
SECRET_KEY = os.getenv('SECRET_KEY')


# SECURITY WARNING: don't run with debug turned on in production!
#DEBUG = True
DEBUG = False

####################################
# new constants
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True
####################################


ALLOWED_HOSTS = [
    "guruweb.pythonanywhere.com"
]

MEDIA_URL = '/uploads/'
MEDIA_ROOT = BASE_DIR / 'uploads'
# Application definition

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "users",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "guru.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [
            BASE_DIR / "templates",
        ],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "guru.wsgi.application"


# Database
# https://docs.djangoproject.com/en/5.0/ref/settings/#databases

DATABASES = {
    "default": {
        #"ENGINE": "django.db.backends.sqlite3",
        #"NAME": BASE_DIR / "db.sqlite3",
        "ENGINE": "django.db.backends.mysql",
        "NAME": "guruweb$default",
        "USER": "guruweb",
        "PASSWORD": os.getenv("MYSQL_PASSWORD"),
        "HOST": "guruweb.mysql.pythonanywhere-services.com",
        "OPTION": {
            "init_command": "SET NAMES 'utf8mb4';SET sql_mode='STRICT_TRANS_TABLES'",
            "charset": "utf8mb4",
        },        
    }
}


# Password validation
# https://docs.djangoproject.com/en/5.0/ref/settings/#auth-password-validators

AUTH_PASSWORD_VALIDATORS = [
    {
        "NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.MinimumLengthValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.CommonPasswordValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.NumericPasswordValidator",
    },
]


# Internationalization
# https://docs.djangoproject.com/en/5.0/topics/i18n/

LANGUAGE_CODE = "en-us"
#LANGUAGE_CODE = "ru-RU"

TIME_ZONE = "UTC"

USE_I18N = True

USE_TZ = True


# Static files (CSS, JavaScript, Images)
# https://docs.djangoproject.com/en/5.0/howto/static-files/

STATIC_URL = "static/"
STATIC_ROOT = BASE_DIR / 'static/'

STATICFILES_DIRS = [
   os.path.join(BASE_DIR, "static"),
]

# Default primary key field type
# https://docs.djangoproject.com/en/5.0/ref/settings/#default-auto-field

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
