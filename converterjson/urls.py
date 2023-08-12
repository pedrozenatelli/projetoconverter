from django.urls import path, re_path

from . import views

urlpatterns = [
    path('', views.start, name='start'),
    path('success/', views.conversion_success, name='conversion_success'),
    re_path(r'^download/(?P<file_name>.+)$',
            views.download_file, name='download_file'),

]
