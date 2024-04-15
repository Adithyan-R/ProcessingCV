from django.urls import path
from .views import cv_upload_view, cv_download_view

urlpatterns = [
    path('upload/', cv_upload_view, name='cv_upload'),
    path('download/<int:cv_id>/', cv_download_view, name='cv_download'),
]
