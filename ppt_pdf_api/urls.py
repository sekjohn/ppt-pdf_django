from django.urls import path
from ppt_pdf_api import views
urlpatterns = [
 
    path('conversion', views.conversion_main),
]