from django.urls import path

from . import views

urlpatterns = [
    path("delete/", views.delete, name="delete"),
    path("update/", views.update, name="update"),
    path("add/", views.add, name="add"),
    path("", views.index, name="index"),
]