from django.urls import path
from mega.views import index_view, final_view

urlpatterns = [
    path("", index_view, name='home'),
    path("final/", final_view, name='final'),
]