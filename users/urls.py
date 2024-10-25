from django.urls import path
from users.views import login_view, logout_view

urlpatterns = [

    path("", login_view, name='login'),
    path("logout/", logout_view, name='logout'),

]