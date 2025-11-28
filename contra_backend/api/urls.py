from django.urls import path
from .views import FormatStatement

urlpatterns = [
    path('format-statement/', FormatStatement.as_view(), name='foramt-statement'),
]
