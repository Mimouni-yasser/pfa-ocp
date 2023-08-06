from django.db import models
from django.db import models

# Create your models here.

class IP_field(models.Model):
    IP = models.CharField("IP address", max_length=50)
    MAC = models.CharField("MAC address", max_length=50) 
    comment = models.CharField("Comment", max_length=200)
    device_type = models.CharField( "device type", default="NON DEFINIE", max_length=200)
    DateTime = models.DateTimeField("date added", auto_now_add=True)
    
    def __str__(self) -> str:
        return f"IP: {self.IP} MAC: {self.MAC} COMMENT: {self.comment}"


