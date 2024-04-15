from django.db import models

# Create your models here.
from django.db import models

class CV(models.Model):
    file = models.FileField(upload_to='cv_files/')
    email = models.EmailField(blank=True, null=True)
    contact_number = models.CharField(max_length=20, blank=True, null=True)
    text = models.TextField(blank=True, null=True)
