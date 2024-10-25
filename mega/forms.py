from django import forms
from django.core.exceptions import ValidationError
import re
import datetime


class FirstStepForm(forms.Form):
    file1 = forms.FileField(label="Müqavilələr jurnalı: ", required=True,
                            widget=forms.ClearableFileInput(attrs={'id': '1'}),
                            )
    file2 = forms.FileField(label="Zərərlər jurnalı: ", required=True,
                            widget=forms.ClearableFileInput(attrs={'id': '2'}))
    file3 = forms.FileField(label="Subriqasiya: ", required=True,
                            widget=forms.ClearableFileInput(attrs={'id': '3'}))
    date = forms.CharField(label="Tarix: ", required=True,
                           widget=forms.TextInput(attrs={'class': 'form-control',
                                                         'placeholder': 'YYYY-MM-DD',
                                                         'id': 'dateInput'}))

    def clean_file1(self):
        file1 = self.cleaned_data.get('file1')
        if not file1:
            raise ValidationError("Fayl 1 seçilməlidir.")
        return file1

    def clean_file2(self):
        file2 = self.cleaned_data.get('file2')
        if not file2:
            raise ValidationError("Fayl 2 seçilməlidir.")
        return file2

    def clean_file3(self):
        file3 = self.cleaned_data.get('file3')
        if not file3:
            raise ValidationError("Fayl 3 seçilməlidir.")
        return file3

    def clean_date(self):
        date = self.cleaned_data.get('date')
        pattern = r'^\d{4}-\d{2}-\d{2}$'

        if not re.match(pattern, date):
            raise ValidationError("Tarixi YYYY-MM-DD formatında daxil edin.")
        try:
            year, month, day = map(int, date.split('-'))
            validated_date = datetime.date(year, month, day)
        except ValueError:
            raise ValidationError("Tarixi düzgün daxil edin.")
        return date

    def __init__(self, *args, **kwargs):
        super(FirstStepForm, self).__init__(*args, **kwargs)
        for field_name, field in self.fields.items():
            field.widget.attrs['class'] = 'form-control'


class NextStepForm(forms.Form):
    file1 = forms.FileField(label="Fayl 1: ", required=True,
                            widget=forms.ClearableFileInput(attrs={'id': '1'}),
                            )
    file2 = forms.FileField(
        label="Fayl 2: ", required=True,
        widget=forms.ClearableFileInput(attrs={'id': '2'}),
    )

    def clean_file1(self):
        file1 = self.cleaned_data.get('file1')
        if not file1:
            raise ValidationError("Fayl 1 seçilməlidir.")
        return file1

    def clean_file2(self):
        file2 = self.cleaned_data.get('file2')
        if not file2:
            raise ValidationError("Fayl 2 seçilməlidir.")
        return file2

    def __init__(self, *args, **kwargs):
        super(NextStepForm, self).__init__(*args, **kwargs)
        for field_name, field in self.fields.items():
            field.widget.attrs['class'] = 'form-control'
