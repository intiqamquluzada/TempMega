


import os
import re
import unicodedata


from django.http import JsonResponse
import pandas as pd
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.shortcuts import render, redirect
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth.decorators import login_required
from mega.forms import FirstStepForm, NextStepForm
from services.tables import total, group_insurance
from services.excel_creator import resulter

from django.core.cache import cache


@login_required
@csrf_exempt
def index_view(request):
    form = FirstStepForm()

    if request.method == "POST":
        form = FirstStepForm(request.POST, request.FILES)
        if form.is_valid():
            file1 = form.cleaned_data.get('file1')
            file2 = form.cleaned_data.get('file2')
            file3 = form.cleaned_data.get('file3')
            date = form.cleaned_data.get('date')

            file_paths = {}
            for file, key in zip([file1, file2, file3], ['file1', 'file2', 'file3']):
                filename = unicodedata.normalize('NFKD', file.name).encode('ascii', 'ignore').decode('ascii')
                filename = re.sub(r'[^\w\s-]', '', filename)
                filename = filename.replace(' ', '_')

                fs = FileSystemStorage(location=settings.MEDIA_ROOT)
                saved_filename = fs.save(filename, file)
                file_path = fs.path(saved_filename)
                file_paths[key] = file_path

            request.session['file1'] = file_paths.get('file1')
            request.session['file2'] = file_paths.get('file2')
            request.session['file3'] = file_paths.get('file3')
            request.session['date'] = date

            return redirect("final")

    context = {
        "form": form
    }
    return render(request, "index.html", context)


@login_required
def final_view(request):
    fs = FileSystemStorage(location=settings.MEDIA_ROOT)
    form = NextStepForm()
    
    df_path = request.session.get("file1")
    dz_path = request.session.get("file2")
    dfs_path = request.session.get("file3")
    date = request.session.get("date")
    
    #df = pd.read_excel(df_path)

    #dz = pd.read_excel(dz_path, engine='openpyxl')

    #dfs = pd.read_excel(dfs_path)
    df = cache.get('df_data')
    if df is None:  
        df = pd.read_excel(df_path)
        cache.set('df_data', df, timeout=3600)  

    dz = cache.get('dz_data')
    if dz is None:  
        dz = pd.read_excel(dz_path)
        cache.set('dz_data', dz, timeout=3600) 

    dfs = cache.get('dfs_data')
    if dfs is None:  
        dfs = pd.read_excel(dfs_path)
        cache.set('dfs_data', dfs, timeout=3600)  

    download_url = None
    if request.method == 'POST':
        form = NextStepForm(request.POST, request.FILES)
        if form.is_valid():
            file1 = form.cleaned_data['file1']
            file2 = form.cleaned_data['file2']

            insurance_class_3 = request.POST.get('insurance_class_3')
            insurance_class_5 = request.POST.get('insurance_class_5')

            main_class = None

            if insurance_class_5:
                main_class = insurance_class_5
            elif insurance_class_3:
                main_class = insurance_class_3

            x = total(
                df, dz, dfs,
                pd.read_excel(file1), pd.read_excel(file2),
                date, main_class,
                5 if insurance_class_5 else 3
            )
            y = resulter(x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], x[8], x[9], x[10],
                         x[11], x[12], x[13], x[14], x[15], x[16], x[17], x[18], main_class
                         )
            download_url = fs.url(y)

            download_url = download_url.replace('media/media', 'media')
            
    class_data = group_insurance(df)
    context = {
        "five": class_data[0],
        "three": class_data[1],
        "form": form,
        "download_url": download_url,  
    }
    return render(request, "final.html", context)

