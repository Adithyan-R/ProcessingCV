# views.py
from django.shortcuts import render, redirect
from .forms import CVUploadForm
from .models import CV
import re
import os
import string
from openpyxl import Workbook
from django.http import HttpResponse

from .utils import extract_text_from_pdf # Import the text extraction function
import string
import subprocess

from docx import Document
import textract
import shutil
from win32com import client as wc
import pythoncom  # Import the pythoncom module for COM initialization

def sanitize_text(text):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    return ''.join(c for c in text if c in valid_chars)

def extract_text_from_doc(doc_path):
    try:
        # Execute antiword command and capture its output
        result = subprocess.run(['antiword', doc_path], capture_output=True, text=True)
        if result.returncode == 0:
            text = result.stdout
        else:
            print(f"Error extracting text from .doc file: {result.stderr}")
            text = ''
    except Exception as e:
        print(f"Error executing antiword: {e}")
        text = ''
    return text

def extract_text_from_docx(docx_path):
    try:
        # Extract text from .docx file using textract
        text = textract.process(docx_path).decode('utf-8')
    except Exception as e:
        print(f"Error extracting text from .docx file: {e}")
        text = ''
    return text

def convert_doc_to_docx(doc_path, docx_path):
    try:
        pythoncom.CoInitialize()  # Initialize COM
        word = wc.Dispatch("Word.Application")  # Create an instance of Word application
        doc = word.Documents.Open(doc_path)  # Open the .doc file
        doc.SaveAs(docx_path, FileFormat=16)  # Save as .docx format
        doc.Close()  # Close the document
        word.Quit()  # Quit Word application
        return True
    except Exception as e:
        print(f"Error converting {doc_path} to {docx_path}: {e}")
        return False

def process_cv(cv_file):
    _, file_extension = os.path.splitext(cv_file.name)
    text = ''  # Initialize text variable with a default value

    if file_extension.lower() == '.pdf':
        text = extract_text_from_pdf(cv_file.path)
    elif file_extension.lower() in ['.doc', '.docx']:
        if file_extension.lower() == '.doc':
            docx_path = cv_file.path + 'x'  # Create .docx file path
            if not convert_doc_to_docx(cv_file.path, docx_path):
                print("Conversion from .doc to .docx failed.")
                return {'email': [], 'contact_number': [], 'text': ''}
            text = extract_text_from_docx(docx_path)
            os.remove(docx_path)  # Remove temporary .docx file after extraction
        else:
            text = extract_text_from_docx(cv_file.path)
    elif file_extension.lower() == '.txt':
        with open(cv_file.path, 'r', encoding='utf-8') as txt_file:
            text = txt_file.read()

    # Extract email and contact number from text
    email = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    contact_number = re.findall(r'\b(?:\+91)?[1-9][0-9]{9}\b', text)  # Match 10-digit phone numbers with optional +91 prefix

    # Sanitize the text
    text = sanitize_text(text)


    return {'email': email, 'contact_number': contact_number, 'text': text}


def cv_upload_view(request):
    if request.method == 'POST':
        form = CVUploadForm(request.POST, request.FILES)
        if form.is_valid():
            cv = form.save()
            processed_data = process_cv(cv.file)
            cv.email = ', '.join(processed_data['email'])
            cv.contact_number = ', '.join(processed_data['contact_number'])
            cv.text = processed_data['text']
            cv.save()
            return redirect('cv_download', cv_id=cv.id)
    else:
        form = CVUploadForm()
    return render(request, 'cv_upload.html', {'form': form})

def cv_download_view(request, cv_id):
    cv = CV.objects.get(id=cv_id)
    wb = Workbook()
    ws = wb.active
    ws.append(['Email', 'Contact Number', 'Text'])
    ws.append([cv.email, cv.contact_number, cv.text])
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="cv_data.xlsx"'
    wb.save(response)
    return response




