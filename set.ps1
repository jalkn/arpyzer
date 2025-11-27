function arpa {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow
    
    $GREEN = [ConsoleColor]::Green

    Write-Host "🚀 Creating ARPA" -ForegroundColor $YELLOW

    # Create python3 virtual environment
    python3 -m venv .venv

    if ($isWindows) {
        .\.venv\scripts\activate
    } else {
        . ./.venv/bin/Activate.ps1
    }

    # Install required python3 packages
    python3 -m pip install --upgrade pip
    python3 -m pip install pyinstaller django whitenoise django-bootstrap-v5 xlsxwriter openpyxl pandas xlrd>=2.0.1 pdfplumber PyMuPDF msoffcrypto-tool fuzzywuzzy python-Levenshtein psycopg2-binary PyPDF2
    
    # Create Django project
    django-admin startproject arpa
    cd arpa

    # Create core app
    python3 manage.py startapp core

    # Create templates directory structure
    $directories = @(
        "core/src",
        "core/static",
        "core/static/css",
        "core/static/js",
        "core/templates",
        "core/templatetags",
        "core/templates/admin",
        "core/templates/registration"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

# Create runserver.py with basic Django setup
Set-Content -Path "runserver.py" -Value @" 
import os
import sys
import webbrowser
from django.core.management import execute_from_command_line

# IMPORTANT: Set this to your actual project name ('arpa')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'arpa.settings')

def main():
    port = '8000'
    server_address = f'http://127.0.0.1:{port}'

    # Check if we are running in the packaged executable
    if getattr(sys, 'frozen', False):
        print(f"Starting Django server at {server_address}...")

        # Automatically open the browser
        import threading
        def open_browser():
             import time
             time.sleep(1) # Give the server a moment to start
             webbrowser.open_new(server_address)

        threading.Thread(target=open_browser).start()

        # Run the Django server command, disabling the reloader
        execute_from_command_line(['manage.py', 'runserver', f'127.0.0.1:{port}', '--noreload'])

    else:
        # Standard development run
        print("Running Django in development mode...")
        execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()
"@

# Create models.py with cedula as primary key
Set-Content -Path "core/models.py" -Value @" 
from django.db import models
from django.contrib.auth.models import User

class Person(models.Model):
    cedula = models.CharField(max_length=20, primary_key=True)
    nombre_completo = models.CharField(max_length=255)
    correo = models.EmailField(max_length=255, blank=True)
    estado = models.CharField(max_length=50, default='Activo')
    compania = models.CharField(max_length=255, blank=True)
    cargo = models.CharField(max_length=255, blank=True)
    area = models.CharField(max_length=255, blank=True)
    revisar = models.BooleanField(default=False)
    comments = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.nombre_completo} ({self.cedula})"    

class CreditCard(models.Model):
    id = models.AutoField(primary_key=True)

    person = models.ForeignKey(
        Person,
        on_delete=models.SET_NULL,
        related_name='credit_cards', 
        to_field='cedula',
        db_column='cedula', 
        null=True,
        blank=True
    )

    tipo_tarjeta = models.CharField(max_length=50, null=True, blank=True)
    numero_tarjeta = models.CharField(max_length=20, null=True, blank=True) # Número de Tarjeta
    moneda = models.CharField(max_length=10, null=True, blank=True)
    trm_cierre = models.CharField(max_length=50, null=True, blank=True) # Renombrado y tipo ajustado a cómo lo genera tcs.py (string)
    valor_original = models.CharField(max_length=50, null=True, blank=True) # Tipo ajustado a string
    valor_cop = models.CharField(max_length=50, null=True, blank=True) # Agregado de nuevo, tipo ajustado a string
    numero_autorizacion = models.CharField(max_length=100, null=True, blank=True)
    fecha_transaccion = models.DateField(null=True, blank=True)
    dia = models.CharField(max_length=20, null=True, blank=True)
    año = models.CharField(max_length=4, null=True, blank=True)
    descripcion = models.TextField(null=True, blank=True)
    categoria = models.CharField(max_length=255, null=True, blank=True)
    subcategoria = models.CharField(max_length=255, null=True, blank=True)
    zona = models.CharField(max_length=255, null=True, blank=True)
    
    archivo_nombre = models.CharField(max_length=255, null=True, blank=True) 

    def __str__(self):
        return f"{self.descripcion} - {self.valor_cop} (Tarjeta: {self.numero_tarjeta})"

    class Meta:
        verbose_name = "Tarjeta de Crédito"
        verbose_name_plural = "Tarjetas de Crédito"
        unique_together = ('person', 'fecha_transaccion', 'numero_autorizacion', 'valor_original')
"@

# Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from django import forms
from django.utils.html import format_html
from django.urls import reverse
from core.models import Person, CreditCard 

@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = ('cedula', 'nombre_completo', 'cargo', 'area', 'compania', 'estado', 'revisar')
    search_fields = ('cedula', 'nombre_completo', 'correo')
    list_filter = ('estado', 'compania', 'revisar')
    list_editable = ('revisar',)

    # Custom fields to show in detail view
    readonly_fields = ('cedula_with_actions',)  # Added trailing comma to make it a proper tuple

    fieldsets = (
        (None, {
            'fields': ('cedula_with_actions', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'area', 'revisar', 'comments')
        }),
        ('Related Records', {
            'fields': (),
            'classes': ('collapse',)
        }),
    )

    def cedula_with_actions(self, obj):
        if obj.pk:
            change_url = reverse('admin:core_person_change', args=[obj.pk])
            history_url = reverse('admin:core_person_history', args=[obj.pk])
            add_url = reverse('admin:core_person_add')

            return format_html(
                '{} <div class="nowrap">'
                '<a href="{}" class="changelink">Change</a> &nbsp;'
                '<a href="{}" class="historylink">History</a> &nbsp;'
                '<a href="{}" class="addlink">Add another</a>'
                '</div>',
                obj.cedula,
                change_url,
                history_url,
                add_url
            )
        return obj.cedula
    cedula_with_actions.short_description = 'Cedula'

    def get_fieldsets(self, request, obj=None):
        if obj is None:  # Add view
            return [(None, {'fields': ('cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')})]
        return super().get_fieldsets(request, obj)

# Reemplaza la clase CreditCardAdmin con esta:
@admin.register(CreditCard)
class CreditCardAdmin(admin.ModelAdmin):
    # list_display ahora usa el nombre correcto del campo: 'archivo_nombre'
    list_display = (
        'person_link', 'tipo_tarjeta', 'numero_tarjeta', 'fecha_transaccion', 
        'descripcion', 'valor_cop', 'categoria', 'subcategoria', 'archivo_nombre' 
    )
    
    # search_fields también debe usar el nombre correcto
    search_fields = (
        'person__cedula', 'person__nombre_completo', 'numero_tarjeta', 
        'descripcion', 'categoria', 'subcategoria', 'archivo_nombre' 
    )
    
    # Filtros laterales
    list_filter = (
        'tipo_tarjeta', 'moneda', 'categoria', 'subcategoria', 
        'zona', 'person__compania', 'person__cargo', 'fecha_transaccion'
    )
    
    # ... (El resto del código de la clase CreditCardAdmin, incluyendo person_link y raw_id_fields, es correcto)
    
    def person_link(self, obj):
        link = reverse("admin:core_person_change", args=[obj.person.cedula])
        return format_html('<a href="{}">{} ({})</a>', link, obj.person.nombre_completo, obj.person.cedula)

    person_link.short_description = 'Persona'
    person_link.admin_order_field = 'person__nombre_completo'

    raw_id_fields = ('person',)
"@

# Create urls.py for core app
Set-Content -Path "core/urls.py" -Value @"
from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect
from django.urls import path
from django.contrib.auth import views as auth_views
from .views import (main, register_superuser, ImportView, person_list,
                   import_persons,
                   person_details, 
                   export_persons_excel, alerts_list, save_comment, delete_comment, import_tcs, import_categorias,
                   tcs_list, import_personas_tc, export_credit_card_excel) 

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')

        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')

        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')

        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')

    return render(request, 'registration/register.html')

urlpatterns = [
    path('', main, name='main'),
    path('login/', auth_views.LoginView.as_view(template_name='registration/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('register/', register_superuser, name='register'),
    path('import/', ImportView.as_view(), name='import'),
    path('import/persons/', import_persons, name='import_persons'),
    path('import/personas_tc/', import_personas_tc, name='import_personas_tc'),
    path('import/tcs/', import_tcs, name='import_tcs'),
    path('import/categorias/', import_categorias, name='import_categorias'),
    path('persons/', person_list, name='person_list'),
    path('persons/<str:cedula>/', person_details, name='person_details'),
    path('persons/export/excel/', export_persons_excel, name='export_persons_excel'),
    path('alerts/', alerts_list, name='alerts_list'),
    path('persons/<str:cedula>/toggle_revisar/', views.toggle_revisar_status, name='toggle_revisar_status'),
    path('persons/<str:cedula>/save_comment/', save_comment, name='save_comment'),
    path('persons/<str:cedula>/delete_comment/<int:comment_index>/', delete_comment, name='delete_comment'),
    path('tcs/', tcs_list, name='tcs_list'),
    path('tcs/export/excel/', export_credit_card_excel, name='export_credit_card_excel'),
    
]
"@

# Update core/views.py with financial import
Set-Content -Path "core/views.py" -Value @"
import pandas as pd
from datetime import datetime
import os
from django.conf import settings
from django.http import HttpResponseRedirect, HttpResponse
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.mixins import LoginRequiredMixin
from django.views.generic import TemplateView
from django.core.paginator import Paginator
from core.models import Person, CreditCard
from django.db.models import Q
import subprocess
import msoffcrypto
import io
import re
from django.views.decorators.http import require_POST
from django.shortcuts import get_object_or_404, redirect
from . import tcs 
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows 

@login_required
@require_POST
def toggle_revisar_status(request, cedula):
    """
    Toggles the 'revisar' status for a given person.
    Expects a POST request with the person's cedula.
    """
    person = get_object_or_404(Person, cedula=cedula)
    person.revisar = not person.revisar  # Toggle the boolean value
    person.save()

    messages.success(request, f"Revisar status for {person.nombre_completo} ({person.cedula}) updated successfully.")

    # Redirect back to the page the request came from
    next_url = request.META.get('HTTP_REFERER')
    if next_url:
        return redirect(next_url)
    else:
        return redirect('main') 

# Helper function to clean and convert numeric values from strings
def _clean_numeric_value(value):
    if pd.isna(value):
        return None

    str_value = str(value).strip()
    if not str_value:
        return None

    numeric_part = re.sub(r'[^\d.%\-]', '', str_value)

    try:
        if '%' in numeric_part:
            # If it's a percentage, convert to float and divide by 100
            return float(numeric_part.replace('%', '')) / 100
        else:
            # Otherwise, just convert to float
            return float(numeric_part)
    except ValueError:
        return None # Return None if conversion fails

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')

        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')

        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')

        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')

    return render(request, 'registration/register.html')


class ImportView(LoginRequiredMixin, TemplateView):
    template_name = 'import.html'

    def get_context_data(self, **kwargs):
        """
        Overrides the default get_context_data to add counts for persons
        and to gather analysis results from the core/src directory.
        """
        context = super().get_context_data(**kwargs)
        # These counts are fetched from models directly
        context['person_count'] = Person.objects.count()
        context['alerts_count'] = Person.objects.filter(revisar=True).count()
        # Updated to count CreditCard entries
        context['tc_count'] = CreditCard.objects.count()


        analysis_results = []
        core_src_dir = os.path.join(settings.BASE_DIR, 'core', 'src')

        # Helper function to get file status
        def get_file_status(filename, directory=core_src_dir):
            file_path = os.path.join(directory, filename)
            status_info = {'filename': filename, 'records': '-', 'status': 'pending', 'last_updated': None, 'error': None}
            if os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path)
                    status_info['records'] = len(df)
                    status_info['status'] = 'success'
                    status_info['last_updated'] = datetime.fromtimestamp(os.path.getmtime(file_path))
                except Exception as e:
                    status_info['status'] = 'error'
                    status_info['error'] = f"Error reading file: {str(e)}"
            return status_info

        # --- Status for tcs.xlsx ---
        tcs_excel_status = get_file_status('tcs.xlsx')
        analysis_results.append(tcs_excel_status)
        if tcs_excel_status['status'] == 'success':
            context['tc_count'] = tcs_excel_status['records'] # Update tcs_count in context

        # --- Status for categorias.xlsx ---
        categorias_status = get_file_status('categorias.xlsx')
        analysis_results.append(categorias_status)
        if categorias_status['status'] == 'success':
            context['categorias_count'] = categorias_status['records']
        else:
            context['categorias_count'] = 0


        # --- Status for Nets.py output files ---
        analysis_results.append(get_file_status('PersonasTC.xlsx'))

        # Expose PersonasTC.xlsx record count to the template as 'personas_tc_count' when available
        personas_tc_count = 0
        for r in analysis_results:
            if not isinstance(r, dict):
                continue
            if r.get('filename') == 'PersonasTC.xlsx':
                try:
                    if r.get('status') == 'success':
                        personas_tc_count = int(r.get('records') or 0)
                    else:
                        personas_tc_count = 0
                except Exception:
                    personas_tc_count = 0
                break

        context['personas_tc_count'] = personas_tc_count

        context['analysis_results'] = analysis_results
        return context

@login_required
def tcs_list(request):
    """
    Function-based view to display credit card transactions from tcs.xlsx.
    """
    context = {}
    core_src_dir = os.path.join(settings.BASE_DIR, 'core', 'src')
    tcs_excel_path = os.path.join(core_src_dir, 'tcs.xlsx')
    personas_excel_path = os.path.join(core_src_dir, 'PersonasTC.xlsx') # Assuming PersonasTC.xlsx is also in core/src

    transactions_list = []

    print(f"DEBUG: Checking for tcs.xlsx at {tcs_excel_path}")
    if os.path.exists(tcs_excel_path):
        try:
            # Read tcs.xlsx, forcing 'Cedula' column to be read as a string
            tcs_df = pd.read_excel(tcs_excel_path, dtype={'Cedula': str})
            print(f"DEBUG: tcs_df loaded. Shape: {tcs_df.shape}")
            print(f"DEBUG: tcs_df columns RAW: {tcs_df.columns.tolist()}")

            # Standardize tcs_df column names for internal use (remove accents, lowercase, replace spaces with underscores)
            tcs_df.columns = [col.strip().lower().replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('.', '') for col in tcs_df.columns]
            print(f"DEBUG: tcs_df columns STANDARDIZED: {tcs_df.columns.tolist()}")

            # Read PersonasTC.xlsx (if exists)
            personas_df = pd.DataFrame()
            print(f"DEBUG: Checking for PersonasTC.xlsx at {personas_excel_path}")
            if os.path.exists(personas_excel_path):
                try:
                    personas_df = pd.read_excel(personas_excel_path)
                    print(f"DEBUG: personas_df loaded. Shape: {personas_df.shape}")
                    print(f"DEBUG: personas_df columns RAW: {personas_df.columns.tolist()}")
                    personas_df.columns = [col.strip().lower().replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('.', '') for col in personas_df.columns]
                    print(f"DEBUG: personas_df columns STANDARDIZED: {personas_df.columns.tolist()}")
                    
                    if 'cedula' in personas_df.columns:
                        # Use the clean_cedula_format function for robust cleaning
                        personas_df['cedula'] = personas_df['cedula'].apply(tcs.clean_cedula_format)
                        print("DEBUG: Personas 'cedula' column cleaned and standardized.")
                    else:
                        print("WARNING: 'cedula' column not found in PersonasTC.xlsx.")
                        
                except Exception as e:
                    messages.warning(request, f"Error loading PersonasTC.xlsx for joining: {e}")
                    print(f"ERROR: Loading PersonasTC.xlsx: {e}")
            else:
                messages.warning(request, "PersonasTC.xlsx not found. Person details might be missing.")
                print("WARNING: PersonasTC.xlsx not found.")

            # Apply the cleaning function to the 'cedula' column of tcs_df after reading
            if 'cedula' in tcs_df.columns:
                tcs_df['cedula'] = tcs_df['cedula'].apply(tcs.clean_cedula_format)
                print("DEBUG: tcs_df 'cedula' column cleaned and standardized.")
            else:
                print("WARNING: 'cedula' column not found in tcs_df. Cannot merge with Personas data.")
                # If 'cedula' is missing in tcs_df, proceed without person data merge
                merged_df = tcs_df.copy()
                merged_df['nombre_completo'] = None
                merged_df['cargo'] = None
                merged_df['compania'] = None
                merged_df['area'] = None
                print("DEBUG: Proceeding without merging Personas data due to missing 'cedula' in tcs_df.")

            # Perform merge if both DFs and 'cedula' column are present
            if not personas_df.empty and 'cedula' in tcs_df.columns and 'cedula' in personas_df.columns:
                # Select only the necessary columns from personas_df to avoid conflicts
                personas_cols_to_merge = ['cedula', 'nombre_completo', 'cargo', 'compania']
                # Filter to only include columns that actually exist in personas_df
                existing_personas_cols = [col for col in personas_cols_to_merge if col in personas_df.columns]

                merged_df = pd.merge(tcs_df, personas_df[existing_personas_cols],
                                     on='cedula', how='left', suffixes=('', '_person'))
                print(f"DEBUG: DataFrames merged successfully. Merged_df shape: {merged_df.shape}")
                print(f"DEBUG: Merged_df columns AFTER MERGE: {merged_df.columns.tolist()}")
                print(f"DEBUG: First 5 rows of merged_df:\n{merged_df.head()}")
            else:
                # If merge cannot happen, use tcs_df as is and add placeholder columns
                merged_df = tcs_df.copy()
                if 'nombre_completo' not in merged_df.columns: merged_df['nombre_completo'] = None
                if 'cargo' not in merged_df.columns: merged_df['cargo'] = None
                if 'compania' not in merged_df.columns: merged_df['compania'] = None
                if 'area' not in merged_df.columns: merged_df['area'] = None
                print("WARNING: Merge skipped (personas_df empty or 'cedula' column missing). Merged_df is a copy of tcs_df.")
                print(f"DEBUG: Merged_df (unmerged) columns: {merged_df.columns.tolist()}")


            # Map standardized Excel columns to dictionary keys for the template
            for index, row in merged_df.iterrows():
                # Robust date parsing
                fecha_transaccion_raw = row.get('fecha_de_transaccion', None)
                if pd.isna(fecha_transaccion_raw): # Check for pandas NaN values
                    fecha_transaccion = None
                else:
                    try:
                        # Try converting to datetime if it's already a datetime object or string
                        if isinstance(fecha_transaccion_raw, datetime):
                            fecha_transaccion = fecha_transaccion_raw.date() # Get only date part
                        else:
                            # Attempt to parse common date formats if it's a string
                            fecha_transaccion = pd.to_datetime(str(fecha_transaccion_raw)).date()
                    except ValueError:
                        fecha_transaccion = None # If parsing fails, set to None

                # Determine if this cedula maps to an existing Person in the DB and include 'revisar' flag
                person_cedula_val = row.get('cedula', '')
                revisar_flag = False
                try:
                    if person_cedula_val is not None and str(person_cedula_val).strip() != '':
                        p_obj = Person.objects.filter(cedula=str(person_cedula_val).strip()).first()
                        if p_obj:
                            revisar_flag = bool(getattr(p_obj, 'revisar', False))
                except Exception as e:
                    # Defensive: if DB lookup fails for any reason, default to False
                    revisar_flag = False

                person_data = {
                    'cedula': person_cedula_val, # Ensure cedula is always present for URL reversal
                    'nombre_completo': row.get('nombre_completo', 'N/A'),
                    'cargo': row.get('cargo', 'N/A'),
                    'compania': row.get('compania', 'N/A'),
                    'area': row.get('area', 'N/A'),
                    'revisar': revisar_flag,
                }
                
                transaction = {
                    'person': person_data,
                    # Map standardized Excel column names (from tcs.xlsx columns list) to desired keys
                    'tipo_tarjeta': row.get('tipo_de_tarjeta', 'N/A'),
                    'numero_tarjeta': row.get('numero_de_tarjeta', 'N/A'),
                    'moneda': row.get('moneda', 'N/A'),
                    'trm_cierre': row.get('trm_cierre', 'N/A'),
                    'valor_original': row.get('valor_original', 'N/A'),
                    'valor_cop': row.get('valor_cop', 'N/A'),
                    'numero_autorizacion': row.get('numero_de_autorizacion', 'N/A'),
                    'fecha_transaccion': fecha_transaccion, # Use the parsed date
                    'dia': row.get('dia', 'N/A'),
                    'año': row.get('año', 'N/A'), 
                    'descripcion': row.get('descripcion', 'N/A'),
                    'categoria': row.get('categoria', 'N/A'),
                    'subcategoria': row.get('subcategoria', 'N/A'),
                    'zona': row.get('zona', 'N/A'),
                    'tasa_pactada': row.get('tasa_pactada', 'N/A'), # Added
                    'tasa_ea_facturada': row.get('tasa_ea_facturada', 'N/A'), # Added
                    'cargos_y_abonos': row.get('cargos_y_abonos', 'N/A'), # Added
                    'saldo_a_diferir': row.get('saldo_a_diferir', 'N/A'), # Added
                    'cuotas': row.get('cuotas', 'N/A'), # Added
                    'pagina': row.get('pagina', 'N/A'), # Added
                    'tar_x_per': row.get('tar_x_per', 'N/A'), # Added for 'Tar. x Per.'
                    'archivo': row.get('archivo', 'N/A'), # Added for 'Archivo'
                }
                transactions_list.append(transaction)

            # --- START FILTERING LOGIC ---
            q = request.GET.get('q')
            tipo_tarjeta = request.GET.get('tipo_tarjeta')
            categoria = request.GET.get('categoria')
            subcategoria = request.GET.get('subcategoria')
            zona = request.GET.get('zona')
            # New filters from request
            cargo = request.GET.get('cargo')
            compania = request.GET.get('compania')
            area = request.GET.get('area')
            moneda = request.GET.get('moneda')
            dia = request.GET.get('dia')
            año = request.GET.get('año')
            numero_tarjeta_filter = request.GET.get('numero_tarjeta')

            # Get unique values for filter dropdowns (normalize everything to strings)
            tipos_tarjeta = set()
            categorias = set()
            subcategorias = set()
            zonas = set()
            cargos = set()
            companias = set()
            areas = set()
            monedas = set()
            dias = set()
            años = set()
            numeros_tarjeta = set()

            def _safe_str(v):
                # Normalize values to strings; return None for empty/N/A/None
                if v is None:
                    return None
                s = str(v).strip()
                if not s:
                    return None
                if s.upper() == 'N/A':
                    return None
                return s

            # First pass to collect unique values (as strings) to avoid mixed-type sorting errors
            for transaction in transactions_list:
                tt = _safe_str(transaction.get('tipo_tarjeta'))
                if tt:
                    tipos_tarjeta.add(tt)
                cat = _safe_str(transaction.get('categoria'))
                if cat:
                    categorias.add(cat)
                sub = _safe_str(transaction.get('subcategoria'))
                if sub:
                    subcategorias.add(sub)
                z = _safe_str(transaction.get('zona'))
                if z:
                    zonas.add(z)
                # Collect new filter values
                c = _safe_str(transaction['person'].get('cargo'))
                if c:
                    cargos.add(c)
                co = _safe_str(transaction['person'].get('compania'))
                if co:
                    companias.add(co)
                ar = _safe_str(transaction['person'].get('area'))
                if ar:
                    areas.add(ar)
                m = _safe_str(transaction.get('moneda'))
                if m:
                    monedas.add(m)
                d = _safe_str(transaction.get('dia'))
                if d:
                    dias.add(d)
                a = _safe_str(transaction.get('año'))
                if a:
                    años.add(a)
                nt = _safe_str(transaction.get('numero_tarjeta'))
                if nt:
                    numeros_tarjeta.add(nt)

            # Add the sets to context (sorted as strings)
            context['tipos_tarjeta'] = sorted(list(tipos_tarjeta))
            context['categorias'] = sorted(list(categorias))
            context['subcategorias'] = sorted(list(subcategorias))
            context['zonas'] = sorted(list(zonas))
            context['cargos'] = sorted(list(cargos))
            context['companias'] = sorted(list(companias))
            context['areas'] = sorted(list(areas))
            context['monedas'] = sorted(list(monedas))
            context['dias'] = sorted(list(dias))
            context['años'] = sorted(list(años))
            context['numeros_tarjeta'] = sorted(list(numeros_tarjeta))

            # Apply filters
            filtered_list = []
            for transaction in transactions_list:
                # Global search filter
                matches_global = True
                if q:
                    query_lower = q.lower()
                    matches_global = any(
                        query_lower in (str(value) if value is not None else '').lower()
                        for value in [
                            transaction['person'].get('nombre_completo', ''),
                            transaction['person'].get('cedula', ''),
                            transaction['person'].get('cargo', ''),
                            transaction['person'].get('compania', ''),
                            transaction.get('trm_cierre', ''),
                            transaction.get('valor_original', ''),
                            transaction.get('valor_cop', ''),
                            transaction.get('numero_autorizacion', ''),
                            transaction.get('fecha_transaccion', ''),
                            transaction.get('descripcion', ''),
                            transaction.get('tipo_tarjeta', ''),
                            transaction.get('numero_tarjeta', ''),
                            transaction.get('categoria', ''),
                            transaction.get('subcategoria', ''),
                            transaction.get('zona', ''),
                            transaction.get('dia', '')
                        ]
                    )

                # Specific field filters (use normalized strings and case-insensitive compare)
                trans_tipo = _safe_str(transaction.get('tipo_tarjeta'))
                trans_cat = _safe_str(transaction.get('categoria'))
                trans_sub = _safe_str(transaction.get('subcategoria'))
                trans_zona = _safe_str(transaction.get('zona'))
                trans_cargo = _safe_str(transaction['person'].get('cargo'))
                trans_compania = _safe_str(transaction['person'].get('compania'))
                trans_area = _safe_str(transaction['person'].get('area'))
                trans_moneda = _safe_str(transaction.get('moneda'))
                trans_dia = _safe_str(transaction.get('dia'))
                trans_año = _safe_str(transaction.get('año'))
                trans_numero_tarjeta = _safe_str(transaction.get('numero_tarjeta'))

                matches_tipo = (not tipo_tarjeta) or (trans_tipo and trans_tipo.lower() == tipo_tarjeta.lower())
                matches_categoria = (not categoria) or (trans_cat and trans_cat.lower() == categoria.lower())
                matches_subcategoria = (not subcategoria) or (trans_sub and trans_sub.lower() == subcategoria.lower())
                matches_zona = (not zona) or (trans_zona and trans_zona.lower() == zona.lower())
                matches_cargo = (not cargo) or (trans_cargo and trans_cargo.lower() == cargo.lower())
                matches_compania = (not compania) or (trans_compania and trans_compania.lower() == compania.lower())
                matches_area = (not area) or (trans_area and trans_area.lower() == area.lower())
                matches_moneda = (not moneda) or (trans_moneda and trans_moneda.lower() == moneda.lower())
                matches_dia = (not dia) or (trans_dia and trans_dia.lower() == dia.lower())
                matches_año = (not año) or (trans_año and str(trans_año).lower() == str(año).lower())
                matches_numero_tarjeta = (not numero_tarjeta_filter) or (trans_numero_tarjeta and trans_numero_tarjeta.lower() == numero_tarjeta_filter.lower())

                # Add transaction if it matches all active filters
                if matches_global and matches_tipo and matches_categoria and matches_subcategoria and matches_zona and matches_cargo and matches_compania and matches_area and matches_moneda and matches_dia and matches_año and matches_numero_tarjeta:
                    filtered_list.append(transaction)

            transactions_list = filtered_list
            # --- END FILTERING LOGIC ---
            
            print(f"DEBUG: Total transactions in transactions_list: {len(transactions_list)}")
            if transactions_list:
                print(f"DEBUG: First transaction in list: {transactions_list[0]}")
                print(f"DEBUG: Cedula for first transaction: {transactions_list[0].get('person', {}).get('cedula', 'N/A')}")
                print(f"DEBUG: Fecha_transaccion for first transaction: {transactions_list[0].get('fecha_transaccion', 'N/A')}")
            else:
                print("DEBUG: transactions_list is empty.")

            paginator = Paginator(transactions_list, 100)
            page_number = request.GET.get('page')
            page_obj = paginator.get_page(page_number)
            
            context['transactions'] = page_obj
            context['page_obj'] = page_obj
            context['paginator'] = paginator

        except FileNotFoundError:
            messages.error(request, "Error: tcs.xlsx not found in 'core/src/'. Please ensure the PDF processing has been run.")
            print(f"ERROR: tcs.xlsx not found at {tcs_excel_path}")
        except Exception as e:
            messages.error(request, f"Error reading or processing tcs.xlsx: {e}")
            print(f"CRITICAL ERROR: {e}")
    else:
        messages.warning(request, "tcs.xlsx not found. No transaction data to display.")
        print(f"WARNING: tcs.xlsx does not exist.")

    return render(request, 'tcs.html', context)

from django.db.models import Count
import json

@login_required
def main(request):
    """
    Main dashboard view. Gathers counts for various data types and passes them to the home template.
    """
    context = {
        'person_count': Person.objects.count(),
        'alerts_count': Person.objects.filter(revisar=True).count(),
        # Count for active persons
        'active_person_count': Person.objects.filter(estado='Activo').count(),
        # Count for retired persons
        'retired_person_count': Person.objects.filter(estado='Retirado').count(),
        'tc_count': CreditCard.objects.count(),
        # Count for transactions on Sunday
        'domingo_tc_count': CreditCard.objects.filter(dia='Domingo').count(),
    }

    # --- Pivot Table Data: Category per Year (Counts and Totals) ---
    # Fetch all relevant transactions
    transactions = CreditCard.objects.exclude(categoria__isnull=True).exclude(categoria='').exclude(año__isnull=True).exclude(año='').values('categoria', 'año', 'valor_cop')
    
    # Prepare data structure for the pivot table
    pivot_data = {}
    # Get a sorted list of all unique years from the database to use as headers
    db_years = {t['año'] for t in transactions}
    
    # Define the required years and combine them with the years from the database
    required_years = {'2022', '2023', '2024'}
    all_years = sorted([y for y in list(set(db_years) | required_years) if y != 'nan'])

    # Process transactions in Python to calculate counts and sums
    for item in transactions:
        category = item['categoria']
        year = item['año']
        
        # Initialize the nested dictionary for the category if it's not already there
        if category not in pivot_data:
            pivot_data[category] = {y: {'count': 0, 'total_cop': 0.0} for y in all_years}

        # Increment count
        if year in pivot_data[category]:
            pivot_data[category][year]['count'] += 1

            # Safely convert valor_cop string to float and add to total
            try:
                valor_cop_str = item.get('valor_cop', '0').replace('.', '').replace(',', '.')
                pivot_data[category][year]['total_cop'] += float(valor_cop_str)
            except (ValueError, TypeError):
                pass  # Ignore if conversion fails

    # Calculate row and column totals for the pivot table
    column_totals = {year: {'count': 0, 'total_cop': 0.0} for year in all_years}
    grand_total = {'count': 0, 'total_cop': 0.0}

    for category, year_data in pivot_data.items():
        row_total_count = sum(data['count'] for data in year_data.values())
        row_total_cop = sum(data['total_cop'] for data in year_data.values())
        pivot_data[category]['total'] = {'count': row_total_count, 'total_cop': row_total_cop}

        for year, data in year_data.items():
            if year in column_totals:
                column_totals[year]['count'] += data['count']
                column_totals[year]['total_cop'] += data['total_cop']

    grand_total['count'] = sum(totals['count'] for totals in column_totals.values())
    grand_total['total_cop'] = sum(totals['total_cop'] for totals in column_totals.values())

    context['pivot_data'] = pivot_data
    context['pivot_years'] = all_years
    context['pivot_column_totals'] = column_totals
    context['pivot_grand_total'] = grand_total

    # --- Pivot Table Data: Category per Area ---
    area_transactions = CreditCard.objects.select_related('person').exclude(
        categoria__isnull=True
    ).exclude(
        categoria=''
    ).values('categoria', 'person__area', 'valor_cop')

    area_pivot_data = {}
    all_areas = sorted(list(
        Person.objects.exclude(area__isnull=True).exclude(area='').values_list('area', flat=True).distinct()
    ))

    for item in area_transactions:
        category = item['categoria']
        area = item['person__area']
        if not area: continue

        if category not in area_pivot_data:
            area_pivot_data[category] = {a: {'count': 0, 'total_cop': 0.0} for a in all_areas}
        
        if area in area_pivot_data[category]:
            area_pivot_data[category][area]['count'] += 1
            try:
                valor_cop_str = item.get('valor_cop', '0').replace('.', '').replace(',', '.')
                area_pivot_data[category][area]['total_cop'] += float(valor_cop_str)
            except (ValueError, TypeError):
                pass

    for category, area_data in area_pivot_data.items():
        area_pivot_data[category]['total'] = {
            'count': sum(data['count'] for data in area_data.values()),
            'total_cop': sum(data['total_cop'] for data in area_data.values())
        }

    context['area_pivot_data'] = area_pivot_data
    context['all_areas'] = all_areas

    return render(request, 'home.html', context)

@login_required
def import_persons(request):
    """View for importing persons data from Excel files"""
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            # Define the path to save the uploaded file temporarily
            temp_upload_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'uploaded_persons_temp.xlsx')
            with open(temp_upload_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(temp_upload_path)

            # Remove the temporary uploaded file
            os.remove(temp_upload_path)

            # Strip whitespace and convert column names to lowercase for consistent mapping
            df.columns = df.columns.str.strip().str.lower()

            # Define column mapping from Excel columns to model fields
            column_mapping = {
                'id': 'id',
                'nombre completo': 'nombre_completo',
                'correo_normalizado': 'raw_correo',
                'cedula': 'cedula',
                'estado': 'estado',
                'compania': 'compania',
                'cargo': 'cargo',
                'activo': 'activo',
                'business unit': 'area', # <-- CORREGIDO: ahora coincide con la columna en minúsculas
            }

            # Rename columns based on the mapping
            df = df.rename(columns=column_mapping)

            # Ensure 'estado' column exists, if 'activo' is present, use it to determine 'estado'
            if 'activo' in df.columns and 'estado' not in df.columns:
                df['estado'] = df['activo'].apply(lambda x: 'Activo' if x else 'Retirado')
            elif 'estado' not in df.columns:
                df['estado'] = 'Activo' # Default to 'Activo' if neither 'estado' nor 'activo' is present

            # Convert 'cedula' to string type to prevent issues with mixed types
            if 'cedula' in df.columns:
                df['cedula'] = df['cedula'].astype(str)
            else:
                messages.error(request, "Error: 'Cedula' column not found in the Excel file.")
                return HttpResponseRedirect('/import/')

            # Convert nombre_completo to title case if it exists
            if 'nombre_completo' in df.columns:
                df['nombre_completo'] = df['nombre_completo'].str.title()

            # Process 'raw_correo' to create 'correo_to_use' for the database and output
            if 'raw_correo' in df.columns:
                df['correo_to_use'] = df['raw_correo'].str.lower()
            else:
                df['correo_to_use'] = '' # Initialize if no raw email is present

            # Define the columns for the output Excel file including 'Id', 'Estado', and the new 'correo' and 'AREA'
            output_columns = ['Id', 'NOMBRE COMPLETO', 'Cedula', 'Estado', 'Compania', 'CARGO', 'correo', 'AREA']
            output_columns_df = pd.DataFrame(columns=output_columns)

            # Populate the output DataFrame with data from the processed DataFrame
            if 'id' in df.columns:
                output_columns_df['Id'] = df['id']
            if 'nombre_completo' in df.columns:
                output_columns_df['NOMBRE COMPLETO'] = df['nombre_completo']
            if 'cedula' in df.columns:
                output_columns_df['Cedula'] = df['cedula']
            if 'estado' in df.columns:
                output_columns_df['Estado'] = df['estado']
            if 'compania' in df.columns:
                output_columns_df['Compania'] = df['compania']
            if 'cargo' in df.columns:
                output_columns_df['CARGO'] = df['cargo']
            if 'correo_to_use' in df.columns:
                output_columns_df['correo'] = df['correo_to_use']
            # Add 'AREA' to the output DataFrame
            if 'area' in df.columns:
                output_columns_df['AREA'] = df['area']
            else:
                output_columns_df['AREA'] = '' # Ensure column exists even if no data

            # Define the path for the output Excel file
            output_excel_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'Personas.xlsx')

            # Save the filtered and formatted DataFrame to a new Excel file
            output_columns_df.to_excel(output_excel_path, index=False)

            # Iterate over the DataFrame and update/create Person objects in the database
            for _, row in df.iterrows():
                Person.objects.update_or_create(
                    cedula=row['cedula'],
                    defaults={
                        'nombre_completo': row.get('nombre_completo', ''),
                        'correo': row.get('correo_to_use', ''),
                        'estado': row.get('estado', 'Activo'),
                        'compania': row.get('compania', ''),
                        'cargo': row.get('cargo', ''),
                        'area': row.get('area', ''), 
                    }
                )

            messages.success(request, f'Archivo de personas importado exitosamente! {len(df)} registros procesados y Personas.xlsx generado.')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de personas: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
@require_POST
def import_personas_tc(request):
    """
    Handles the upload and processing of the PersonasTC.xlsx file.
    Updates Person records in the database and saves a clean version of the Excel file
    con la columna 'NOMBRE COMPLETO'.
    """
    if request.method == 'POST':
        if 'excel_file' not in request.FILES: 
            messages.error(request, 'No se seleccionó ningún archivo de personas.', extra_tags='import_personas_tc')
            return redirect('import')

        uploaded_file = request.FILES['excel_file'] 
        file_name = "PersonasTC.xlsx"

        # Define el directorio de destino (ej. core/src) y guarda el archivo
        target_directory = os.path.join(settings.BASE_DIR, "core", "src")
        os.makedirs(target_directory, exist_ok=True)
        file_path = os.path.join(target_directory, file_name)

        try:
            # Guarda el archivo cargado
            with open(file_path, 'wb+') as destination:
                for chunk in uploaded_file.chunks():
                    destination.write(chunk)

            # 1. Leer el archivo Excel
            # Read Excel file with cedula as string to prevent float conversion
            df = pd.read_excel(file_path, dtype={'National ID': str, 'Cedula': str, 'Person ID': str})

            # 2. Normalizar encabezados (minúsculas y sin espacios iniciales/finales)
            df.columns = df.columns.str.strip().str.lower()
            
            # --- START: Transformaciones de Datos ---
            
            # Clean cedula field to remove any decimal points and ensure proper format
            def clean_cedula(cedula):
                if pd.isna(cedula):
                    return None
                # Convert to string and handle scientific notation
                cedula_str = str(cedula)
                if 'e' in cedula_str.lower():
                    # Convert scientific notation to regular number
                    cedula_str = f"{float(cedula):.0f}"
                # Remove any decimal points and trailing zeros
                cedula_str = cedula_str.split('.')[0]
                return cedula_str.strip()

            # Apply cedula cleaning to the relevant column
            if 'national id' in df.columns:
                df['national id'] = df['national id'].apply(clean_cedula)
            if 'cedula' in df.columns:
                df['cedula'] = df['cedula'].apply(clean_cedula)
            
            # 3. CONCATENACIÓN: Crear la columna 'nombre_completo'
            name_columns = ['first name', 'middle name', 'last name', 'second last name']
            
            def create_full_name(row):
                """Concatena los nombres, ignorando valores NaN y espacios en blanco."""
                parts = []
                for col in name_columns:
                    if col in row.index and pd.notna(row[col]):
                        parts.append(str(row[col]).strip())
                return ' '.join(parts).strip()

            df['nombre_completo'] = df.apply(create_full_name, axis=1)

            # 4. DEFAULT: Crear la columna 'estado' si no existe y poner 'Activo'
            if 'estado' not in df.columns:
                df['estado'] = 'Activo'
            
            # 5. Mapeo de columnas (del Excel normalizado al campo del modelo/interno)
            user_requested_mapping = {
                'person id': 'id_temp',      
                'national id': 'cedula',     
                'company': 'compania',
                'job title': 'cargo',
                'business unit': 'area', 
                'correo': 'raw_correo',      
                'estado': 'estado',
            }

            final_mapping = {excel_col: model_field for excel_col, model_field in user_requested_mapping.items() if excel_col in df.columns}
            df.rename(columns=final_mapping, inplace=True)
            
            if 'raw_correo' in df.columns:
                df.rename(columns={'raw_correo': 'correo'}, inplace=True)

            # 6. Preparar DataFrame para guardar en la BD y Excel
            model_fields = ['cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'area']
            
            df_to_save = df[[col for col in model_fields if col in df.columns]].copy()
            
            df_to_save.dropna(subset=['cedula'], inplace=True)
            
            # --- END: Transformaciones de Datos ---

            # 7. Lógica de guardado en la Base de Datos (Update or Create)
            for index, row in df_to_save.iterrows():
                Person.objects.update_or_create(
                    cedula=str(row['cedula']).strip(),
                    defaults={
                        'nombre_completo': row.get('nombre_completo', ''),
                        'correo': row.get('correo', ''),
                        'estado': row.get('estado', 'Activo'),
                        'compania': row.get('compania', ''),
                        'cargo': row.get('cargo', ''),
                        'area': row.get('area', ''),
                    }
                )

            # --- NUEVA LÓGICA: Guardar el DataFrame modificado en un archivo Excel ---
            
            # Asegurarse de que las columnas deseadas existan antes de seleccionarlas
            output_columns = ['cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'area']
            df_output = df_to_save[[col for col in output_columns if col in df_to_save.columns]].copy()
            
            # Renombrar la columna 'nombre_completo' a 'NOMBRE COMPLETO' para el archivo de salida
            if 'nombre_completo' in df_output.columns:
                df_output.rename(columns={'nombre_completo': 'NOMBRE COMPLETO'}, inplace=True)
                
            df_output.to_excel(file_path, index=False)
            
            # --- FIN de la nueva lógica ---

            messages.success(request, f'Archivo \"{file_name}\" importado y datos de {len(df_to_save)} personas actualizados correctamente. Archivo de salida con NOMBRE COMPLETO generado.', extra_tags='import_personas_tc')
            
        except Exception as e:
            messages.error(request, f'Error al importar el archivo de personas: {e}', extra_tags='import_personas_tc')
            
    return redirect('import')

@login_required
def person_list(request):
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    area_filter = request.GET.get('area', '')

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query) |
            Q(area__icontains=search_query) 
        )

    if status_filter:
        persons = persons.filter(estado=status_filter)

    if cargo_filter:
        persons = persons.filter(cargo=cargo_filter)

    if compania_filter:
        persons = persons.filter(compania=compania_filter)

    if area_filter: # New: Apply area filter
        persons = persons.filter(area=area_filter)

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by)

    # Convert names to title case for display
    for person in persons:
        person.nombre_completo = person.nombre_completo.title()

    cargos = Person.objects.exclude(cargo='').values_list('cargo', flat=True).distinct().order_by('cargo')
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')
    areas = Person.objects.exclude(area='').values_list('area', flat=True).distinct().order_by('area') 

    paginator = Paginator(persons, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'persons': page_obj,
        'page_obj': page_obj,
        'cargos': cargos,
        'companias': companias,
        'areas': areas,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']},
        'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
    }

    return render(request, 'persons.html', context)

@login_required
def export_persons_excel(request):
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    revisar_filter = request.GET.get('revisar', '') # <--- Add this line to get the 'revisar' parameter

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

    # Apply the 'revisar' filter if present in the URL
    if revisar_filter == 'True': # <--- Add this block
        persons = persons.filter(revisar=True)

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query)
        )

    if status_filter:
        persons = persons.filter(estado=status_filter)

    if cargo_filter:
        persons = persons.filter(cargo=cargo_filter)

    if compania_filter:
        persons = persons.filter(compania=compania_filter)

    # --- Add dynamic column filtering for FinancialReport fields ---

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by).distinct() # Use .distinct() to avoid duplicate persons if related objects cause issues

    # Prepare data for DataFrame
    data = []
    for person in persons:

        row_data = {
            'ID': person.cedula,
            'Nombre Completo': person.nombre_completo,
            'Correo': person.correo,
            'Estado': person.estado,
            'Compañía': person.compania,
            'Cargo': person.cargo,
            'Revisar': 'Sí' if person.revisar else 'No',
            'Comentarios': person.comments,
            'Creado En': person.created_at.strftime('%Y-%m-%d %H:%M:%S') if person.created_at else '',
            'Actualizado En': person.updated_at.strftime('%Y-%m-%d %H:%M:%S') if person.updated_at else '',
        }

    df = pd.DataFrame(data)

    # Create an in-memory Excel file
    excel_file = io.BytesIO()
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Persons', index=False)
    excel_file.seek(0)

    # Create the HTTP response
    response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="persons_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
    return response

from django.db.models import Count
import json

@login_required
def person_details(request, cedula):
    """
    View to display the details of a specific person, including related and credit card transactions.
    """
    myperson = get_object_or_404(Person, cedula=cedula)
    all_transactions = CreditCard.objects.filter(person=myperson).order_by('-fecha_transaccion')

    # --- Pivot Table & Chart Data Calculation ---
    # Fetch all relevant transaction data for the person
    person_transactions_data = all_transactions.exclude(categoria__isnull=True).exclude(categoria='').exclude(año__isnull=True).exclude(año='').values('categoria', 'año', 'valor_cop')

    # Prepare data structure for the pivot table
    person_pivot_data = {}
    db_years = {t['año'] for t in person_transactions_data}
    required_years = {'2022', '2023', '2024'}
    all_person_years = sorted(list(set(db_years) | required_years))

    # Process transactions to calculate counts and sums
    for item in person_transactions_data:
        category = item['categoria']
        year = item['año']
        if category not in person_pivot_data:
            person_pivot_data[category] = {y: {'count': 0, 'total_cop': 0.0} for y in all_person_years}
        if year in person_pivot_data[category]:
            person_pivot_data[category][year]['count'] += 1
            try:
                valor_cop_str = item.get('valor_cop', '0').replace('.', '').replace(',', '.')
                person_pivot_data[category][year]['total_cop'] += float(valor_cop_str)
            except (ValueError, TypeError):
                pass

    # Calculate row and column totals for the person-specific pivot table
    person_column_totals = {year: {'count': 0, 'total_cop': 0.0} for year in all_person_years}
    person_grand_total = {'count': 0, 'total_cop': 0.0}

    for category, year_data in person_pivot_data.items():
        row_total_count = sum(data['count'] for data in year_data.values())
        row_total_cop = sum(data['total_cop'] for data in year_data.values())
        person_pivot_data[category]['total'] = {'count': row_total_count, 'total_cop': row_total_cop}

        for year, data in year_data.items():
            if year in person_column_totals:
                person_column_totals[year]['count'] += data['count']
                person_column_totals[year]['total_cop'] += data['total_cop']

    person_grand_total['count'] = sum(totals['count'] for totals in person_column_totals.values())
    person_grand_total['total_cop'] = sum(totals['total_cop'] for totals in person_column_totals.values())


    # Aggregate transaction counts by year for the person
    yearly_counts = all_transactions.values('año').annotate(count=Count('año')).order_by('año')

    yearly_chart_labels = [item['año'] if item['año'] else 'Sin Año' for item in yearly_counts]
    yearly_chart_data = [item['count'] for item in yearly_counts]

    # --- Filtering Logic ---
    q = request.GET.get('q', '')
    tipo_tarjeta = request.GET.get('tipo_tarjeta', '')
    categoria = request.GET.get('categoria', '')
    subcategoria = request.GET.get('subcategoria', '')
    zona = request.GET.get('zona', '')
    dia = request.GET.get('dia', '')
    año = request.GET.get('año', '')

    # Get unique values for filter dropdowns from all transactions for this person
    tipos_tarjeta = sorted(all_transactions.values_list('tipo_tarjeta', flat=True).distinct())
    categorias = sorted(all_transactions.values_list('categoria', flat=True).distinct())
    subcategorias = sorted(all_transactions.values_list('subcategoria', flat=True).distinct())
    zonas = sorted(all_transactions.values_list('zona', flat=True).distinct())
    dias = sorted(all_transactions.values_list('dia', flat=True).distinct())
    años = sorted(all_transactions.values_list('año', flat=True).distinct())

    # Start with all transactions and apply filters
    filtered_transactions = all_transactions

    if q:
        filtered_transactions = filtered_transactions.filter(
            Q(descripcion__icontains=q) |
            Q(categoria__icontains=q) |
            Q(subcategoria__icontains=q) |
            Q(tipo_tarjeta__icontains=q) |
            Q(dia__icontains=q) |
            Q(año__icontains=q) |
            Q(valor_cop__icontains=q) |
            Q(zona__icontains=q) |
            Q(numero_autorizacion__icontains=q)
        )

    if tipo_tarjeta:
        filtered_transactions = filtered_transactions.filter(tipo_tarjeta=tipo_tarjeta)
    if categoria:
        filtered_transactions = filtered_transactions.filter(categoria=categoria)
    if subcategoria:
        filtered_transactions = filtered_transactions.filter(subcategoria=subcategoria)
    if zona:
        filtered_transactions = filtered_transactions.filter(zona=zona)
    if dia:
        filtered_transactions = filtered_transactions.filter(dia=dia)
    if año:
        filtered_transactions = filtered_transactions.filter(año=año)

    # Pagination
    paginator = Paginator(filtered_transactions, 25)  # Show 25 transactions per page
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'myperson': myperson,
        'alerts_count': Person.objects.filter(revisar=True).count(),
        'credit_card_transactions': page_obj,  # Pass paginated and filtered transactions
        'page_obj': page_obj,
        'paginator': paginator,
        
        # Dropdown options
        'tipos_tarjeta': tipos_tarjeta,
        'categorias': categorias,
        'subcategorias': subcategorias,
        'zonas': zonas,
        'dias': dias,
        'años': años,

        # For preserving filter parameters in pagination links
        'all_params': {k: v for k, v in request.GET.items() if k != 'page'},

        # Pivot table data for the person
        'person_pivot_data': person_pivot_data,
        'person_pivot_years': all_person_years,
        'person_pivot_column_totals': person_column_totals,
        'person_pivot_grand_total': person_grand_total,

        # Yearly chart data
        'yearly_chart_labels': json.dumps(yearly_chart_labels),
        'yearly_chart_data': json.dumps(yearly_chart_data),
    }
    
    return render(request, 'details.html', context)



@login_required
def alerts_list(request):
    """
    View to display persons marked for review (revisar=True).
    """
    search_query = request.GET.get('q', '')
    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.filter(revisar=True)

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query))

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by)

    # Convert names to title case for display
    for person in persons:
        person.nombre_completo = person.nombre_completo.title()

    paginator = Paginator(persons, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'persons': page_obj,
        'page_obj': page_obj,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']},
        'alerts_count': Person.objects.filter(revisar=True).count(), # Pass alerts count
    }

    return render(request, 'alerts.html', context)

@require_POST
def save_comment(request, cedula):
    person = get_object_or_404(Person, cedula=cedula)
    new_comment = request.POST.get('new_comment')

    if new_comment:
        now = datetime.now() 
        timestamp = now.strftime("%d/%m/%Y %H:%M:%S")

        formatted_comment = f"[{timestamp}] {new_comment}"
        
        if person.comments:
            person.comments += f"\n{formatted_comment}"
        else:
            person.comments = formatted_comment
        
        person.save()

    return redirect('person_details', cedula=cedula)

@require_POST
def delete_comment(request, cedula, comment_index):
    person = get_object_or_404(Person, cedula=cedula)

    if person.comments:
        comments_list = person.comments.splitlines()
        
        # Filter out empty strings that might result from splitlines
        comments_list = [comment.strip() for comment in comments_list if comment.strip()]

        if 0 <= comment_index < len(comments_list):
            # Remove the comment at the specified index
            comments_list.pop(comment_index)
            
            # Join the remaining comments back into a single string
            person.comments = "\n".join(comments_list)
            person.save()
        else:
            # Handle invalid index, e.g., log an error or show a message
            pass # For now, silently ignore invalid index
    
    return redirect('person_details', cedula=cedula)

def import_tcs(request):
    if request.method == 'POST':
        pdf_files = request.FILES.getlist('tcs_pdf_files') # Changed to a single input name
        pdf_password = request.POST.get('tcs_pdf_password', '')

        # Check if at least one file type was uploaded
        if not pdf_files:
            messages.error(request, 'No se seleccionaron archivos PDF.', extra_tags='import_tcs')
            return redirect('import')

        # --- Process all uploaded files ---
        input_pdf_dir = os.path.join(settings.BASE_DIR, 'core', 'src', 'extractos')
        output_excel_dir = os.path.join(settings.BASE_DIR, 'core', 'src')
        tcs_excel_path = os.path.join(output_excel_dir, "tcs.xlsx") # Path to the output Excel

        os.makedirs(input_pdf_dir, exist_ok=True) # Ensure input directory exists

        # Clear existing PDFs in the input_pdf_dir before saving new ones.
        for filename in os.listdir(input_pdf_dir):
            if filename.endswith(".pdf"):
                os.remove(os.path.join(input_pdf_dir, filename))

        files_saved = 0
        # Combine both lists of files to save them
        for pdf_file in pdf_files:
            file_path = os.path.join(input_pdf_dir, pdf_file.name)
            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)
                files_saved += 1
            except Exception as e:
                messages.error(request, f"Error saving PDF '{pdf_file.name}': {e}", extra_tags='import_tcs')

        # --- Run processing if files were saved ---
        if files_saved > 0:
            try:
                tcs.pdf_password = pdf_password
                tcs.run_pdf_processing(settings.BASE_DIR, input_pdf_dir, output_excel_dir)

                # --- NEW LOGIC: Load tcs.xlsx into CreditCard model ---
                if os.path.exists(tcs_excel_path):
                    # Force 'Cedula' column to be read as string to prevent float interpretation
                    df_tcs = pd.read_excel(tcs_excel_path, dtype={'Cedula': str})
                    transactions_created = 0
                    transactions_updated = 0

                    # --- NEW: Define clean_cedula_format locally for views.py ---
                    def clean_cedula_format(value):
                        try:
                            if isinstance(value, float) and value.is_integer():
                                return str(int(value))
                            return str(value)
                        except (ValueError, TypeError):
                            return str(value)

                    for index, row in df_tcs.iterrows():
                        raw_cedula = row.get('Cedula') # Get Cedula from the DataFrame row
                        cleaned_cedula = clean_cedula_format(raw_cedula) # Apply cleaning here

                        if pd.isna(cleaned_cedula) or cleaned_cedula == '':
                            print(f"Skipping row {index}: Missing or invalid Cedula for transaction {row.get('Descripción')}")
                            continue

                        # Find or create the Person.
                        person_obj, created = Person.objects.get_or_create(
                            cedula=cleaned_cedula, # Use the cleaned cedula
                            defaults={
                                'nombre_completo': row.get('Tarjetahabiente', ''),
                                'cargo': row.get('CARGO', ''),
                                'compania': '',
                                'area': ''
                            }
                        )
                        if created:
                            print(f"Created new Person: {person_obj.cedula}")

                        card_data = {
                            'person': person_obj,
                            'tipo_tarjeta': row.get('Tipo de Tarjeta', ''),
                            'numero_tarjeta': str(row.get('Número de Tarjeta', '')),
                            'moneda': row.get('Moneda', ''),
                            'trm_cierre': str(row.get('TRM Cierre', '')),
                            'valor_original': str(row.get('Valor Original', '')),
                            'valor_cop': str(row.get('Valor COP', '')),
                            'numero_autorizacion': str(row.get('Número de Autorización', '')),
                            'fecha_transaccion': pd.to_datetime(row.get('Fecha de Transacción'), errors='coerce').date() if pd.notna(row.get('Fecha de Transacción')) else None,
                            'dia': row.get('Día', ''),
                            'año': row.get('Año', ''),
                            'descripcion': row.get('Descripción', ''),
                            'categoria': row.get('Categoría', ''),
                            'subcategoria': row.get('Subcategoría', ''),
                            'zona': row.get('Zona', ''),
                            # No 'cant_tarjetas', 'cargos_abonos', 'archivo', 'cedula_TC' as per new model
                        }

                        # Check for existing transaction to avoid duplicates on re-import
                        lookup_fields = {
                            'person': person_obj,
                            'fecha_transaccion': card_data['fecha_transaccion'],
                            'valor_original': card_data['valor_original']
                        }
                        if card_data['numero_autorizacion'] and card_data['numero_autorizacion'] != 'Sin transacciones':
                            lookup_fields['numero_autorizacion'] = card_data['numero_autorizacion']
                        else:
                            # If no auth number, try to use description to make it somewhat unique
                            lookup_fields['descripcion'] = card_data['descripcion']

                        lookup_fields = {k: v for k, v in lookup_fields.items() if v is not None}


                        try:
                            obj, created = CreditCard.objects.update_or_create( # Changed to CreditCard
                                defaults=card_data,
                                **lookup_fields
                            )
                            if created:
                                transactions_created += 1
                            else:
                                transactions_updated += 1
                        except Exception as e:
                            messages.error(request, f"Error saving transaction row {index} for {cleaned_cedula}: {e}", extra_tags='import_tcs')
                            print(f"Error saving transaction row {index} for {cleaned_cedula}: {e} - Data: {card_data}")

                    messages.success(request, f'Datos de extractos cargados a la base de datos. {transactions_created} creados, {transactions_updated} actualizados.', extra_tags='import_tcs')

                messages.success(request, f'Se procesaron {files_saved} archivos PDF de extractos.', extra_tags='import_tcs')
            except Exception as e:
                messages.error(request, f'Error durante el procesamiento de los PDFs de extractos: {e}', extra_tags='import_tcs')
        else:
            messages.warning(request, 'No se pudieron guardar los archivos PDF para procesar.', extra_tags='import_tcs')

    return redirect('import')


# Function for handling categorias.xlsx upload
def import_categorias(request):
    if request.method == 'POST':
        if 'categorias_excel_file' in request.FILES:
            uploaded_file = request.FILES['categorias_excel_file']
            file_name = "categorias.xlsx"

            target_directory = os.path.join(settings.BASE_DIR, "core", "src")
            os.makedirs(target_directory, exist_ok=True)

            file_path = os.path.join(target_directory, file_name)

            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in uploaded_file.chunks():
                        destination.write(chunk)

                df = pd.read_excel(file_path)
                messages.success(request, f'Archivo "{file_name}" importado correctamente. {len(df)} registros procesados.', extra_tags='import_categorias')
            except Exception as e:
                messages.error(request, f'Error al importar el archivo de categorías: {e}', extra_tags='import_categorias')
        else:
            messages.error(request, 'No se seleccionó ningún archivo de categorías.', extra_tags='import_categorias')
    return redirect('import')


@login_required
def export_credit_card_excel(request):
    """
    Exports the filtered list of CreditCard (TC) objects to an Excel file.
    Handles timezone conversion for Excel compatibility.
    """
    
    # --- Start Filter Logic (Use the same logic as your tcs_list view, if applicable) ---
    # Fetch data including foreign key relationship data (e.g., person details)
    # FIX: Removed .select_related('categoria') as 'categoria' and 'subcategoria' are direct fields.
    queryset = CreditCard.objects.all().select_related('person')
    
    # The filter logic in tcs.html seems to be client-side (JS), so we'll export all
    # or apply server-side filters if they exist in request.GET (e.g., search query 'q')
    q = request.GET.get('q')
    if q:
        # Example server-side filtering for credit cards:
        # FIX: Changed 'comercio' to the correct field name 'descripcion' in Q filter
        queryset = queryset.filter(
            Q(descripcion__icontains=q) | 
            Q(person__cedula__icontains=q) |
            Q(person__nombre_completo__icontains=q)
        )
    # --- End Filter Logic ---

    # Prepare data for DataFrame, explicitly joining related fields
    # FIX: Corrected field names based on Django model choices.
    data = queryset.values(
        'id',
        'person__cedula',
        'person__nombre_completo',
        'fecha_transaccion', 
        'tipo_tarjeta',
        'descripcion',          # Corrected from 'comercio'
        'moneda',               
        'valor_original',       
        'valor_cop',            
        'categoria',            # Corrected from 'categoria__categoria' to direct field
        'subcategoria',         # Corrected from 'categoria__subcategoria' to direct field
        'zona',                 # The field 'zona' is available directly on CreditCard model
        'person__revisar'       # FIX: Access 'revisar' through the 'person' relationship
    )
    df = pd.DataFrame(list(data))

    # Rename columns to be more readable in the Excel file. The order must match queryset.values()
    df.columns = [
        'ID',
        'Cédula',
        'Nombre Completo',
        'Fecha Transacción', 
        'Tipo Tarjeta',
        'Descripción (Comercio)',  # Renamed for clarity
        'Moneda',
        'Valor Original',
        'Valor COP',
        'Categoría',
        'Subcategoría',
        'Zona',
        'Revisar' # The column name remains the same, but the source field changed
    ]

    # FIX: Remove Timezone Information from Datetime Columns
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            # Use 'dt.tz_localize(None)' to remove the timezone information
            df[col] = df[col].dt.tz_localize(None)

    # Prepare Excel workbook and response headers
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    # Set filename
    response['Content-Disposition'] = 'attachment; filename=reporte_tarjetas_credito.xlsx'

    # Create a workbook and sheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Transacciones TC"

    # Write the DataFrame to the worksheet
    for row in dataframe_to_rows(df, header=True, index=False):
        worksheet.append(row)
    
    # Save the workbook to the response
    workbook.save(response)
    
    return response
"@


# Create tcs.py
Set-Content -Path "core/tcs.py" -Value @"
import os
import re
import fitz
import PyPDF2
import pdfplumber
import pandas as pd
from datetime import datetime
import locale
import unicodedata

# --- Configuration (can be modified by views.py) ---
categorias_file = "" # Will be set dynamically
cedulas_file = "" 
pdf_password = ""


# --- TRM Data (Hardcoded - Pre-calculated Monthly Averages) ---
trm_data = {
    "2024/01/01": 3907.86, "2024/02/01": 3932.79, "2024/03/01": 3902.16,
    "2024/04/01": 3871.93, "2024/05/01": 3866.50, "2024/06/01": 4030.73,
    "2024/07/01": 4040.82, "2024/08/01": 4068.79, "2024/09/01": 4188.08,
    "2024/10/01": 4242.02, "2024/11/01": 4398.81, "2024/12/01": 4381.16,
    "2025/01/01": 4296.84, "2025/02/01": 4125.79, "2025/03/01": 4136.21,
    "2025/04/01": 4272.93, "2025/05/01": 4216.79, "2025/06/01": 4110.15,
    "2025/07/01": 4037.03
}

trm_df = pd.DataFrame()
trm_loaded = False

# Function to load TRM data from the hardcoded dictionary (monthly averages)
def load_trm_data():
    global trm_df, trm_loaded
    try:
        # Convert dictionary to DataFrame
        trm_df = pd.DataFrame(list(trm_data.items()), columns=["Fecha", "TRM"])
        # Convert 'Fecha' column to datetime objects (specifically to the first day of the month)
        trm_df["Fecha"] = pd.to_datetime(trm_df["Fecha"], errors='coerce').dt.date
        trm_loaded = True
        print(f"✅ TRM data loaded successfully from hardcoded monthly average dictionary.")
    except Exception as e:
        print(f"⚠️ Error loading TRM data from dictionary: {e}. MC currency conversion will not be available.")


def obtener_trm(fecha):
    # Ensure fecha is a date object for comparison
    if isinstance(fecha, pd.Timestamp):
        fecha = fecha.date()

    if trm_loaded and pd.isna(fecha):
        return ""
    if trm_loaded:
        # Create a "YYYY/MM/01" string for the given date's month
        lookup_month_start = datetime(fecha.year, fecha.month, 1).date()

        # Find the row in trm_df that matches the start of the month
        fila = trm_df[trm_df["Fecha"] == lookup_month_start]
        if not fila.empty:
            return fila["TRM"].values[0]
    return ""

def formato_excel(valor):
    try:
        if isinstance(valor, (int, float)):
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # Handle cases where value might be a string with ',' for thousands and '.' for decimals (e.g., '1.234,56')
        # or just a string with '.' for thousands and ',' for decimals (e.g., '1,234.56')
        # First, remove thousands separators (both '.' and ',') then replace decimal separator to '.'
        s_valor = str(valor).strip()
        if re.match(r'^\d{1,3}(,\d{3})*(\.\d+)?$', s_valor): # Matches 1,234.56 or 1234.56
            numero = float(s_valor.replace(",", ""))
        elif re.match(r'^\d{1,3}(\.\d{3})*(,\d+)?$', s_valor): # Matches 1.234,56 or 1234,56
            numero = float(s_valor.replace(".", "").replace(",", "."))
        else: # Attempt direct conversion
            numero = float(s_valor)

        return f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, AttributeError):
        return valor


# --- Categorias Data Loading ---
categorias_df = pd.DataFrame()
categorias_loaded = False

# Modified to accept base_dir
def load_categorias_data(base_dir):
    global categorias_df, categorias_loaded, categorias_file
    categorias_file = os.path.join(base_dir, "core", "src", "categorias.xlsx") # Set the full path here
    if os.path.exists(categorias_file):
        try:
            categorias_df = pd.read_excel(categorias_file)
            if 'Descripción' in categorias_df.columns:
                categorias_df['Descripción'] = categorias_df['Descripción'].astype(str).str.strip()
                categorias_loaded = True
                print(f"✅ Categorias file '{categorias_file}' loaded successfully.")
            else:
                print(f"⚠️ Categorias file '{categorias_file}' loaded, but 'Descripción' column not found. Categorization will not be available.")
        except Exception as e:
            print(f"⚠️ Error loading Categorias file '{categorias_file}': {e}. Categorization will not be available.")
    else:
        print(f"⚠️ Categorias file '{categorias_file}' not found. Categorization will not be available.")

# --- Cedulas Data Loading (now for PersonasTC.xlsx) ---
cedulas_df = pd.DataFrame()
cedulas_loaded = False

# Modified to accept base_dir and reflect new filename/columns
def load_cedulas_data(base_dir):
    global cedulas_df, cedulas_loaded, cedulas_file
    cedulas_file = os.path.join(base_dir, "core", "src", "PersonasTC.xlsx") # Changed filename here
    if os.path.exists(cedulas_file):
        try:
            df = pd.read_excel(cedulas_file)
            
            # --- START FIX: Normalizar columnas para hacer la verificación insensible a mayúsculas/minúsculas ---
            # 1. Normalizar las columnas del DataFrame cargado a minúsculas/snake_case
            df.columns = df.columns.str.lower().str.replace(' ', '_').str.strip()
            
            # 2. Definir los nombres de columna requeridos en formato normalizado
            required_normalized = ['nombre_completo', 'cedula', 'cargo']
            
            # 3. Verificar la existencia de las columnas normalizadas
            if all(col in df.columns for col in required_normalized):
                
                # 4. Renombrar las columnas a la capitalización específica que el resto del código espera
                df = df.rename(columns={
                    'nombre_completo': 'NOMBRE COMPLETO',
                    'cedula': 'Cedula',
                    'cargo': 'CARGO',
                    # Aseguramos que otras columnas importantes también tengan el casing esperado
                    'compania': 'compania',
                    'area': 'AREA',
                }, errors='ignore') # errors='ignore' para columnas opcionales
                
                # 5. Aplicar la lógica de limpieza y asignar al global cedulas_df
                if 'Cedula' in df.columns:
                    # Apply the clean_cedula_format to the 'Cedula' column upon loading
                    df['Cedula'] = df['Cedula'].apply(clean_cedula_format)
                
                if 'NOMBRE COMPLETO' in df.columns:
                    df['NOMBRE COMPLETO'] = df['NOMBRE COMPLETO'].astype(str).str.title().str.strip()
                    
                cedulas_df = df
                cedulas_loaded = True
                print(f"✅ Personas file '{cedulas_file}' loaded successfully.")
            
            else:
                # Si fallan las columnas, se mantiene el mensaje de advertencia y se evita el merge
                print(f"⚠️ Personas file '{cedulas_file}' loaded, but expected columns ('NOMBRE COMPLETO', 'Cedula', 'CARGO') not found. Personas data will not be available.")
                cedulas_df = None
                cedulas_loaded = False
            # --- END FIX ---
            
        except Exception as e:
            print(f"⚠️ Error loading Personas file '{cedulas_file}': {e}. Personas data will not be available.")
            cedulas_df = None
            cedulas_loaded = False
    else:
        print(f"⚠️ Personas file '{cedulas_file}' not found. Personas data will not be available.")
        cedulas_loaded = False
        cedulas_df = None

# --- NEW: Function to clean Cedula format (e.g., 123.0 to 123) ---
def clean_cedula_format(value):
    try:
        # If it's a float that represents an integer (e.g., 123.0)
        if isinstance(value, float) and value.is_integer():
            return str(int(value)) # Convert to int, then to string
        # For any other type (string, non-integer float, etc.), convert to string and return as is
        return str(value)
    except (ValueError, TypeError):
        # Handles cases where conversion isn't straightforward (e.g., NaN)
        return str(value) # Ensure it's a string even if it's NaN or an unhandled type


# --- Regex for MC (from mc.py) ---
mc_transaccion_regex = re.compile(
    r"(\w{5,})\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+)"
)
mc_nombre_regex = re.compile(r"SEÑOR \(A\):\s*(.*)")
mc_tarjeta_regex = re.compile(r"TARJETA:\s+\*{12}(\d{4})")
mc_moneda_regex = re.compile(r"ESTADO DE CUENTA EN:\s+(DOLARES|PESOS)")

# --- Regex for Visa (from visa.py) ---
visa_pattern_transaccion = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
visa_pattern_tarjeta = re.compile(r"TARJETA:\s+\*{12}(\d{4})")

# --- Regex for Clara ---
clara_card_block_pattern = re.compile(r'(Tarjeta\s+\*\s*\d{4}.*?)(?=Tarjeta\s+\*|\Z)', re.IGNORECASE | re.DOTALL)
clara_card_details_pattern = re.compile(r'Tarjeta\s+\*\s*(\d{4})\s+·\s+(Virtual|Física)\s+(.*?)\s+·\s+ID\s+\d{8}', re.IGNORECASE | re.DOTALL)
clara_transaction_line_pattern = re.compile(r'(\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2})\s+(.*?)\s+(\d{6})\s*(\`$?\s*[\d\.,]+)\s*(\`$?\s*[\d\.,]+)?\s*(\bUSD\b|\bEUR\b|\bPEN\b)?', re.IGNORECASE | re.MULTILINE)

# Modified to accept base_dir, input_folder, and output_folder
def run_pdf_processing(base_dir, input_folder, output_folder):
    """
    Main function to process all PDFs in the input_folder.
    This function should be called from views.py.
    """
    global input_base_folder, output_base_folder 
    global cedulas_df, cedulas_loaded 

    input_base_folder = input_folder
    output_base_folder = output_folder

    all_resultados = [] 

    # Asegúrate de que estas funciones existan y carguen los DataFrames globales
    load_trm_data() 
    load_categorias_data(base_dir) 
    load_cedulas_data(base_dir) # Falla la validación interna por casing

    if os.path.exists(input_base_folder):
        for archivo in sorted(os.listdir(input_base_folder)):
            if archivo.endswith(".pdf"):
                ruta_pdf = os.path.join(input_base_folder, archivo)

                # Use file name to determine card type
                card_type_is_mc = "MC" in archivo.upper() or "MASTERCARD" in archivo.upper()
                card_type_is_visa = "VISA" in archivo.upper()
                card_type_is_clara = "CLARA" in archivo.upper()

                if card_type_is_mc:
                    print(f"📄 Procesando Mastercard: {archivo}")
                    try:
                        with fitz.open(ruta_pdf) as doc:
                            if doc.needs_pass:
                                doc.authenticate(pdf_password)

                            moneda_actual = ""
                            nombre = ""
                            ultimos_digitos = ""
                            tiene_transacciones_mc = False

                            for page_num, page in enumerate(doc, start=1):
                                texto = page.get_text()

                                moneda_match = mc_moneda_regex.search(texto)
                                if moneda_match:
                                    moneda_actual = "USD" if moneda_match.group(1) == "DOLARES" else "COP"

                                if not nombre:
                                    nombre_match = mc_nombre_regex.search(texto)
                                    if nombre_match:
                                        nombre = nombre_match.group(1).strip() 

                                if not ultimos_digitos:
                                    tarjeta_match = mc_tarjeta_regex.search(texto)
                                    if tarjeta_match:
                                        ultimos_digitos = tarjeta_match.group(1).strip()

                                for match in mc_transaccion_regex.finditer(texto):
                                    autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match.groups()

                                    if "ABONO DEBITO AUTOMATICO" in descripcion.upper():
                                        continue

                                    try:
                                        fecha_transaccion = pd.to_datetime(fecha_str, dayfirst=True).date()
                                    except:
                                        fecha_transaccion = None

                                    tipo_cambio = obtener_trm(fecha_transaccion) if moneda_actual == "USD" else ""

                                    trm_cierre_value = formato_excel(str(tipo_cambio)) if tipo_cambio else "1"

                                    all_resultados.append({
                                        "Archivo": archivo,
                                        "Tipo de Tarjeta": "Mastercard", 
                                        "Tarjetahabiente": nombre, 
                                        "Número de Tarjeta": ultimos_digitos,
                                        "Moneda": moneda_actual,
                                        "TRM Cierre": trm_cierre_value, 
                                        "Valor Original": formato_excel(valor_original), 
                                        "Número de Autorización": autorizacion,
                                        "Fecha de Transacción": fecha_transaccion,
                                        "Descripción": descripcion.strip(), 
                                        "Tasa Pactada": formato_excel(tasa_pactada),
                                        "Tasa EA Facturada": formato_excel(tasa_ea),
                                        "Cargos y Abonos": formato_excel(cargo),
                                        "Saldo a Diferir": formato_excel(saldo),
                                        "Cuotas": cuotas,
                                        "Página": page_num,
                                    })
                                    tiene_transacciones_mc = True

                            if not tiene_transacciones_mc and (nombre or ultimos_digitos): 
                                all_resultados.append({
                                    "Archivo": archivo,
                                    "Tipo de Tarjeta": "Mastercard", 
                                    "Tarjetahabiente": nombre, 
                                    "Número de Tarjeta": ultimos_digitos,
                                    "Moneda": "",
                                    "TRM Cierre": "1", 
                                    "Valor Original": "", 
                                    "Número de Autorización": "Sin transacciones",
                                    "Fecha de Transacción": "",
                                    "Descripción": "",
                                    "Tasa Pactada": "",
                                    "Tasa EA Facturada": "",
                                    "Cargos y Abonos": "",
                                    "Saldo a Diferir": "",
                                    "Cuotas": "",
                                    "Página": "",
                                })

                    except Exception as e:
                        print(f"⚠️ Error procesando MC '{archivo}': {e}")

                elif card_type_is_visa:
                    print(f"📄 Procesando Visa: {archivo}")
                    try:
                        with pdfplumber.open(ruta_pdf, password=pdf_password) as pdf:
                            tarjetahabiente_visa = ""
                            tarjeta_visa = ""
                            tiene_transacciones_visa = False
                            last_page_number_visa = 1

                            for page_number, page in enumerate(pdf.pages, start=1):
                                text = page.extract_text()
                                if not text:
                                    continue

                                last_page_number_visa = page_number
                                lines = text.split("\n")

                                for idx, line in enumerate(lines):
                                    line = line.strip()

                                    tarjeta_match_visa = visa_pattern_tarjeta.search(line)
                                    if tarjeta_match_visa:
                                        if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                                            all_resultados.append({
                                                "Archivo": archivo,
                                                "Tipo de Tarjeta": "Visa", 
                                                "Tarjetahabiente": tarjetahabiente_visa,
                                                "Número de Tarjeta": tarjeta_visa,
                                                "Moneda": "",
                                                "TRM Cierre": "1", 
                                                "Valor Original": "", 
                                                "Número de Autorización": "Sin transacciones",
                                                "Fecha de Transacción": "",
                                                "Descripción": "",
                                                "Tasa Pactada": "",
                                                "Tasa EA Facturada": "",
                                                "Cargos y Abonos": "",
                                                "Saldo a Diferir": "",
                                                "Cuotas": "",
                                                "Página": last_page_number_visa,
                                            })

                                        tarjeta_visa = tarjeta_match_visa.group(1)
                                        tiene_transacciones_visa = False 

                                        if idx > 0:
                                            posible_nombre = lines[idx - 1].strip()
                                            posible_nombre = (
                                                posible_nombre
                                                .replace("SEÑOR (A):", "")
                                                .replace("Señor (A):", "")
                                                .replace("SEÑOR:", "")
                                                .replace("Señor:", "")
                                                .strip()
                                            )
                                            if len(posible_nombre.split()) >= 2:
                                                tarjetahabiente_visa = posible_nombre
                                        continue

                                    match_visa = visa_pattern_transaccion.search(' '.join(line.split()))
                                    if match_visa and tarjetahabiente_visa and tarjeta_visa:
                                        autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match_visa.groups()

                                        valor_original_formatted = valor_original.replace(".", "").replace(",", ".")
                                        cargo_formatted = cargo.replace(".", "").replace(",", ".")
                                        saldo_formatted = saldo.replace(".", "").replace(",", ".")

                                        all_resultados.append({
                                            "Archivo": archivo,
                                            "Tipo de Tarjeta": "Visa", 
                                            "Tarjetahabiente": tarjetahabiente_visa, 
                                            "Número de Tarjeta": tarjeta_visa,
                                            "Moneda": "COP", 
                                            "TRM Cierre": "1", 
                                            "Valor Original": formato_excel(valor_original_formatted), 
                                            "Número de Autorización": autorizacion,
                                            "Fecha de Transacción": pd.to_datetime(fecha_str, dayfirst=True).date() if fecha_str else None,
                                            "Descripción": descripcion.strip(), 
                                            "Tasa Pactada": formato_excel(tasa_pactada),
                                            "Tasa EA Facturada": formato_excel(tasa_ea),
                                            "Cargos y Abonos": formato_excel(cargo_formatted),
                                            "Saldo a Diferir": formato_excel(saldo_formatted),
                                            "Cuotas": cuotas,
                                            "Página": page_number,
                                        })
                                        tiene_transacciones_visa = True

                            if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                                all_resultados.append({
                                    "Archivo": archivo,
                                    "Tipo de Tarjeta": "Visa", 
                                    "Tarjetahabiente": tarjetahabiente_visa, 
                                    "Número de Tarjeta": tarjeta_visa,
                                    "Moneda": "",
                                    "TRM Cierre": "1", 
                                    "Valor Original": "", 
                                    "Número de Autorización": "Sin transacciones",
                                    "Fecha de Transacción": "",
                                    "Descripción": "",
                                    "Tasa Pactada": "",
                                    "Tasa EA Facturada": "",
                                    "Cargos y Abonos": "",
                                    "Saldo a Diferir": "",
                                    "Cuotas": "",
                                    "Página": last_page_number_visa,
                                })

                    except Exception as e:
                        print(f"⚠️ Error al procesar Visa '{archivo}': {e}")
                
                elif card_type_is_clara:
                    print(f"📄 Procesando Clara: {archivo}")
                    try:
                        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
                    except locale.Error:
                        try:
                            locale.setlocale(locale.LC_TIME, 'es_ES')
                        except locale.Error:
                            pass

                    try:
                        with open(ruta_pdf, 'rb') as file:
                            reader = PyPDF2.PdfReader(file)
                            for page_number, page in enumerate(reader.pages, 1):
                                full_text = page.extract_text()
                                card_blocks = clara_card_block_pattern.findall(full_text)

                                for block_text in card_blocks:
                                    card_details = clara_card_details_pattern.search(block_text)
                                    if card_details:
                                        card_number = card_details.group(1)
                                        card_type = card_details.group(2)
                                        cardholder_name = card_details.group(3).strip()
                                        
                                        transactions = clara_transaction_line_pattern.findall(block_text)
                                        
                                        if not transactions:
                                            # Add entry for cards with no transactions
                                            all_resultados.append({
                                                "Archivo": archivo, "Tipo de Tarjeta": f"Clara {card_type}", "Tarjetahabiente": cardholder_name,
                                                "Número de Tarjeta": card_number, "Moneda": "", "TRM Cierre": "1", "Valor Original": "",
                                                "Número de Autorización": "Sin transacciones", "Fecha de Transacción": "", "Descripción": "",
                                                "Tasa Pactada": "", "Tasa EA Facturada": "", "Cargos y Abonos": "", "Saldo a Diferir": "",
                                                "Cuotas": "", "Página": page_number,
                                            })

                                        for transaction in transactions:
                                            date_str, description, auth_num, primary_value, secondary_value, moneda = transaction
                                            
                                            moneda = moneda.strip() if moneda else "COP"
                                            valor_cop = primary_value.strip()
                                            valor_original = secondary_value.strip() if secondary_value and (moneda != "COP") else primary_value.strip()

                                            try:
                                                date_obj = datetime.strptime(date_str.strip(), '%d %b %y')
                                                fecha_transaccion = date_obj.date()
                                            except ValueError:
                                                fecha_transaccion = None

                                            tipo_cambio = obtener_trm(fecha_transaccion) if moneda == "USD" else ""
                                            trm_cierre_value = formato_excel(str(tipo_cambio)) if tipo_cambio else "1"

                                            all_resultados.append({
                                                "Archivo": archivo,
                                                "Tipo de Tarjeta": f"Clara {card_type}",
                                                "Tarjetahabiente": cardholder_name,
                                                "Número de Tarjeta": card_number,
                                                "Moneda": moneda,
                                                "TRM Cierre": trm_cierre_value,
                                                "Valor Original": formato_excel(valor_original),
                                                "Número de Autorización": auth_num.strip(),
                                                "Fecha de Transacción": fecha_transaccion,
                                                "Descripción": description.strip(),
                                                "Tasa Pactada": "", "Tasa EA Facturada": "", "Cargos y Abonos": "",
                                                "Saldo a Diferir": "", "Cuotas": "", "Página": page_number,
                                            })

                    except PyPDF2.errors.PdfReadError as e:
                        print(f"⚠️ Error: No se pudo leer '{archivo}'. El archivo puede estar corrupto o encriptado. Error: {e}")
                    except Exception as e:
                        print(f"⚠️ Error inesperado procesando Clara '{archivo}': {e}")

                else:
                    print(f"⏩ Archivo '{archivo}' no reconocido como Mastercard o Visa. Saltando.")

    else:
        print(f"⏩ Carpeta de origen '{input_base_folder}' no encontrada. No hay archivos para procesar.")


    # --- Save All Results to a Single Excel File ---
    if all_resultados:
        df_resultado_final = pd.DataFrame(all_resultados)

        # 1. STANDARDIZACION DE CLAVE DE UNION EN EL DATAFRAME DE RESULTADOS
        df_resultado_final['Tarjetahabiente'] = df_resultado_final['Tarjetahabiente'].astype(str).str.title().str.strip()
        # Crear clave de unión en MAYÚSCULAS para match robusto (ya estaba bien)
        df_resultado_final['Join_Key'] = df_resultado_final['Tarjetahabiente'].str.upper().str.strip()


        # Conversions and Calculations (omitted for brevity, assume correct)
        df_resultado_final['Fecha de Transacción'] = pd.to_datetime(df_resultado_final['Fecha de Transacción'], errors='coerce')
        df_resultado_final['Día'] = df_resultado_final['Fecha de Transacción'].dt.day_name(locale='es_ES').fillna('') 
        df_resultado_final['Año'] = df_resultado_final['Fecha de Transacción'].dt.year.apply(
            lambda x: str(int(x)) if pd.notna(x) else ''
        )
        df_resultado_final['Tar. x Per.'] = df_resultado_final.groupby('Tarjetahabiente')['Número de Tarjeta'].transform('nunique')
        
        def safe_float_conversion(value):
            try:
                if isinstance(value, str):
                    s_value = value.replace(".", "").replace(",", ".")
                    return float(s_value)
                return float(value)
            except (ValueError, TypeError):
                return pd.NA 

        df_resultado_final['Valor Original Num'] = df_resultado_final['Valor Original'].apply(safe_float_conversion)
        df_resultado_final['TRM Cierre Num'] = df_resultado_final['TRM Cierre'].apply(safe_float_conversion)
        df_resultado_final['Valor COP'] = (df_resultado_final['Valor Original Num'] * df_resultado_final['TRM Cierre Num']).apply(lambda x: formato_excel(x) if pd.notna(x) else '')
        df_resultado_final = df_resultado_final.drop(columns=['Valor Original Num', 'TRM Cierre Num'])
        
        # Merge with categorias_df (omitted for brevity, assume correct)
        if categorias_loaded:
            print("Merging all results with categorias.xlsx...")
            df_resultado_final = pd.merge(df_resultado_final, categorias_df[['Descripción', 'Categoría', 'Subcategoría', 'Zona']],
                                    on='Descripción', how='left')
        else:
            df_resultado_final['Categoría'] = ''
            df_resultado_final['Subcategoría'] = ''
            df_resultado_final['Zona'] = ''

        # --- START: MODIFIED MERGE WITH PersonasTC.xlsx (cedulas_df) ---
        # Verificamos si cedulas_df tiene contenido, independientemente de la bandera 'cedulas_loaded'
        if 'cedulas_df' in globals() and cedulas_df is not None and not cedulas_df.empty:
            
            print("\nMerging all results with PersonasTC.xlsx (Cedula, compania, CARGO, AREA) using UPPERCASE keys and **case-insensitive column names**...")
            
            # 1. Prepare Personas DataFrame (cedulas_df)
            df_personas_to_merge = cedulas_df.copy()
            
            # **PASO CRUCIAL:** Normalizar todos los nombres de columna a minúsculas y snake_case para hacerlos case-insensitive
            df_personas_to_merge.columns = df_personas_to_merge.columns.str.lower().str.replace(' ', '_')
            
            # El nombre de la columna que contiene el nombre completo ahora es 'nombre_completo'
            nombre_completo_col = 'nombre_completo'

            if nombre_completo_col in df_personas_to_merge.columns:
                
                # A. Crear la clave de unión en MAYÚSCULAS
                df_personas_to_merge['Join_Key'] = df_personas_to_merge[nombre_completo_col].astype(str).str.upper().str.strip() 
                
                # B. Seleccionar y renombrar las columnas usando los nombres estandarizados (minúsculas)
                # Esto garantiza que 'cedula', 'compania', 'cargo', 'area' sean seleccionados correctamente
                rename_map = {
                    'cedula': 'Cedula',
                    'compania': 'compania',
                    'cargo': 'CARGO',
                    'area': 'AREA'
                }
                
                valid_rename_map = {k: v for k, v in rename_map.items() if k in df_personas_to_merge.columns}
                
                df_personas_to_merge = df_personas_to_merge.rename(columns=valid_rename_map)
                
                # C. Seleccionar solo las columnas necesarias para la unión
                merge_cols_final = ['Join_Key', 'Cedula', 'compania', 'CARGO', 'AREA']
                df_personas_to_merge = df_personas_to_merge[[col for col in merge_cols_final if col in df_personas_to_merge.columns]]
                
                # 2. Realizar la unión Left Merge
                df_resultado_final = pd.merge(
                    df_resultado_final,
                    df_personas_to_merge,
                    on='Join_Key',
                    how='left'
                )
                
                # 3. Finalizar: Limpieza
                df_resultado_final.drop(columns=['Join_Key'], errors='ignore', inplace=True) 
                
                for col in ['Cedula', 'compania', 'CARGO', 'AREA']:
                    if col in df_resultado_final.columns:
                        df_resultado_final[col].fillna('', inplace=True)
                    else:
                        df_resultado_final[col] = '' 
                
                print("✅ Merge completado. Columnas de persona (Cedula, compania, CARGO, AREA) añadidas.")

            else:
                print("FATAL ERROR: 'NOMBRE COMPLETO' column not found in PersonasTC.xlsx, even after standardization. Merge aborted.")
                df_resultado_final['Cedula'] = ''
                df_resultado_final['compania'] = ''
                df_resultado_final['CARGO'] = ''
                df_resultado_final['AREA'] = ''

        else:
            # Bloque original si el DataFrame nunca se cargó o está vacío
            df_resultado_final['Cedula'] = ''
            df_resultado_final['compania'] = ''
            df_resultado_final['CARGO'] = ''
            df_resultado_final['AREA'] = ''
            print("⚠️ Merge no realizado. Se añadieron columnas vacías.")
        # --- END: MODIFIED MERGE WITH PersonasTC.xlsx (cedulas_df) ---

        # Define all expected columns in their desired order (omitted for brevity, assume correct)
        ordered_columns = [
            "Cedula", "compania", "CARGO", "AREA", "Tarjetahabiente", "Tipo de Tarjeta",
            "Número de Tarjeta", "Tar. x Per.", "Moneda", "TRM Cierre", "Valor Original", 
            "Valor COP", "Número de Autorización", "Fecha de Transacción", "Día", "Año",
            "Descripción", "Categoría", "Subcategoría", "Zona", "Tasa Pactada", 
            "Tasa EA Facturada", "Cargos y Abonos", "Saldo a Diferir", "Cuotas", "Página"
        ]

        if "Archivo" in df_resultado_final.columns:
            ordered_columns.append("Archivo")

        df_resultado_final = df_resultado_final[[col for col in ordered_columns if col in df_resultado_final.columns]]

        if 'Cedula' in df_resultado_final.columns:
            df_resultado_final['Cedula'] = df_resultado_final['Cedula'].apply(clean_cedula_format)


        archivo_salida_unificado = "tcs.xlsx"
        ruta_salida_unificado = os.path.join(output_base_folder, archivo_salida_unificado)
        df_resultado_final.to_excel(ruta_salida_unificado, index=False)
        print(f"\n✅ Archivo unificado de extractos generado correctamente en:\n{ruta_salida_unificado}")
        print("\nPrimeras 5 filas del resultado unificado:")
        print(df_resultado_final.head())
    else:
        print("\n⚠️ No se extrajo ningún dato de los archivos PDF (MC o VISA).")
"@

# Create clara.py
Set-Content -Path "core/clara.py" -Value @"
import PyPDF2
import pandas as pd
import re
import datetime
import locale
from PyPDF2.errors import PdfReadError
# Importación necesaria para la normalización de texto
import unicodedata 

def extract_and_parse_data(pdf_path):
    """
    Extracts and parses transaction data from a PDF file, grouped by card.
    """
    parsed_rows = []

    try:
        # Set the locale to Spanish for date formatting
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, 'es_ES')
            except locale.Error:
                pass

        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)

            for page_number, page in enumerate(reader.pages, 1):
                full_text = page.extract_text()
                
                # Regex to capture the entire card block.
                card_block_pattern = re.compile(
                    r'(Tarjeta\s+\*\s*\d{4}.*?)(?=Tarjeta\s+\*|\Z)',
                    re.IGNORECASE | re.DOTALL
                )
                
                # Regex to extract details from the card header.
                card_details_pattern = re.compile(
                    r'Tarjeta\s+\*\s*(\d{4})\s+·\s+(Virtual|Física)\s+(.*?)\s+·\s+ID\s+\d{8}',
                    re.IGNORECASE | re.DOTALL
                )

                card_blocks = card_block_pattern.findall(full_text)

                for block_text in card_blocks:
                    card_details = card_details_pattern.search(block_text)
                    
                    if card_details:
                        card_number = card_details.group(1)
                        card_type = card_details.group(2)
                        cardholder_name = card_details.group(3).strip()
                        
                        print(f"--- Found Card Block for {cardholder_name} ({card_number}) ---")
                        
                        # Regex to capture transaction line with both primary and secondary values
                        transaction_line_pattern = re.compile(
                            r'(\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2})\s+' # 1: Date
                            r'(.*?)\s+'                                                              # 2: Description
                            r'(\d{6})\s*'                                                            # 3: Auth Num (6 digits)
                            r'(`$?\s*[\d\.,]+)\s*'                                                   # 4: Value A (COP Value - Primary)
                            r'(`$?\s*[\d\.,]+)?\s*'                                                  # 5: Value B (Secondary value, either COP or Foreign)
                            r'(\bUSD\b|\bEUR\b|\bPEN\b)?',                                            # 6: Currency (optional)
                            re.IGNORECASE | re.MULTILINE
                        )
                        
                        transactions = transaction_line_pattern.findall(block_text)
                        
                        if not transactions:
                            pass 
                        
                        for transaction in transactions:
                            date = transaction[0].strip()
                            description = transaction[1].strip()
                            auth_num = transaction[2].strip()
                            
                            primary_value = transaction[3].strip()
                            secondary_value = transaction[4].strip()
                            moneda = transaction[5].strip() if transaction[5] else "COP"

                            valor_cop = primary_value

                            if secondary_value and not moneda:
                                valor_original = secondary_value
                            elif secondary_value and moneda:
                                valor_original = secondary_value
                            else:
                                valor_original = primary_value

                            try:
                                date_obj = datetime.datetime.strptime(date, '%d %b %y')
                                day_of_week = date_obj.strftime('%A').title()
                                year = date_obj.year
                            except ValueError:
                                day_of_week = "N/A"
                                year = "N/A"
                            
                            formatted_card_type = f"Clara {card_type}"

                            parsed_rows.append([
                                date, day_of_week, year, description, auth_num, valor_original,
                                moneda, valor_cop, formatted_card_type, card_number,
                                cardholder_name, pdf_path, page_number
                            ])
    except FileNotFoundError:
        print(f"Error: The file '{pdf_path}' was not found.")
        return []
    except PdfReadError as e:
        print(f"Error: Could not read '{pdf_path}'. The file might be corrupted or encrypted.")
        print(f"PyPDF2 error: {e}")
        return []
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return []

    if not parsed_rows:
        print("No transaction data was found in the PDF. Please check the file and the data format.")
    
    return parsed_rows

def save_data_to_excel(df, output_excel_path):
    """
    Saves a pandas DataFrame to an Excel file with all required columns.
    """
    if df.empty:
        print("No data to save.")
        return

    df.to_excel(output_excel_path, index=False)
    print(f"Successfully saved {len(df)} rows to '{output_excel_path}' under the following columns:")
    for col in df.columns:
        print(f" - {col}")
    print(f"Output file: {output_excel_path}")

def read_personas_excel(excel_path):
    """
    Reads the PersonasTC.xlsx file into a DataFrame.
    """
    try:
        personas_df = pd.read_excel(excel_path)
        return personas_df
    except FileNotFoundError:
        print(f"Error: The file '{excel_path}' was not found. Check your working directory.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred while reading '{excel_path}': {e}")
        return pd.DataFrame()

# --- FUNCIÓN DE LIMPIEZA MEJORADA ---
def clean_name(name):
    if pd.isna(name):
        return name
    name = str(name)
    # 1. Normalizar a NFD (Forma de Descomposición Canónica) y codificar a ASCII,
    #    ignorando los caracteres que no se puedan mapear (elimina tildes y eñes).
    name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode('utf-8')
    # 2. Convertir a mayúsculas
    name = name.upper()
    # 3. Eliminar todos los caracteres que NO sean letras o espacios
    name = re.sub(r'[^A-Z\s]', '', name)
    # 4. Eliminar espacios múltiples y trim
    name = re.sub(r'\s+', ' ', name).strip()
    return name


if __name__ == "__main__":
    pdf_file_name = "clara.pdf"
    excel_file_name = "Clara.xlsx"
    personas_file_name = "PersonasTC.xlsx"
    
    column_headers = ["Fecha Transacción", "Día", "Año", "Descripción", "Número de Autorización", "Valor Original", "Moneda", "Valor COP", "Tipo de Tarjeta", "Número de Tarjeta", "Tarjetahabiente", "Archivo", "Página"]

    extracted_data = extract_and_parse_data(pdf_file_name)

    if extracted_data:
        df_clara = pd.DataFrame(extracted_data, columns=column_headers)

        df_personas = read_personas_excel(personas_file_name)

        if not df_personas.empty:
            
            # 1. Aplicar la limpieza agresiva a ambas columnas clave
            df_clara['Tarjetahabiente_MergeKey'] = df_clara['Tarjetahabiente'].apply(clean_name)
            
            # Asegurarse de que la columna de nombres exista en el Excel
            if 'NOMBRE COMPLETO' in df_personas.columns:
                 df_personas['NOMBRE COMPLETO_MergeKey'] = df_personas['NOMBRE COMPLETO'].apply(clean_name)
            else:
                 print("\nError: Columna 'NOMBRE COMPLETO' no encontrada en PersonasTC.xlsx. No se puede continuar con la fusión.")
                 save_data_to_excel(df_clara, excel_file_name)
                 exit()

            try:
                # 2. Fusión usando las nuevas claves limpias y usando las columnas correctas en minúsculas
                final_df = pd.merge(df_clara, 
                                    df_personas[['NOMBRE COMPLETO', 'cedula', 'compania', 'cargo', 'area', 'NOMBRE COMPLETO_MergeKey']], 
                                    how='left', 
                                    left_on='Tarjetahabiente_MergeKey', 
                                    right_on='NOMBRE COMPLETO_MergeKey')
                
                # 3. Eliminar las columnas clave temporales
                final_df.drop(['Tarjetahabiente_MergeKey', 'NOMBRE COMPLETO_MergeKey', 'NOMBRE COMPLETO'], axis=1, inplace=True)
                
                save_data_to_excel(final_df, excel_file_name)
                
            except KeyError as e:
                print(f"\n--- ERROR EN NOMBRE DE COLUMNA ---\nError durante la fusión: Una o más columnas que intentó seleccionar de 'PersonasTC.xlsx' no se encontraron en el archivo: {e}")
                print("Verifique la ortografía y el uso de minúsculas/mayúsculas de las columnas: 'cedula', 'compania', 'cargo', 'area'.")
                print("Guardando solo los datos de transacción extraídos en 'Clara.xlsx'.")
                save_data_to_excel(df_clara, excel_file_name)
        else:
            print("Guardando solo los datos de transacción extraídos en 'Clara.xlsx' (Falta o no se pudo leer 'PersonasTC.xlsx').")
            save_data_to_excel(df_clara, excel_file_name)
    else:
        print("No se encontraron datos de transacciones en el PDF. Por favor, verifique el archivo y el formato de datos.")
"@

# Update project urls.py with proper admin configuration
Set-Content -Path "arpa/urls.py" -Value @"
from django.contrib import admin
from django.urls import include, path
from django.contrib.auth import views as auth_views

# Customize default admin interface
admin.site.site_header = 'A R P A'
admin.site.site_title = 'ARPA Admin Portal'
admin.site.index_title = 'Bienvenido a A R P A'

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('core.urls')), 
    path('accounts/', include('django.contrib.auth.urls')),  
    
]
"@

# Update template filters
Set-Content -Path "core/templatetags/my_filters.py" -Value @"
from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Allows accessing dictionary items with a variable key in templates."""
    if hasattr(dictionary, 'get'):
        return dictionary.get(key)
    return None

@register.filter
def split_lines(value):
    """Splits a string by newlines and returns a list of lines."""
    if isinstance(value, str):
        return value.splitlines()
    return []
"@

# Update template filters __init__.py
Set-Content -Path "core/templatetags/__init__.py" -Value @"
"@

#statics css style
@" 
:root {
    --primary-color: #1a4d7d;
    --primary-hover: #153c61;
    --text-on-primary: white;
    --secondary-color: #6c757d;
    --secondary-hover: #5a6268;
    --success-color: #28a745;
    --success-hover: #218838;
    --warning-color: #ffc107;
    --warning-hover: #e0a800;
    --danger-color: #dc3545;
    --danger-hover: #c82333;
    --info-color: #17a2b8;
    --info-hover: #138496;
    --light-bg: #f8f9fa;
    --card-bg: white;
    --border-color: #e9ecef;
    --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
}

body {
    font-family: var(--font-family);
    background-color: var(--light-bg);
    margin: 0;
}

main {
    padding-top: 20px;
    padding-bottom: 20px;
}

/* Navbar */
.logoIN {
    width: 30px;
    height: 30px;
    background-color: var(--primary-color);
    border-radius: 4px;
    position: relative;
    flex-shrink: 0;
}

.logoIN::before {
    content: "";
    position: absolute;
    width: 100%;
    height: 100%;
    border-radius: 50%;
    top: 30%;
    left: 70%;
    transform: translate(-50%, -50%);
    background-image: linear-gradient(to right, 
        #ffffff 2px, transparent 2px);
    background-size: 4px 100%;
}

.navbar-title {
    font-size: 1.25rem;
    font-weight: 500;
    color: var(--primary-color);
}

.btn-outline-primary {
    color: var(--primary-color);
    border-color: var(--primary-color);
}

.btn-outline-primary:hover {
    background-color: var(--primary-color);
    color: var(--text-on-primary);
}

.btn-outline-secondary {
    color: var(--secondary-color);
    border-color: var(--secondary-color);
}

.btn-outline-secondary:hover {
    background-color: var(--secondary-color);
    color: var(--text-on-primary);
}

.btn-primary {
    background-color: var(--primary-color);
    border-color: var(--primary-color);
    color: var(--text-on-primary);
}

.btn-primary:hover {
    background-color: var(--primary-hover);
    border-color: var(--primary-hover);
}

.btn-secondary {
    background-color: var(--secondary-color);
    border-color: var(--secondary-color);
    color: var(--text-on-primary);
}

.btn-secondary:hover {
    background-color: var(--secondary-hover);
    border-color: var(--secondary-hover);
}

.btn-outline-success {
    color: var(--success-color);
    border-color: var(--success-color);
}

.btn-outline-success:hover {
    background-color: var(--success-color);
    color: var(--text-on-primary);
}

.btn-outline-danger {
    color: var(--danger-color);
    border-color: var(--danger-color);
}

.btn-outline-danger:hover {
    background-color: var(--danger-color);
    color: var(--text-on-primary);
}

/* Cards */
.card {
    border: 1px solid var(--border-color);
    border-radius: .5rem;
    overflow: hidden;
    transition: transform 0.2s, box-shadow 0.2s;
}

.card:hover {
    transform: translateY(-3px);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.card-title {
    font-size: 1rem;
    color: #6c757d;
}

.card-text {
    font-size: 1.75rem;
}

/* Table */
.table th, .table td {
    vertical-align: middle;
    padding: .75rem 1rem;
}

.table thead th {
    background-color: #f1f3f5;
    border-bottom: 2px solid var(--border-color);
    font-weight: 600;
}

.table-hover tbody tr:hover {
    background-color: #f1f3f5;
}

.table-warning {
    background-color: #fff3cd !important;
}

/* Badges */
.badge {
    font-weight: 500;
    text-transform: uppercase;
    font-size: 0.75rem;
}

.bg-success {
    background-color: var(--success-color) !important;
}

.bg-danger {
    background-color: var(--danger-color) !important;
}

.bg-warning {
    background-color: var(--warning-color) !important;
    color: #212529 !important;
}

.text-primary { color: var(--primary-color) !important; }
.text-success { color: var(--success-color) !important; }
.text-danger { color: var(--danger-color) !important; }
.text-info { color: var(--info-color) !important; }
.text-warning { color: var(--warning-color) !important; }
.text-secondary { color: var(--secondary-color) !important; }

/* Form Elements */
.form-control, .form-select {
    border-radius: .25rem;
    border-color: #ced4da;
}

.form-control:focus, .form-select:focus {
    border-color: var(--primary-color);
    box-shadow: 0 0 0 0.25rem rgba(26, 77, 125, 0.25);
}

/* Utilities */
.shadow-sm {
    box-shadow: 0 .125rem .25rem rgba(0,0,0,.075)!important;
}

.border-0 {
    border: none !important;
}

.text-dark {
    color: #212529 !important;
}

/* Footer */
footer {
    border-top: 1px solid var(--border-color);
}
"@ | Out-File -FilePath "core/static/css/style.css" -Encoding utf8

@"
.loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    z-index: 9999;
    display: none;
    justify-content: center;
    align-items: center;
}

.loading-content {
    background-color: white;
    padding: 30px;
    border-radius: 8px;
    text-align: center;
    max-width: 500px;
    width: 90%;
}

.progress {
    height: 20px;
    margin: 20px 0;
}

/* Spinner styles for submit buttons */
.btn .spinner-border {
    margin-right: 8px;
}
"@ | Out-File -FilePath "core/static/css/loading.css" -Encoding utf8

@"
/* Table container styles */
.table-container {
    position: relative;
    overflow: auto;
    max-height: calc(100vh - 300px); /* Adjust this value as needed */
}

/* Make the entire table header sticky */
.table-fixed-header {
    position: sticky;
    top: 0;
    z-index: 10; /* Ensure it stays above the table body */
    background-color: white; /* Fallback background for the header area */
}

/* Apply styles to header cells, but remove individual sticky positioning */
.table-fixed-header th {
    background-color: #f8f9fa; /* Match your table header color */
    /* Remove sticky positioning from individual th elements */
    /* position: sticky; */
    /* top: 0; */
    /* z-index: 20; */
}

/* Add a shadow to the fixed header for visual separation */
.table-fixed-header::after {
    content: '';
    position: absolute;
    left: 0;
    right: 0;
    bottom: -5px;
    height: 5px;
    background: linear-gradient(to bottom, rgba(0,0,0,0.1), transparent);
}

/* Styles for fixed columns */
.table-fixed-column {
    position: sticky;
    right: 0;
    background-color: white;
    z-index: 5;
}

.table-fixed-column::before {
    content: '';
    position: absolute;
    top: 0;
    left: -5px;
    width: 5px;
    height: 100%;
    background: linear-gradient(to right, transparent, rgba(0,0,0,0.1));
}

/* New styles for dynamically frozen columns */
.table-frozen-column {
    position: sticky;
    background-color: white; /* Ensure background is solid when frozen */
    z-index: 6; /* Higher than regular cells but lower than fixed-right column if any */
}

.table-frozen-column::after {
    content: '';
    position: absolute;
    top: 0;
    right: -5px; /* Adjust if shadow is desired on the right */
    width: 5px;
    height: 100%;
    background: linear-gradient(to left, rgba(0,0,0,0.1), transparent);
    pointer-events: none; /* Allows clicks on elements behind the shadow */
}


/* Adjust the z-index for header cells to stay above fixed column */
.table-fixed-header th:last-child {
    z-index: 30;
}

/* Ensure the fixed column stays visible when scrolling */
.table-container {
    overflow: auto;
}

/* Table hover effects */
.table-hover tbody tr:hover {
    background-color: rgba(11, 0, 162, 0.05);
}

/* Style for the freeze button to align it nicely */
.freeze-column-btn {
    margin-right: 5px; /* Space between button and text */
    opacity: 0.5; /* Make it subtle when not active */
}

.freeze-column-btn:hover,
.freeze-column-btn.active {
    opacity: 1; /* More visible when hovered or active */
}
"@ | Out-File -FilePath "core/static/css/freeze.css" -Encoding utf8

@"
.loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    z-index: 9999;
    display: none;
    justify-content: center;
    align-items: center;
}

.loading-content {
    background-color: white;
    padding: 30px;
    border-radius: 8px;
    text-align: center;
    max-width: 500px;
    width: 90%;
}

.progress {
    height: 20px;
    margin: 20px 0;
}

/* Spinner styles for submit buttons */
.btn .spinner-border {
    margin-right: 8px;
}
"@ | Out-File -FilePath "core/static/css/loading.css" -Encoding utf8

# Create loading.js
@"
document.addEventListener('DOMContentLoaded', function() {
    // Get all forms that should show loading
    const forms = document.querySelectorAll('form');
    
    forms.forEach(form => {
        form.addEventListener('submit', function(e) {
            // Only show loading for forms that aren't the search form
            if (!form.classList.contains('no-loading')) {
                // Show loading overlay
                const loadingOverlay = document.getElementById('loadingOverlay');
                if (loadingOverlay) {
                    loadingOverlay.style.display = 'flex';
                }
                
                // Optional: Disable submit button to prevent double submission
                const submitButton = form.querySelector('button[type="submit"]');
                if (submitButton) {
                    submitButton.disabled = true;
                    submitButton.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Procesando...';
                }
            }
        });
    });
});
"@ | Out-File -FilePath "core/static/js/loading.js" -Encoding utf8

$jsContent = @"
document.addEventListener('DOMContentLoaded', function() {
    const table = document.querySelector('.table');
    if (!table) return;

    const freezeButtons = document.querySelectorAll('.freeze-column-btn');
    let frozenColumns = JSON.parse(localStorage.getItem('frozenColumns')) || [];

    function applyFrozenColumns() {
        // Clear any existing frozen classes and inline styles
        document.querySelectorAll('.table-frozen-column').forEach(el => {
            el.classList.remove('table-frozen-column');
            el.style.left = ''; // Clear inline style
        });
        document.querySelectorAll('.freeze-column-btn').forEach(btn => {
            btn.classList.remove('active');
        });

        let currentLeft = 0;
        frozenColumns.forEach(colIndex => {
            const cellsInColumn = table.querySelectorAll(``td:nth-child(`$`{colIndex + 1}), th:nth-child(`$`{colIndex + 1})``);
            cellsInColumn.forEach(cell => {
                cell.classList.add('table-frozen-column');
                cell.style.left = ``$`{currentLeft}px``;
            });

            // Mark the corresponding freeze button as active
            const button = document.querySelector(``.freeze-column-btn[data-column-index="`$`{colIndex}"]``);
            if (button) {
                button.classList.add('active');
            }

            // Calculate the width of the frozen column to offset the next one
            // This is a simplified approach, in a real complex table with variable widths,
            // you might need a more robust calculation or a library.
            const headerCell = table.querySelector(``th:nth-child(`$`{colIndex + 1})``);
            if (headerCell) {
                currentLeft += headerCell.offsetWidth;
            }
        });
    }

    freezeButtons.forEach(button => {
        button.addEventListener('click', function() {
            const columnIndex = parseInt(this.dataset.columnIndex);
            const indexInFrozen = frozenColumns.indexOf(columnIndex);

            if (indexInFrozen > -1) {
                // Column is already frozen, unfreeze it
                frozenColumns.splice(indexInFrozen, 1);
            } else {
                // Column is not frozen, freeze it
                frozenColumns.push(columnIndex);
                frozenColumns.sort((a, b) => a - b); // Keep columns ordered by index
            }

            localStorage.setItem('frozenColumns', JSON.stringify(frozenColumns));
            applyFrozenColumns();
        });
    });

    // Apply frozen columns on initial load
    applyFrozenColumns();

    // Re-apply frozen columns on window resize to adjust 'left' positions
    window.addEventListener('resize', applyFrozenColumns);
});
"@
$jsContent | Out-File -FilePath "core/static/js/freeze_columns.js" -Encoding utf8

# Create custom admin base template
@"
{% extends "admin/base.html" %}

{% block title %}{{ title }} | {{ site_title|default:_('A R P A') }}{% endblock %}

{% block branding %}
<h1 id="site-name"><a href="{% url 'admin:index' %}">{{ site_header|default:_('A R P A') }}</a></h1>
{% endblock %}

{% block nav-global %}{% endblock %}
"@ | Out-File -FilePath "core/templates/admin/base_site.html" -Encoding utf8

# Create master template
@"
<!DOCTYPE html>
<html>
<head>
    <title>{% block title %}ARPA{% endblock %}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    {% load static %}
    <link rel="stylesheet" href="{% static 'css/style.css' %}">
    <link rel="stylesheet" href="{% static 'css/freeze.css' %}">
    <style>
        /* Custom styles for alert messages */
        .alert {
            background-color: #fff3cd !important;  /* Light yellow background */
            border-color: #ffeeba !important;
            color: #856404 !important;
        }
        .alert .btn-close {
            color: #856404 !important;
        }
    </style>
</head>
<body>
    {% if user.is_authenticated %}
    <header class="navbar navbar-expand-lg navbar-light bg-white border-bottom shadow-sm">
        <div class="container-fluid">
            <a href="/" class="d-flex align-items-center me-3">
                <div class="logoIN"></div>
            </a>
            <div class="navbar-title">{% block navbar_title %}{% endblock %}</div>
            <div class="ms-auto d-flex align-items-center gap-2">
                {% block navbar_buttons %}
                <a href="/admin/" class="btn btn-custom-primary" title="Admin">
                    <i class="fas fa-wrench"></i>
                </a>
                <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
                    <i class="fas fa-database"></i>
                </a>
                <form method="post" action="{% url 'logout' %}" class="d-inline">
                    {% csrf_token %}
                    <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
                        <i class="fas fa-sign-out-alt"></i>
                    </button>
                </form>
                {% endblock %}
            </div>
        </div>
    </header>
    {% endif %}
    
    <main class="container-fluid py-4">
        {% if messages %}
            {% for message in messages %}
                <div class="alert alert-{{ message.tags }} alert-dismissible fade show">
                    {{ message }}
                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                </div>
            {% endfor %}
        {% endif %}
        
        {% block content %}
        {% endblock %}
    </main>

    <footer class="footer mt-auto py-3 bg-light border-top">
        <div class="container text-center">
            <span class="text-muted">&copy; A R P A 2 0 2 5</span>
        </div>
    </footer>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="{% static 'js/freeze_columns.js' %}"></script>
</body>
</html>
"@ | Out-File -FilePath "core/templates/master.html" -Encoding utf8

# Create home template
@"
{% extends "master.html" %}
{% load humanize %}
{% load static %}
{% load my_filters %}

{% block title %}A R P A{% endblock %}
{% block navbar_title %}Dashboard{% endblock %}

{% block navbar_buttons %}
<a href="{% url 'person_list' %}" class="btn btn-custom-primary" title="Personas">
    <i class="fas fa-users"></i>
</a>
<a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
    <i class="far fa-credit-card"></i>
</a>
<a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
    <i class="fas fa-bell"></i>
</a>
<a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
    <i class="fas fa-database"></i>
</a>
<form method="post" action="{% url 'logout' %}" class="d-inline">
    {% csrf_token %}
    <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
        <i class="fas fa-sign-out-alt"></i>
    </button>
</form>
{% endblock %}

{% block content %}
<div class="row g-4">
    <div class="col-md-3">
        <a href="{% url 'person_list' %}" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="fas fa-users fa-3x text-primary mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Total de Personas</h5>
                <h2 class="card-text fw-bold text-dark">{{ person_count|intcomma }}</h2>
            </div>
        </a>
    </div>
    <div class="col-md-3">
        <a href="{% url 'tcs_list' %}" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="far fa-credit-card fa-3x text-info mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Transacciones TC</h5>
                <h2 class="card-text fw-bold text-dark">{{ tc_count|intcomma }}</h2>
            </div>
        </a>
    </div>
    <!--
    <div class="col-md-3">
        <a href="{% url 'person_list' %}?status=Activo" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="fas fa-running fa-3x text-success mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Personas Activas</h5>
                <h2 class="card-text fw-bold text-dark">{{ active_person_count|intcomma }}</h2>
            </div>
        </a>
    </div>
    <div class="col-md-3">
        <a href="{% url 'person_list' %}?status=Retirado" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="fas fa-user-times fa-3x text-secondary mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Personas Retiradas</h5>
                <h2 class="card-text fw-bold text-dark">{{ retired_person_count|intcomma }}</h2>
            </div>
        </a>
    </div>
    -->
    <div class="col-md-3">
        <a href="{% url 'alerts_list' %}" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="fas fa-bell fa-3x text-danger mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Alertas</h5>
                <h2 class="card-text fw-bold text-dark">{{ alerts_count|intcomma }}</h2>
            </div>
        </a>
    </div>

    <div class="col-md-3">
        <a href="{% url 'tcs_list' %}?dia=Domingo" class="card h-100 shadow-sm border-0 text-decoration-none">
            <div class="card-body text-center p-4">
                <i class="fas fa-calendar-day fa-3x text-warning mb-3"></i>
                <h5 class="card-title fw-normal mb-1">Transacciones (Domingos)</h5>
                <h2 class="card-text fw-bold text-dark">{{ domingo_tc_count|intcomma }}</h2>
            </div>
        </a>
    </div>
</div>

<div class="row g-4 mt-2">
    <div class="col-lg-12">
        <div class="card h-100 shadow-sm border-0">
            <div class="card-header bg-light">
                <h5 class="mb-0">Transacciones por Categoría y Año</h5>
            </div>
            <div class="card-body">
                {% if pivot_data %}
                    <div class="table-responsive">
                        <table class="table table-striped table-hover table-bordered text-center">
                            <thead class="table-light">
                                <tr>
                                    <th class="text-start">Categoría</th>
                                    {% for year in pivot_years %}
                                        <th colspan="2">{{ year|floatformat:"0" }}</th>
                                    {% endfor %}
                                    <th colspan="2">Total</th>
                                </tr>
                                <tr>
                                    <th></th>
                                    {% for year in pivot_years %}
                                        <th># Trans.</th>
                                        <th>Valor COP</th>
                                    {% endfor %}
                                    <th># Trans.</th>
                                    <th>Valor COP</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for category, year_counts in pivot_data.items %}
                                <tr>
                                    <td class="text-start fw-bold">{{ category }}</td>
                                    {% for year in pivot_years %}
                                        <td>{{ year_counts|get_item:year|get_item:'count'|intcomma }}</td>
                                        <td>`${{ year_counts|get_item:year|get_item:'total_cop'|floatformat:2|intcomma }}</td>
                                    {% endfor %}
                                    <td class="fw-bold">{{ year_counts.total.count|intcomma }}</td>
                                    <td class="fw-bold">`${{ year_counts.total.total_cop|floatformat:2|intcomma }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                            <tfoot class="table-group-divider fw-bold">
                                <tr>
                                    <td class="text-start">Total General</td>
                                    {% for year in pivot_years %}
                                        <td>{{ pivot_column_totals|get_item:year|get_item:'count'|intcomma }}</td>
                                        <td>`${{ pivot_column_totals|get_item:year|get_item:'total_cop'|floatformat:2|intcomma }}</td>
                                    {% endfor %}
                                    <td>{{ pivot_grand_total.count|intcomma }}</td>
                                    <td>`${{ pivot_grand_total.total_cop|floatformat:2|intcomma }}</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                {% else %}
                    <p class="text-muted text-center py-5">No hay datos suficientes para mostrar el gráfico.</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>

<div class="row g-4 mt-2">
    <div class="col-lg-12">
        <div class="card h-100 shadow-sm border-0">
            <div class="card-header bg-light">
                <h5 class="mb-0">Transacciones por Categoría y Área</h5>
            </div>
            <div class="card-body">
                {% if area_pivot_data %}
                    <div class="table-responsive">
                        <table class="table table-striped table-hover table-bordered text-center">
                            <thead class="table-light">
                                <tr>
                                    <th class="text-start">Categoría</th>
                                    {% for area in all_areas %}
                                        <th colspan="2">{{ area }}</th>
                                    {% endfor %}
                                    <th colspan="2">Total</th>
                                </tr>
                                <tr>
                                    <th></th>
                                    {% for area in all_areas %}
                                        <th># Trans.</th>
                                        <th>Valor COP</th>
                                    {% endfor %}
                                    <th># Trans.</th>
                                    <th>Valor COP</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for category, area_counts in area_pivot_data.items %}
                                <tr>
                                    <td class="text-start fw-bold">{{ category }}</td>
                                    {% for area in all_areas %}
                                        <td>{{ area_counts|get_item:area|get_item:'count'|intcomma }}</td>
                                        <td>`${{ area_counts|get_item:area|get_item:'total_cop'|floatformat:2|intcomma }}</td>
                                    {% endfor %}
                                    <td class="fw-bold">{{ area_counts.total.count|intcomma }}</td>
                                    <td class="fw-bold">`${{ area_counts.total.total_cop|floatformat:2|intcomma }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                {% else %}
                    <p class="text-muted text-center py-5">No hay datos suficientes para mostrar el gráfico.</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>

{% endblock %}
"@ | Out-File -FilePath "core/templates/home.html" -Encoding utf8

# Create login template
@"
{% extends "master.html" %}

{% block title %}ARPA{% endblock %}
{% block navbar_title %}ARPA{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if form.errors %}
                    <div class="alert alert-danger">
                        Tu nombre de usuario y clave no coinciden. Por favor intenta de nuevo.
                    </div>
                    {% endif %}

                    {% if next %}
                        {% if user.is_authenticated %}
                        <div class="alert alert-warning">
                            Tu cuenta no tiene acceso a esta pagina. Para continuar,
                            por favor ingresa con una cuenta que tenga acceso.
                        </div>
                        {% else %}
                        <div class="alert alert-info">
                            Por favor accede con tu clave para ver esta página.
                        </div>
                        {% endif %}
                    {% endif %}

                    <form method="post" action="{% url 'login' %}">
                        {% csrf_token %}

                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="id_username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-4">
                            <input type="password" name="password" class="form-control form-control-lg" id="id_password" placeholder="Clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-sign-in-alt"style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'register' %}" class="btn btn-custom-primary" title="Registrarse">  
                                    <i class="fas fa-user-plus fa-lg"></i>
                                </a>
                                <a href="{% url 'password_reset' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-key fa-lg" style="color: orange;"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/login.html" -Encoding utf8

# Create register template
@"
{% extends "master.html" %}

{% block title %}Registro{% endblock %}
{% block navbar_title %}Registro{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if messages %}
                        {% for message in messages %}
                            <div class="alert alert-{% if message.tags == 'error' %}danger{% else %}{{ message.tags }}{% endif %}">
                                {{ message }}
                            </div>
                        {% endfor %}
                    {% endif %}

                    <form method="post" action="{% url 'register' %}">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-3">
                            <input type="email" name="email" class="form-control form-control-lg" id="email" placeholder="Correo" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password1" class="form-control form-control-lg" id="password1" placeholder="Clave" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password2" class="form-control form-control-lg" id="password2" placeholder="Repite tu clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-user-plus fa-lg" style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'login' %}" class="btn btn-custom-primary" title="Ingresar">
                                    <i class="fas fa-sign-in-alt" style="color: rgb(0, 0, 255);"></i>
                                </a>
                                <a href="{% url 'password_reset' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-key fa-lg" style="color: orange;"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/register.html" -Encoding utf8

# Create import template
@"
{% extends "master.html" %}
{% load static %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block navbar_buttons %}
<a href="/" class="btn btn-custom-primary" title="Dashboard">
    <i class="fas fa-chart-pie"></i>
</a>
<a href="{% url 'person_list' %}" class="btn btn-custom-primary" title="Personas">
    <i class="fas fa-users"></i>
</a>
<a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Transacciones TC">
    <i class="far fa-credit-card"></i>
</a>
<a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
    <i class="fas fa-bell"></i>
</a>
<a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
    <i class="fas fa-database"></i> 
</a>
<form method="post" action="{% url 'logout' %}" class="d-inline">
    {% csrf_token %}
    <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
        <i class="fas fa-sign-out-alt"></i>
    </button>
</form>
{% endblock %}

{% block content %}
<div class="loading-overlay" id="loadingOverlay">
    <div class="loading-content">
        <h4>Procesando datos...</h4>
        <div class="progress">
            <div class="progress-bar progress-bar-striped progress-bar-animated"
                 role="progressbar"
                 style="width: 100%"></div>
        </div>
        <p>Por favor espere, esto puede tomar unos segundos.</p>
    </div>
</div>

<link rel="stylesheet" href="{% static 'css/loading.css' %}">
<script src="{% static 'js/loading.js' %}"></script>

<style>
    .upload-form {
        display: flex;
        gap: 0.5rem;
        align-items: center;
    }
    .upload-form .form-control {
        flex: 1;
    }
    .upload-btn {
        padding: 0.375rem 0.75rem;
    }
</style>

<div class="row g-3 mb-1">
    
    <!-- Block commented out
    <div class="col-lg-4 col-md-6 col-sm-12 mb-4">
        <div class="card h-100 border-0 shadow text-center">
            <div class="card-body pb-0">
                <i class="fas fa-users fa-3x text-primary mb-2"></i>
                <h5 class="card-title">Personas</h5>
                <form method="post" enctype="multipart/form-data" action="{% url 'import_persons' %}">
                    {% csrf_token %}
                    <div class="upload-form mb-2">
                        <input type="file" class="form-control form-control-sm" name="excel_file" required>
                        <button type="submit" class="btn btn-secondary btn-sm upload-btn" title="Subir archivo">
                            <i class="fas fa-upload"></i>
                        </button>
                    </div>
                </form>
            </div>
            <div class="card-footer bg-transparent border-0 d-flex justify-content-between align-items-center">
                <span class="badge bg-success">
                    {{ person_count }} Registradas
                </span>
            </div>
        </div>
    </div>
    --> 
    <div class="col-lg-4 col-md-6 col-sm-12 mb-4">
        <div class="card h-100 border-0 shadow text-center">
            <div class="card-body pb-0">
                <i class="fas fa-users fa-3x text-primary mb-2"></i>
                <h5 class="card-title">Personas TC</h5>
                <form method="post" enctype="multipart/form-data" action="{% url 'import_personas_tc' %}">
                    {% csrf_token %}
                    <div class="upload-form mb-2">
                        <input type="file" class="form-control form-control-sm" name="excel_file" required>
                        <button type="submit" class="btn btn-secondary btn-sm upload-btn" title="Subir archivo">
                            <i class="fas fa-upload"></i>
                        </button>
                    </div>
                </form>
            </div>
            <div class="card-footer bg-transparent border-0 d-flex justify-content-between align-items-center">
                <span class="badge bg-success">
                    {{ personas_tc_count }} Registradas
                </span>
            </div>
        </div>
    </div>
    <div class="col-lg-4 col-md-6 col-sm-12 mb-4">
        <div class="card h-100 border-0 shadow text-center">
            <div class="card-body pb-0">
                <i class="fas fa-book fa-3x text-warning mb-2"></i>
                <h5 class="card-title">Categorias</h5>
                <form method="post" enctype="multipart/form-data" action="{% url 'import_categorias' %}" id="categorias-form">
                    {% csrf_token %}
                    <div class="upload-form mb-2">
                        <input type="file" class="form-control form-control-sm" name="categorias_excel_file" required
                               accept=".xlsx, .xls">
                        <button type="submit" class="btn btn-secondary btn-sm upload-btn" title="Subir archivo">
                            <i class="fas fa-upload"></i>
                        </button>
                    </div>
                </form>
            </div>
            <div class="card-footer bg-transparent border-0 d-flex justify-content-between align-items-center">
                <span class="badge bg-success">
                    {{ categorias_count }} Registradas
                </span>
            </div>
        </div>
    </div>

    <div class="col-lg-4 col-md-6 col-sm-12 mb-4">
        <div class="card h-100 border-0 shadow text-center">
            <div class="card-body pb-0">
                <i class="far fa-credit-card fa-3x text-info mb-2"></i>
                <h5 class="card-title">Transacciones TC</h5>
                <form method="post" enctype="multipart/form-data" action="{% url 'import_tcs' %}">
                    {% csrf_token %}
                    <div class="upload-form mb-2">
                        <input type="file" class="form-control form-control-sm" name="tcs_pdf_files" multiple accept=".pdf" required>
                        <button type="submit" class="btn btn-secondary btn-sm upload-btn" title="Subir archivo">
                            <i class="fas fa-upload"></i>
                        </button>
                    </div>
                    <div class="upload-form">
                        <input type="password" class="form-control form-control-sm" name="tcs_pdf_password" placeholder="Clave (opcional)">
                    </div>
                </form>
            </div>
            <div class="card-footer bg-transparent border-0 d-flex justify-content-between align-items-center">
                <span class="badge bg-success">
                    {{ tc_count }} Registradas
                </span>
            </div>
        </div>
    </div>
</div>

<div class="row g-3 mb-1">
    <div class="col-lg-12">
        <div class="card h-100 border-0 shadow">
            <div class="card-header bg-white border-0">
                <h5 class="card-title mb-0">Estado de Archivos Procesados</h5>
            </div>
            <div class="card-body">
                {% if analysis_results %}
                <div class="table-responsive">
                    <table class="table table-striped table-hover mb-0">
                        <thead>
                            <tr>
                                <th scope="col">Archivo</th>
                                <th scope="col" class="text-end">Registros</th>
                                <th scope="col">Estado</th>
                                <th scope="col">Ultima Actualizacion</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for result in analysis_results %}
                            <tr>
                                <td>{{ result.filename }}</td>
                                <td class="text-end">{{ result.records }}</td>
                                <td>
                                    <span class="badge {% if result.status == 'success' %}bg-success{% elif result.status == 'error' %}bg-danger{% else %}bg-secondary{% endif %}">
                                        {% if result.status == 'success' %}
                                            Exitoso
                                        {% elif result.status == 'error' %}
                                            Error
                                        {% else %}
                                            {{ result.status|capfirst }}
                                        {% endif %}
                                    </span>
                                    {% if result.status == 'error' and result.error %}
                                    <small class="text-muted d-block">{{ result.error }}</small>
                                    {% endif %}
                                </td>
                                <td>
                                    {% if result.last_updated %}
                                    {{ result.last_updated|date:"d/m/Y H:i" }}
                                    {% else %}
                                    -
                                    {% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
                {% else %}
                <div class="text-center py-4">
                    <i class="fas fa-info-circle fa-3x text-muted mb-3"></i>
                    <p class="text-muted">No hay resultados de analisis disponibles</p>
                </div>
                {% endif %}
            </div>
            <div class="card-footer bg-transparent border-0">
                <small class="text-muted">Los archivos se procesan en: core/src/</small>
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/import.html" -Encoding utf8

# Create persons template
@"
{% extends "master.html" %}
{% load static %}

{% block title %}Personas{% endblock %}
{% block navbar_title %}Personas{% endblock %}

{% block navbar_buttons %}
<a href="/" class="btn btn-custom-primary" title="Dashboard">
    <i class="fas fa-chart-pie"></i>
</a>
<a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
    <i class="far fa-credit-card"></i>
</a>
<a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
    <i class="fas fa-bell"></i>
</a>
<a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
    <i class="fas fa-database"></i> 
</a>
<a href="{% url 'export_persons_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary" title="Exportar a Excel">
    <i class="fas fa-file-excel"></i>
</a>
<form method="post" action="{% url 'logout' %}" class="d-inline">
    {% csrf_token %}
    <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
        <i class="fas fa-sign-out-alt"></i>
    </button>
</form>
{% endblock %}

{% block content %}
<div class="card mb-4 shadow-sm border-0">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <div class="col-12 d-flex align-items-center mb-3">
                <span class="badge bg-success py-2 px-3 fs-6">{{ page_obj.paginator.count }} registros</span>
                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                {% endif %}
            </div>
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control" 
                       placeholder="Buscar persona o cedula" 
                       value="{{ request.GET.q }}">
            </div>
            <div class="col-md-2">
                <select name="status" class="form-select">
                    <option value="">Estado</option>
                    <option value="Activo" {% if request.GET.status == 'Activo' %}selected{% endif %}>Activo</option>
                    <option value="Retirado" {% if request.GET.status == 'Retirado' %}selected{% endif %}>Retirado</option>
                </select>
            </div>
            <div class="col-md-2">
                <select name="cargo" class="form-select">
                    <option value="">Cargo</option>
                    {% for cargo in cargos %}
                        <option value="{{ cargo }}" {% if request.GET.cargo == cargo %}selected{% endif %}>{{ cargo }}</option>
                    {% endfor %}
                </select>
            </div>
            <div class="col-md-2">
                <select name="area" class="form-select">
                    <option value="">Área</option>
                    {% for area in areas %}
                        <option value="{{ area }}" {% if request.GET.area == area %}selected{% endif %}>{{ area }}</option>
                    {% endfor %}
                </select>
            </div>
            <div class="col-md-2">
                <select name="compania" class="form-select">
                    <option value="">Compañía</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-primary flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-secondary flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<div class="card shadow-sm border-0">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-hover mb-0">
                <thead class="table-light table-fixed-header">
                    <tr>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Revisar
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Cédula
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Nombre
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Cargo
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=area&sort_direction={% if current_order == 'area' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Área
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Correo
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Compañía
                            </a>
                        </th>
                        <th class="py-3">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="text-decoration-none text-dark">
                                Estado
                            </a>
                        </th>
                        <th class="py-3 text-dark">Comentarios</th>
                        <th class="table-fixed-column py-3 text-dark">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <form action="{% url 'toggle_revisar_status' person.cedula %}" method="post" class="d-inline">
                                    {% csrf_token %}
                                    <button type="submit" class="btn btn-sm btn-link p-0" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                        <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}"></i>
                                    </button>
                                </form>
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.area }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' person.cedula %}" 
                                   class="btn btn-sm btn-outline-primary"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="10" class="text-center py-4 text-muted">
                                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/persons.html" -Encoding utf8

@"
{% extends "master.html" %}
{% load static %}

{% block title %}Tarjetas{% endblock %}
{% block navbar_title %}Transacciones TC{% endblock %}

{% block navbar_buttons %}
<div class="ms-auto d-flex align-items-center gap-2">
    <a href="/" class="btn btn-custom-primary" title="Dashboard">
        <i class="fas fa-chart-pie"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary" title="Personas">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Transacciones TC">
        <i class="far fa-credit-card"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        <i class="fas fa-bell"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-database"></i> 
    </a>
    <a href="{% url 'export_credit_card_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary" title="Exportar a Excel">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center" id="filter-form">
            <div class="d-flex align-items-center">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} registros
                </span>
                {% if request.GET.q or request.GET.tipo_tarjeta or request.GET.categoria or request.GET.subcategoria or request.GET.zona or request.GET.cargo or request.GET.compania or request.GET.area or request.GET.moneda or request.GET.dia or request.GET.año or request.GET.numero_tarjeta %}
                {% endif %}
            </div>
            
            <div class="col-md-3">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar cualquier valor..." 
                       value="{{ request.GET.q }}"
                       id="global-search-input">
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="cargo">
                    <option value="">Cargo</option>
                    {% for c in cargos %}
                        <option value="{{ c }}" {% if request.GET.cargo == c %}selected{% endif %}>{{ c }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="area">
                    <option value="">Área</option>
                    {% for a in areas %}
                        <option value="{{ a }}" {% if request.GET.area == a %}selected{% endif %}>{{ a }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="compania">
                    <option value="">Compañía</option>
                    {% for c in companias %}
                        <option value="{{ c }}" {% if request.GET.compania == c %}selected{% endif %}>{{ c }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="tipo_tarjeta" id="cardtype-filter-select">
                    <option value="">Tipo Tarjeta</option>
                    {% for tipo in tipos_tarjeta %}
                        <option value="{{ tipo }}" {% if request.GET.tipo_tarjeta == tipo %}selected{% endif %}>{{ tipo }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="moneda">
                    <option value="">Moneda</option>
                    {% for m in monedas %}
                        <option value="{{ m }}" {% if request.GET.moneda == m %}selected{% endif %}>{{ m }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="dia">
                    <option value="">Día</option>
                    {% for d in dias %}
                        <option value="{{ d }}" {% if request.GET.dia == d %}selected{% endif %}>{{ d }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="año">
                    <option value="">Año</option>
                    {% for a in años %}
                        <option value="{{ a }}" {% if request.GET.año == a|stringformat:"s" %}selected{% endif %}>{{ a }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="categoria" id="category-filter-select">
                    <option value="">Categoría</option>
                    {% for cat in categorias %}
                        <option value="{{ cat }}" {% if request.GET.categoria == cat %}selected{% endif %}>{{ cat }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <div class="col-md-1">
                <select class="form-select form-select-lg" name="subcategoria" id="subcategory-filter-select">
                    <option value="">Subcategoría</option>
                    {% for subcat in subcategorias %}
                        <option value="{{ subcat }}" {% if request.GET.subcategoria == subcat %}selected{% endif %}>{{ subcat }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-1">
                <select class="form-select form-select-lg" name="zona" id="zona-filter-select">
                    <option value="">Zona</option>
                    {% for z in zonas %}
                        <option value="{{ z }}" {% if request.GET.zona == z %}selected{% endif %}>{{ z }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <div class="col-md-1 d-flex gap-2 align-items-center">
                <button type="submit" class="btn btn-primary btn-sm" title="Aplicar filtros de formulario">
                    <i class="fas fa-filter"></i>
                </button>
                <a href="{% url 'tcs_list' %}" class="btn btn-secondary btn-sm" title="Quitar todos los filtros">
                    <i class="fas fa-undo"></i>
                </a>
            </div>
            
            {% if request.GET.q or request.GET.tipo_tarjeta or request.GET.categoria or request.GET.subcategoria or request.GET.zona or request.GET.cargo or request.GET.compania or request.GET.area or request.GET.moneda or request.GET.dia or request.GET.año or request.GET.numero_tarjeta %}
                <div class="mt-2">
                    <span class="badge bg-info">
                        Filtros activos:
                        {% if request.GET.q %}<span class="badge bg-primary">Búsqueda: {{ request.GET.q }}</span>{% endif %}
                        {% if request.GET.tipo_tarjeta %}<span class="badge bg-primary">Tipo: {{ request.GET.tipo_tarjeta }}</span>{% endif %}
                        {% if request.GET.categoria %}<span class="badge bg-primary">Categoría: {{ request.GET.categoria }}</span>{% endif %}
                        {% if request.GET.subcategoria %}<span class="badge bg-primary">Subcategoría: {{ request.GET.subcategoria }}</span>{% endif %}
                        {% if request.GET.zona %}<span class="badge bg-primary">Zona: {{ request.GET.zona }}</span>{% endif %}
                        {% if request.GET.cargo %}<span class="badge bg-primary">Cargo: {{ request.GET.cargo }}</span>{% endif %}
                        {% if request.GET.compania %}<span class="badge bg-primary">Compañía: {{ request.GET.compania }}</span>{% endif %}
                        {% if request.GET.area %}<span class="badge bg-primary">Área: {{ request.GET.area }}</span>{% endif %}
                        {% if request.GET.moneda %}<span class="badge bg-primary">Moneda: {{ request.GET.moneda }}</span>{% endif %}
                        {% if request.GET.año %}<span class="badge bg-primary">Año: {{ request.GET.año }}</span>{% endif %}
                        {% if request.GET.numero_tarjeta %}<span class="badge bg-primary">Número Tarjeta: {{ request.GET.numero_tarjeta }}</span>{% endif %}
                    </span>
                </div>
            {% endif %}
        </form>
    </div>
</div>

<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0" id="transactions-table">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__cargo&sort_direction={% if current_order == 'person__cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__compania&sort_direction={% if current_order == 'person__compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compañía
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__area&sort_direction={% if current_order == 'person__area' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Área
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=tipo_tarjeta&sort_direction={% if current_order == 'tipo_tarjeta' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tipo Tarjeta
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=numero_tarjeta&sort_direction={% if current_order == 'numero_tarjeta' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Número de Tarjeta
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=moneda&sort_direction={% if current_order == 'moneda' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Moneda
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=trm_cierre&sort_direction={% if current_order == 'trm_cierre' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                TRM Cierre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=valor_original&sort_direction={% if current_order == 'valor_original' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Valor Original
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=valor_cop&sort_direction={% if current_order == 'valor_cop' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Valor COP
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=numero_autorizacion&sort_direction={% if current_order == 'numero_autorizacion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Numero de Autorización
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=fecha_transaccion&sort_direction={% if current_order == 'fecha_transaccion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Fecha de Transacción
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=dia&sort_direction={% if current_order == 'dia' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Día
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=año&sort_direction={% if current_order == 'año' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Año
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=descripcion&sort_direction={% if current_order == 'descripcion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Descripción
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=categoria&sort_direction={% if current_order == 'categoria' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Categoría
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=subcategoria&sort_direction={% if current_order == 'subcategoria' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Subcategoría
                            </a>
                        </th>
                        <th id="zona-header">
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=zona&sort_direction={% if current_order == 'zona' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Zona
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=archivo&sort_direction={% if current_order == 'archivo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Archivo
                            </a>
                        </th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for transaction in page_obj %}
                        <tr>
                            <td>{{ transaction.person.nombre_completo }}</td>
                            <td>{{ transaction.person.cargo }}</td>
                            <td>{{ transaction.person.compania }}</td>
                            <td>{{ transaction.person.area }}</td>
                            <td>{{ transaction.tipo_tarjeta }}</td>
                            <td>{{ transaction.numero_tarjeta }}</td>
                            <td>{{ transaction.moneda }}</td>
                            <td>{{ transaction.trm_cierre }}</td>
                            <td>{{ transaction.valor_original }}</td>
                            <td>{{ transaction.valor_cop }}</td>
                            <td>{{ transaction.numero_autorizacion }}</td>
                            <td>{{ transaction.fecha_transaccion|date:"Y-m-d" }}</td>
                            <td>{{ transaction.dia }}</td>
                            <td>{{ transaction.año|floatformat:"0" }}</td>
                            <td>{{ transaction.descripcion }}</td>
                            <td>{{ transaction.categoria }}</td> 
                            <td>{{ transaction.subcategoria }}</td>
                            <td>{{ transaction.zona }}</td>
                            <td>{{ transaction.archivo }}</td>
                            <td class="table-fixed-column">
                                {% if transaction.person and transaction.person.cedula %}
                                    <a href="{% url 'person_details' transaction.person.cedula %}"
                                    class="btn btn-primary btn-sm"
                                    title="View person details">
                                        <i class="bi bi-person-vcard-fill"></i>
                                    </a>
                                {% else %}
                                    <span class="text-muted">No person data</span>
                                {% endif %}
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="19" class="text-center py-4" id="no-results-row">
                                {% if request.GET.q or request.GET.tipo_tarjeta or request.GET.categoria or request.GET.subcategoria or request.GET.zona or request.GET.cargo or request.GET.compania or request.GET.area or request.GET.moneda or request.GET.dia or request.GET.año or request.GET.numero_tarjeta %}
                                    Sin transacciones TC que coincidan con los filtros.
                                {% else %}
                                    Sin transacciones TC
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    {% if page_obj.has_other_pages %}
    <div class="card-footer bg-light">
        <nav>
            <ul class="pagination justify-content-center mb-0">
                {% if page_obj.has_previous %}
                    <li class="page-item">
                        <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                            <span aria-hidden="true">&laquo;&laquo;</span>
                        </a>
                    </li>
                    <li class="page-item">
                        <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                            <span aria-hidden="true">&laquo;</span>
                        </a>
                    </li>
                {% endif %}
                
                {% for num in page_obj.paginator.page_range %}
                    {% if page_obj.number == num %}
                        <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                    {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                        <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                    {% endif %}
                {% endfor %}
                
                {% if page_obj.has_next %}
                    <li class="page-item">
                        <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                            <span aria-hidden="true">&raquo;</span>
                        </a>
                    </li>
                    <li class="page-item">
                        <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                            <span aria-hidden="true">&raquo;&raquo;</span>
                        </a>
                    </li>
                {% endif %}
            </ul>
        </nav>
    </div>
    {% endif %}
</div>

<script>
// Client-side filtering script removed as filtering is now handled by the server.
</script>
{% endblock %}
"@ | Out-File -FilePath "core/templates/tcs.html" -Encoding utf8


# details template
@" 
{% extends "master.html" %}
{% load humanize %}
{% load my_filters %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
<a href="/" class="btn btn-custom-primary" title="Dashboard">
    <i class="fas fa-chart-pie"></i>
</a>
<a href="{% url 'person_list' %}" class="btn btn-custom-primary" title="Personas">
    <i class="fas fa-users"></i>
</a>
<a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Transacciones TC">
    <i class="far fa-credit-card"></i>
</a>
<a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
    <i class="fas fa-bell"></i>
</a>
<a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
    <i class="fas fa-database"></i> 
</a>
<form method="post" action="{% url 'logout' %}" class="d-inline">
    {% csrf_token %}
    <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
        <i class="fas fa-sign-out-alt"></i>
    </button>
</form>
{% endblock %}

{% block content %}
<div class="row">
    <div class="col-md-12 mb-4">
        <div class="card h-100"> {# Added h-100 for equal height #}
            <div class="card-header bg-light">
                <h5 class="mb-0">Informacion Personal</h5>
            </div>
            <div class="card-body">
                <table class="table">
                    <tr>
                        <th>ID:</th>
                        <td>{{ myperson.cedula|floatformat:"0" }}</td>
                    </tr>
                    <tr>
                        <th>Nombre:</th>
                        <td>{{ myperson.nombre_completo }}</td>
                    </tr>
                    <tr>
                        <th>Cargo:</th>
                        <td>{{ myperson.cargo }}</td>
                    </tr>
                    <tr>
                        <th>Area:</th>
                        <td>{{ myperson.area }}</td>
                    </tr>
                    <tr>
                        <th>Correo:</th>
                        <td>{{ myperson.correo }}</td>
                    </tr>
                    <tr>
                        <th>Compania:</th>
                        <td>{{ myperson.compania }}</td>
                    </tr>
                    <tr>
                        <th>Estado:</th>
                        <td>
                            <span class="badge bg-{% if myperson.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                {{ myperson.estado }}
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>Por revisar:</th>
                        <td>
                            <form action="{% url 'toggle_revisar_status' myperson.cedula %}" method="post" style="display:inline;">
                                {% csrf_token %}
                                <button type="submit"
                                        class="btn btn-link p-0 border-0 bg-transparent"
                                        title="{% if myperson.revisar %}Desmarcar para Revisar{% else %}Marcar para Revisar{% endif %}">
                                    <i class="fas fa-{% if myperson.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}"
                                    style="padding-left: 20px; font-size: 1.25rem;"></i>
                                </button>
                            </form>
                        </td>
                    </tr>
                    <tr>
                        <th>Comentarios:</th>
                        <td>
                            <div class="mb-3">
                                <strong>Historial:</strong><br>
                                <div id="comment-list" class="border p-2" style="max-height: 200px; overflow-y: auto;">
                                    {% if myperson.comments %}
                                        {% for comment_line in myperson.comments|split_lines %}
                                            {% if comment_line %}
                                                <div class="d-flex justify-content-between align-items-center mb-1">
                                                    <span>{{ comment_line }}</span>
                                                    <form action="{% url 'delete_comment' myperson.cedula forloop.counter0 %}" method="post" style="display:inline;">
                                                        {% csrf_token %}
                                                        <button type="submit" class="btn btn-danger btn-sm" title="Eliminar Comentario">
                                                            <i class="fas fa-trash-alt"></i>
                                                        </button>
                                                    </form>
                                                </div>
                                            {% endif %}
                                        {% endfor %}
                                    {% else %}
                                        <p class="text-muted">No hay comentarios existentes.</p>
                                    {% endif %}
                                </div>
                            </div>

                            <form action="{% url 'save_comment' myperson.cedula %}" method="post">
                                {% csrf_token %}
                                <div class="form-group mt-3">
                                    <label for="new_comment">Agregar nuevo comentario:</label>
                                    <textarea class="form-control" id="new_comment" name="new_comment" rows="3" placeholder="Escribe tu comentario..."></textarea>
                                </div>
                                <button type="submit" class="btn btn-primary btn-sm mt-2">Comentar</button>
                            </form>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
    </div>
</div>
<div class="row">
    <div class="col-md-6 mb-4">
        <div class="card h-100">
            <div class="card-header bg-light">
                <h5 class="mb-0">Resumen de Transacciones por Categoría y Año</h5>
            </div>
            <div class="card-body">
                {% if person_pivot_data %}
                    <div class="table-responsive">
                        <table class="table table-striped table-hover table-bordered text-center">
                            <thead class="table-light">
                                <tr>
                                    <th class="text-start">Categoría</th>
                                    {% for year in person_pivot_years %}
                                        <th colspan="2">{{ year|floatformat:"0" }}</th>
                                    {% endfor %}
                                    <th colspan="2">Total</th>
                                </tr>
                                <tr>
                                    <th></th>
                                    {% for year in person_pivot_years %}
                                        <th># Trans.</th>
                                        <th>Valor COP</th>
                                    {% endfor %}
                                    <th># Trans.</th>
                                    <th>Valor COP</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for category, year_data in person_pivot_data.items %}
                                <tr>
                                    <td class="text-start fw-bold">{{ category }}</td>
                                    {% for year in person_pivot_years %}
                                        <td>{{ year_data|get_item:year|get_item:'count'|intcomma }}</td>
                                        <td>`${{ year_data|get_item:year|get_item:'total_cop'|floatformat:2|intcomma }}</td>
                                    {% endfor %}
                                    <td class="fw-bold">{{ year_data.total.count|intcomma }}</td>
                                    <td class="fw-bold">`${{ year_data.total.total_cop|floatformat:2|intcomma }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                            <tfoot class="table-group-divider fw-bold">
                                <tr>
                                    <td class="text-start">Total General</td>
                                    {% for year in person_pivot_years %}
                                        <td>{{ person_pivot_column_totals|get_item:year|get_item:'count'|intcomma }}</td>
                                        <td>`${{ person_pivot_column_totals|get_item:year|get_item:'total_cop'|floatformat:2|intcomma }}</td>
                                    {% endfor %}
                                    <td>{{ person_pivot_grand_total.count|intcomma }}</td>
                                    <td>`${{ person_pivot_grand_total.total_cop|floatformat:2|intcomma }}</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                {% else %}
                    <p class="text-muted">No hay datos de transacciones para mostrar en el gráfico.</p>
                {% endif %}
            </div>
        </div>
    </div>
    <div class="col-md-6 mb-4">
        <div class="card h-100">
            <div class="card-header bg-light">
                <h5 class="mb-0">Transacciones por Año</h5>
            </div>
            <div class="card-body d-flex justify-content-center align-items-center">
                {% if yearly_chart_data and yearly_chart_data != '[]' %}
                    <canvas id="yearlyChart"></canvas>
                {% else %}
                    <p class="text-muted">No hay datos de transacciones para mostrar en el gráfico.</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>

<div class="row mt-4">
    <div class="col-md-12">
        <div class="card">
            <div class="card-header bg-light">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0">Transacciones TC</h5>
                    <span class="badge bg-success">{{ page_obj.paginator.count }} registros</span>
                </div>
            </div>
            <div class="card-body">
                <!-- Filter Form -->
                <form method="get" action="" class="row g-3 align-items-center mb-4" id="filter-form">
                    <div class="col-md-3">
                        <input type="text" name="q" class="form-control" placeholder="Buscar en transacciones..." value="{{ request.GET.q }}">
                    </div>
                    <div class="col-md-2">
                        <select class="form-select" name="tipo_tarjeta">
                            <option value="">Tipo Tarjeta</option>
                            {% for tipo in tipos_tarjeta %}
                                <option value="{{ tipo }}" {% if request.GET.tipo_tarjeta == tipo %}selected{% endif %}>{{ tipo }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-2">
                        <select class="form-select" name="categoria">
                            <option value="">Categoría</option>
                            {% for cat in categorias %}
                                <option value="{{ cat }}" {% if request.GET.categoria == cat %}selected{% endif %}>{{ cat }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-2">
                        <select class="form-select" name="subcategoria">
                            <option value="">Subcategoría</option>
                            {% for subcat in subcategorias %}
                                <option value="{{ subcat }}" {% if request.GET.subcategoria == subcat %}selected{% endif %}>{{ subcat }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-1">
                        <select class="form-select" name="zona">
                            <option value="">Zona</option>
                            {% for z in zonas %}
                                <option value="{{ z }}" {% if request.GET.zona == z %}selected{% endif %}>{{ z }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-1">
                        <select class="form-select" name="año">
                            <option value="">Año</option>
                            {% for a in años %}
                                <option value="{{ a }}" {% if request.GET.año == a|stringformat:"s" %}selected{% endif %}>{{ a }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-1 d-flex gap-2">
                        <button type="submit" class="btn btn-primary btn-sm" title="Aplicar filtros">
                            <i class="fas fa-filter"></i>
                        </button>
                        <a href="{% url 'person_details' myperson.cedula %}" class="btn btn-secondary btn-sm" title="Quitar filtros">
                            <i class="fas fa-undo"></i>
                        </a>
                    </div>
                     {% if request.GET.q or request.GET.tipo_tarjeta or request.GET.categoria or request.GET.subcategoria or request.GET.zona or request.GET.año %}
                        <div class="mt-2">
                            <span class="badge bg-info">
                                Filtros activos:
                                {% if request.GET.q %}<span class="badge bg-primary">Búsqueda: {{ request.GET.q }}</span>{% endif %}
                                {% if request.GET.tipo_tarjeta %}<span class="badge bg-primary">Tipo: {{ request.GET.tipo_tarjeta }}</span>{% endif %}
                                {% if request.GET.categoria %}<span class="badge bg-primary">Categoría: {{ request.GET.categoria }}</span>{% endif %}
                                {% if request.GET.subcategoria %}<span class="badge bg-primary">Subcategoría: {{ request.GET.subcategoria }}</span>{% endif %}
                                {% if request.GET.zona %}<span class="badge bg-primary">Zona: {{ request.GET.zona }}</span>{% endif %}
                                {% if request.GET.año %}<span class="badge bg-primary">Año: {{ request.GET.año }}</span>{% endif %}
                            </span>
                        </div>
                    {% endif %}
                </form>

                {% if credit_card_transactions %}
                    <div class="table-responsive">
                        <table class="table table-striped table-hover mb-0">
                            <thead>
                                <tr>
                                    <th>Fecha</th>
                                    <th>Categoría</th>
                                    <th>Subcategoría</th>
                                    <th>Tipo de Tarjeta</th>
                                    <th>Día</th>
                                    <th>Año</th>
                                    <th>Valor COP</th>
                                    <th>Descripción</th>
                                    <th>Zona</th>
                                    <th>No. Autorización</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for transaction in page_obj %}
                                <tr>
                                    <td>{{ transaction.fecha_transaccion|date:"d/m/Y" }}</td>
                                    <td>{{ transaction.categoria }}</td>
                                    <td>{{ transaction.subcategoria }}</td>
                                    <td>{{ transaction.tipo_tarjeta }}</td>
                                    <td>{{ transaction.dia }}</td>
                                    <td>{{ transaction.año|floatformat:"0" }}</td>
                                    <td>{{ transaction.valor_cop }}</td>
                                    <td>{{ transaction.descripcion }}</td>
                                    <td>{{ transaction.zona }}</td>
                                    <td>{{ transaction.numero_autorizacion }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                {% else %}
                    <p class="text-muted text-center">
                        {% if request.GET.q or request.GET.tipo_tarjeta or request.GET.categoria or request.GET.subcategoria or request.GET.zona or request.GET.año %}
                            No hay transacciones que coincidan con los filtros.
                        {% else %}
                            No hay transacciones de tarjeta de crédito disponibles para esta persona.
                        {% endif %}
                    </p>
                {% endif %}
            </div>
            {% if page_obj.has_other_pages %}
            <div class="card-footer bg-light">
                <nav>
                    <ul class="pagination justify-content-center mb-0">
                        {% if page_obj.has_previous %}
                            <li class="page-item"><a class="page-link" href="?page=1{% for key, value in all_params.items %}&{{ key }}={{ value }}{% endfor %}">&laquo;&laquo;</a></li>
                            <li class="page-item"><a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in all_params.items %}&{{ key }}={{ value }}{% endfor %}">&laquo;</a></li>
                        {% endif %}

                        {% for num in page_obj.paginator.page_range %}
                            {% if page_obj.number == num %}
                                <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                            {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                                <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in all_params.items %}&{{ key }}={{ value }}{% endfor %}">{{ num }}</a></li>
                            {% endif %}
                        {% endfor %}

                        {% if page_obj.has_next %}
                            <li class="page-item"><a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in all_params.items %}&{{ key }}={{ value }}{% endfor %}">&raquo;</a></li>
                            <li class="page-item"><a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in all_params.items %}&{{ key }}={{ value }}{% endfor %}">&raquo;&raquo;</a></li>
                        {% endif %}
                    </ul>
                </nav>
            </div>
            {% endif %}
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script>
document.addEventListener('DOMContentLoaded', function() {
    const yearlyChartData = {{ yearly_chart_data|safe }};
    if (yearlyChartData && yearlyChartData.length > 0) {
        const yearlyCtx = document.getElementById('yearlyChart').getContext('2d');
        const yearlyChart = new Chart(yearlyCtx, {
            type: 'line', // Line chart is great for time-series data
            data: {
                labels: {{ yearly_chart_labels|safe }},
                datasets: [{
                    label: 'Número de Transacciones',
                    data: yearlyChartData,
                    backgroundColor: 'rgba(255, 159, 64, 0.2)',
                    borderColor: 'rgba(255, 159, 64, 1)',
                    borderWidth: 2,
                    pointBackgroundColor: 'rgba(255, 159, 64, 1)',
                    pointRadius: 4,
                    tension: 0.1 // Makes the line slightly curved
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { 
                    legend: { display: false } 
                },
                scales: { 
                    y: { 
                        beginAtZero: true,
                        ticks: {
                            // Ensure only whole numbers are shown on the y-axis
                            stepSize: 1
                        }
                    } 
                }
            }
        });
    }
});
</script>

{% endblock %}
"@ | Out-File -FilePath "core/templates/details.html" -Encoding utf8

# Create alert template
@"
{% extends "master.html" %}
{% load static %}

{% block title %}Alertas{% endblock %}
{% block navbar_title %}Alertas{% endblock %}

{% block navbar_buttons %}
<div class="ms-auto d-flex align-items-center gap-2">
    <a href="/" class="btn btn-custom-primary" title="Dashboard">
        <i class="fas fa-chart-pie"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary" title="Personas">
    <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        <i class="fas fa-bell"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-database"></i> 
    </a>
    <a href="{% url 'export_persons_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="card mb-4 border-0 shadow">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <div class="d-flex align-items-center mb-3 col-12">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} alertas
                </span>
            </div>
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona o cédula" 
                       value="{{ request.GET.q }}">
            </div>
            
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-primary btn-lg flex-grow-1" title="Filtrar"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-secondary btn-lg flex-grow-1" title="Limpiar"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Cédula
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Correo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" class="link-dark text-decoration-none">
                                Estado
                            </a>
                        </th>
                        <th class="text-dark">Comentarios</th>
                        <th class="table-fixed-column">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr>
                            <td>
                                {# Replace the existing <a> tag with a form that submits to your toggle view #}
                                <form action="{% url 'toggle_revisar_status' person.cedula %}" method="post" style="display:inline;">
                                    {% csrf_token %}
                                    <button type="submit"
                                            class="btn btn-link p-0 border-0 bg-transparent" {# Style button to look like a clickable icon #}
                                            title="{% if person.revisar %}Desmarcar para Revisar{% else %}Marcar para Revisar{% endif %}">
                                        <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}"
                                           style="font-size: 1.25rem;"></i>
                                    </button>
                                </form>
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' person.cedula %}" 
                                   class="btn btn-sm btn-outline-primary"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="9" class="text-center py-4">
                                No hay personas marcadas para revisar.
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/alerts.html" -Encoding utf8

# Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'core.apps.CoreConfig',
    'django.contrib.humanize',"
    $settingsContent = $settingsContent -replace "from pathlib import Path", "from pathlib import Path
import os"
    $settingsContent | Set-Content -Path ".\arpa\settings.py"
    
    # Configure database to use SQLite
    $dbSettings = @"
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}
"@
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "(?s)DATABASES =.*?}.*?}", $dbSettings
    $settingsContent | Set-Content -Path ".\arpa\settings.py"


# Add static files configuration
Add-Content -Path ".\arpa\settings.py" -Value @"

# Static files (CSS, JavaScript, Images)
STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_DIRS = [
    BASE_DIR / "core/static",
]

MEDIA_URL = 'media/'
MEDIA_ROOT = BASE_DIR / 'media'

# Custom admin skin
ADMIN_SITE_HEADER = "A R P A"
ADMIN_SITE_TITLE = "ARPA Admin Portal"
ADMIN_INDEX_TITLE = "Bienvenido a A R P A"

LOGIN_REDIRECT_URL = '/'  
LOGOUT_REDIRECT_URL = '/accounts/login/'  
"@

    # Run migrations
    python3 manage.py makemigrations core
    python3 manage.py migrate

    python3 manage.py collectstatic --noinput

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python3 manage.py runserver

}

arpa