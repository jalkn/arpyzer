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
from core.models import Person, Conflict, FinancialReport, CreditCard
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
        return redirect('financial_report_list') # Or 'main' or 'alerts_list' as a default

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
        Overrides the default get_context_data to add counts for persons and conflicts,
        and to gather analysis results from the core/src directory.
        """
        context = super().get_context_data(**kwargs)
        # These counts are fetched from models directly
        context['conflict_count'] = Conflict.objects.count()
        context['person_count'] = Person.objects.count()
        context['finances_count'] = FinancialReport.objects.count()
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

        # --- Status for Personas.xlsx ---
        """
        personas_status = get_file_status('Personas.xlsx')
        analysis_results.append(personas_status)
        if personas_status['status'] == 'success':
            context['person_count'] = personas_status['records']

        # --- Status for conflicts.xlsx ---
        
        conflicts_status = get_file_status('conflicts.xlsx')
        analysis_results.append(conflicts_status)
        if conflicts_status['status'] == 'success':
            context['conflict_count'] = conflicts_status['records']
        """
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
        """
        analysis_results.append(get_file_status('bankNets.xlsx'))
        analysis_results.append(get_file_status('debtNets.xlsx'))
        analysis_results.append(get_file_status('goodNets.xlsx'))
        analysis_results.append(get_file_status('incomeNets.xlsx'))
        analysis_results.append(get_file_status('investNets.xlsx'))
        analysis_results.append(get_file_status('assetNets.xlsx'))
        analysis_results.append(get_file_status('worthNets.xlsx'))

        # --- Status for Trends.py output files ---
        analysis_results.append(get_file_status('trends.xlsx'))

        # --- Status for idTrends.xlsx ---
        idtrends_status = get_file_status('idTrends.xlsx')
        analysis_results.append(idtrends_status)
        if idtrends_status['status'] == 'success':
            context['financial_report_count'] = idtrends_status['records']
        """

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

            # Get unique values for filter dropdowns (normalize everything to strings)
            tipos_tarjeta = set()
            categorias = set()
            subcategorias = set()
            zonas = set()

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

            # Add the sets to context (sorted as strings)
            context['tipos_tarjeta'] = sorted(list(tipos_tarjeta))
            context['categorias'] = sorted(list(categorias))
            context['subcategorias'] = sorted(list(subcategorias))
            context['zonas'] = sorted(list(zonas))

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
                            transaction.get('descripcion', ''),
                            transaction.get('tipo_tarjeta', ''),
                            transaction.get('numero_tarjeta', ''),
                            transaction.get('fecha_transaccion', ''),
                            transaction.get('numero_autorizacion', ''),
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

                matches_tipo = (not tipo_tarjeta) or (trans_tipo and trans_tipo.lower() == tipo_tarjeta.lower())
                matches_categoria = (not categoria) or (trans_cat and trans_cat.lower() == categoria.lower())
                matches_subcategoria = (not subcategoria) or (trans_sub and trans_sub.lower() == subcategoria.lower())
                matches_zona = (not zona) or (trans_zona and trans_zona.lower() == zona.lower())

                # Add transaction if it matches all active filters
                if matches_global and matches_tipo and matches_categoria and matches_subcategoria and matches_zona:
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

@login_required
def main(request):
    """
    Main dashboard view. Gathers counts for various data types and passes them to the home template.
    """
    context = {
        'person_count': Person.objects.count(),
        'conflict_count': Conflict.objects.count(),
        'finances_count': FinancialReport.objects.count(),
        'alerts_count': Person.objects.filter(revisar=True).count(),
        # Corrected count for Accionista del Grupo to count distinct persons
        'accionista_grupo_count': Person.objects.filter(conflicts__q3=True).distinct().count(), # Changed from conflict__q3 to conflicts__q3
        # Count for Aum. Pat. Subito > 2, as seen in the original home.html
        'aum_pat_subito_alert_count': FinancialReport.objects.filter(aum_pat_subito__gt=2).count(),
        # New counts for declarations per year
        'declarations_2021_count': FinancialReport.objects.filter(ano_declaracion=2021).count(),
        'declarations_2022_count': FinancialReport.objects.filter(ano_declaracion=2022).count(),
        'declarations_2023_count': FinancialReport.objects.filter(ano_declaracion=2023).count(),
        'declarations_2024_count': FinancialReport.objects.filter(ano_declaracion=2024).count(),
        # Corrected counts for conflicts per year, based on the 'fecha_inicio' field from the Conflict model
        'conflicts_2021_count': Conflict.objects.filter(fecha_inicio__year=2021).count(),
        'conflicts_2022_count': Conflict.objects.filter(fecha_inicio__year=2022).count(),
        'conflicts_2023_count': Conflict.objects.filter(fecha_inicio__year=2023).count(),
        'conflicts_2024_count': Conflict.objects.filter(fecha_inicio__year=2024).count(),
        # Count for active persons
        'active_person_count': Person.objects.filter(estado='Activo').count(),
        # Count for retired persons
        'retired_person_count': Person.objects.filter(estado='Retirado').count(),
        'tc_count': CreditCard.objects.count(),
    }

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
def import_conflicts(request):
    """View for importing conflicts data from Excel files"""
    if request.method == 'POST' and request.FILES.get('conflict_excel_file'):
        excel_file = request.FILES['conflict_excel_file']
        try:
            dest_path = os.path.join(settings.BASE_DIR, "core", "src", "conflictos.xlsx")
            with open(dest_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            subprocess.run(['python3', 'core/conflicts.py'], check=True, cwd=settings.BASE_DIR)

            processed_file = os.path.join(settings.BASE_DIR, "core", "src", "conflicts.xlsx")
            df = pd.read_excel(processed_file)
            df.columns = df.columns.str.lower().str.replace(' ', '_')

            # Helper function to process boolean fields
            def get_boolean_value(value):
                if pd.isna(value):
                    return None  # Return None for NaN/empty values
                return bool(value) # Convert to boolean otherwise

            for _, row in df.iterrows():
                try:
                    person, created = Person.objects.get_or_create(
                        cedula=str(row['cedula']),
                        defaults={
                            'nombre_completo': row.get('nombre', ''),
                            'correo': row.get('email', ''),
                            'compania': row.get('compañía', ''),
                            'cargo': row.get('cargo', '')
                        }
                    )

                    fecha_inicio_str = row.get('fecha_de_inicio')
                    fecha_inicio_date = None
                    if pd.notna(fecha_inicio_str):
                        try:
                            fecha_inicio_date = pd.to_datetime(fecha_inicio_str).date()
                        except ValueError:
                            messages.warning(request, f"Could not parse date '{fecha_inicio_str}' for conflict. Skipping row.")
                            continue

                    Conflict.objects.update_or_create(
                        person=person,
                        fecha_inicio=fecha_inicio_date,
                        defaults={
                            'q1': get_boolean_value(row.get('q1')), # Changed
                            'q1_detalle': row.get('q1_detalle', ''),
                            'q2': get_boolean_value(row.get('q2')), # Changed
                            'q2_detalle': row.get('q2_detalle', ''),
                            'q3': get_boolean_value(row.get('q3')), # Changed
                            'q3_detalle': row.get('q3_detalle', ''),
                            'q4': get_boolean_value(row.get('q4')), # Changed
                            'q4_detalle': row.get('q4_detalle', ''),
                            'q5': get_boolean_value(row.get('q5')), # Changed
                            'q5_detalle': row.get('q5_detalle', ''),
                            'q6': get_boolean_value(row.get('q6')), # Changed
                            'q6_detalle': row.get('q6_detalle', ''),
                            'q7': get_boolean_value(row.get('q7')), # Changed
                            'q7_detalle': row.get('q7_detalle', ''),
                            'q8': get_boolean_value(row.get('q8')), # Changed
                            'q9': get_boolean_value(row.get('q9')), # Changed
                            'q10': get_boolean_value(row.get('q10')), # Changed
                            'q10_detalle': row.get('q10_detalle', ''),
                            'q11': get_boolean_value(row.get('q11')), # Changed
                            'q11_detalle': row.get('q11_detalle', '')
                        }
                    )

                except Exception as e:
                    messages.error(request, f"Error processing row with cedula {row.get('cedula', 'N/A')}: {str(e)}")
                    continue

            messages.success(request, f'Archivo de conflictos importado exitosamente! {len(df)} registros procesados.')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de conflictos: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
def import_financial_reports(request):
    """View for importing financial reports data from idTrends.xlsx"""
    # This function is called internally after idTrends.py generates the file
    # It does not expect a file upload directly from the user.
    try:
        file_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'idTrends.xlsx')
        if not os.path.exists(file_path):
            messages.error(request, "Error: idTrends.xlsx not found. Please ensure analysis scripts run first.")
            return

        df = pd.read_excel(file_path)
        # Ensure column names are consistently lowercased and spaces/dots are replaced
        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('.', '', regex=False).str.replace('á', 'a').str.replace('é', 'e').str.replace('í', 'i').str.replace('ó', 'o').str.replace('ú', 'u')

        # No need for column_mapping dictionary if direct access is used after cleaning column names
        for _, row in df.iterrows():
            try:
                cedula = str(row.get('cedula'))
                if not cedula:
                    messages.warning(request, f"Skipping row due to missing cedula: {row.to_dict()}")
                    continue # Skip rows without a cedula

                person, created = Person.objects.get_or_create(
                    cedula=cedula,
                    defaults={
                        'nombre_completo': row.get('nombre_completo', ''),
                        'correo': row.get('correo', ''),
                        'compania': row.get('compania_y', ''), # Use compania_y from idTrends.py output
                        'cargo': row.get('cargo', '')
                    }
                )

                apalancamiento_val = _clean_numeric_value(row.get('apalancamiento'))
                # If apalancamiento_val is a number > 1 and it didn't originally have a '%' sign,
                # it's likely a percentage like 12.45 which needs to be stored as 0.1245
                if isinstance(apalancamiento_val, (int, float)) and apalancamiento_val is not None and apalancamiento_val > 1.0 and '%' not in str(row.get('apalancamiento', '')):
                    apalancamiento_val /= 100

                endeudamiento_val = _clean_numeric_value(row.get('endeudamiento'))
                # Similar logic for endeudamiento_val
                if isinstance(endeudamiento_val, (int, float)) and endeudamiento_val is not None and endeudamiento_val > 1.0 and '%' not in str(row.get('endeudamiento', '')):
                    endeudamiento_val /= 100

                # Prepare data for FinancialReport, handling potential NaN values and cleaning numeric fields
                report_data = {
                    'person': person,
                    'fk_id_periodo': _clean_numeric_value(row.get('fkidperiodo')), # Corrected column name access
                    'ano_declaracion': _clean_numeric_value(row.get('año_declaracion')),
                    'ano_creacion': _clean_numeric_value(row.get('año_creacion')),
                    'activos': _clean_numeric_value(row.get('activos')),
                    'cant_bienes': _clean_numeric_value(row.get('cant_bienes')),
                    'cant_bancos': _clean_numeric_value(row.get('cant_bancos')),
                    'cant_cuentas': _clean_numeric_value(row.get('cant_cuentas')),
                    'cant_inversiones': _clean_numeric_value(row.get('cant_inversiones')),
                    'pasivos': _clean_numeric_value(row.get('pasivos')),
                    'cant_deudas': _clean_numeric_value(row.get('cant_deudas')),
                    'patrimonio': _clean_numeric_value(row.get('patrimonio')),
                    'apalancamiento': apalancamiento_val, # Use the processed value
                    'endeudamiento': endeudamiento_val,
                    'capital': _clean_numeric_value(row.get('capital')),
                    'aum_pat_subito': _clean_numeric_value(row.get('aum_pat_subito')), # Apply cleaning here
                    'activos_var_abs': _clean_numeric_value(row.get('activos_var_abs')),
                    'activos_var_rel': str(row.get('activos_var_rel')) if pd.notna(row.get('activos_var_rel')) else '',
                    'pasivos_var_abs': _clean_numeric_value(row.get('pasivos_var_abs')),
                    'pasivos_var_rel': str(row.get('pasivos_var_rel')) if pd.notna(row.get('pasivos_var_rel')) else '',
                    'patrimonio_var_abs': _clean_numeric_value(row.get('patrimonio_var_abs')),
                    'patrimonio_var_rel': str(row.get('patrimonio_var_rel')) if pd.notna(row.get('patrimonio_var_rel')) else '',
                    'apalancamiento_var_abs': _clean_numeric_value(row.get('apalancamiento_var_abs')),
                    'apalancamiento_var_rel': str(row.get('apalancamiento_var_rel')) if pd.notna(row.get('apalancamiento_var_rel')) else '',
                    'endeudamiento_var_abs': _clean_numeric_value(row.get('endeudamiento_var_abs')),
                    'endeudamiento_var_rel': str(row.get('endeudamiento_var_rel')) if pd.notna(row.get('endeudamiento_var_rel')) else '',
                    'banco_saldo': _clean_numeric_value(row.get('banco_saldo')),
                    'bienes': _clean_numeric_value(row.get('bienes')),
                    'inversiones': _clean_numeric_value(row.get('inversiones')),
                    'banco_saldo_var_abs': _clean_numeric_value(row.get('banco_saldo_var_abs')),
                    'banco_saldo_var_rel': str(row.get('banco_saldo_var_rel')) if pd.notna(row.get('banco_saldo_var_rel')) else '',
                    'bienes_var_abs': _clean_numeric_value(row.get('bienes_var_abs')),
                    'bienes_var_rel': str(row.get('bienes_var_rel')) if pd.notna(row.get('bienes_var_rel')) else '',
                    'inversiones_var_abs': _clean_numeric_value(row.get('inversiones_var_abs')),
                    'inversiones_var_rel': str(row.get('inversiones_var_rel')) if pd.notna(row.get('inversiones_var_rel')) else '',
                    'ingresos': _clean_numeric_value(row.get('ingresos')),
                    'cant_ingresos': _clean_numeric_value(row.get('cant_ingresos')),
                    'ingresos_var_abs': _clean_numeric_value(row.get('ingresos_var_abs')),
                    'ingresos_var_rel': str(row.get('ingresos_var_rel')) if pd.notna(row.get('ingresos_var_rel')) else '',
                }

                # Use update_or_create based on person and fk_id_periodo to ensure uniqueness per period
                # Ensure fk_id_periodo is not None for update_or_create to work correctly
                if report_data['fk_id_periodo'] is not None:
                    FinancialReport.objects.update_or_create(
                        person=person,
                        fk_id_periodo=report_data['fk_id_periodo'],
                        defaults=report_data
                    )
                else:
                    messages.warning(request, f"Skipping row for {person.nombre_completo} due to missing fk_id_periodo.")

            except Exception as e:
                messages.error(request, f"Error processing financial report for row: {row.to_dict()}. Error: {e}")
                
        messages.success(request, f"Se importaron exitosamente los datos de {len(df)} reportes financieros.")

    except FileNotFoundError:
        messages.error(request, "Error: idTrends.xlsx no se encontró. Asegúrese de que los scripts de análisis se hayan ejecutado correctamente.")
    except Exception as e:
        messages.error(request, f"Ocurrió un error al procesar el archivo idTrends.xlsx: {e}")

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
    i = 0
    while f'column_{i}' in request.GET:
        column = request.GET.get(f'column_{i}')
        operator = request.GET.get(f'operator_{i}')
        value1 = request.GET.get(f'value_{i}')
        value2 = request.GET.get(f'value2_{i}')

        if column and operator and value1:
            # Corrected: Use 'financial_reports' as the related name from Person to FinancialReport
            filter_key = f'financial_reports__{column}'

            try:
                # Remove commas from value1 and value2 before conversion
                if isinstance(value1, str):
                    value1 = value1.replace(',', '')
                if isinstance(value2, str):
                    value2 = value2.replace(',', '')

                # Convert value1 to appropriate type based on common financial fields
                if column in ['fk_id_periodo', 'ano_declaracion', 'cant_bienes', 'cant_bancos', 'cant_cuentas',
                              'cant_inversiones', 'cant_deudas', 'cant_ingresos']:
                    value1 = int(float(value1)) # Convert to int if it's a count/ID
                    if value2: value2 = int(float(value2))
                else: # Assume float for monetary values and percentages
                    value1 = float(value1)
                    if value2: value2 = float(value2)
            except (ValueError, TypeError):
                # Handle cases where conversion fails (e.g., non-numeric input for numeric fields)
                # You might want to log this or provide user feedback
                value1 = None # Invalidate the filter if value is not convertible
                value2 = None

            if value1 is not None:
                if operator == '>':
                    persons = persons.filter(**{f'{filter_key}__gt': value1})
                elif operator == '<':
                    persons = persons.filter(**{f'{filter_key}__lt': value1})
                elif operator == '=':
                    persons = persons.filter(**{f'{filter_key}': value1})
                elif operator == '>=':
                    persons = persons.filter(**{f'{filter_key}__gte': value1})
                elif operator == '<=':
                    persons = persons.filter(**{f'{filter_key}__lte': value1})
                elif operator == 'between' and value2 is not None:
                    persons = persons.filter(**{f'{filter_key}__range': (min(value1, value2), max(value1, value2))})
                elif operator == 'contains':
                    # 'contains' operator is typically for text fields.
                    # Ensure the column is a text field or handle accordingly.
                    # For numeric fields, 'contains' usually doesn't make sense.
                    persons = persons.filter(**{f'{filter_key}__icontains': str(value1)})
        i += 1


    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by).distinct() # Use .distinct() to avoid duplicate persons if related objects cause issues

    # Prepare data for DataFrame
    data = []
    for person in persons:
        # You might need to adjust this logic based on how you want to handle multiple reports per person
        financial_report = FinancialReport.objects.filter(person=person).order_by('-ano_declaracion', '-fk_id_periodo').first()

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

        # Add financial report data if available
        if financial_report:
            row_data.update({
                'Periodo': financial_report.fk_id_periodo,
                'Ano': financial_report.ano_declaracion,
                'Aum. Pat. Subito': financial_report.aum_pat_subito,
                '% Endeudamiento': financial_report.endeudamiento,
                'Patrimonio': financial_report.patrimonio,
                'Patrimonio Var. Rel. %': financial_report.patrimonio_var_rel,
                'Patrimonio Var. Abs. $': financial_report.patrimonio_var_abs,
                'Activos': financial_report.activos,
                'Activos Var. Rel. %': financial_report.activos_var_rel,
                'Activos Var. Abs. $': financial_report.activos_var_abs,
                'Pasivos': financial_report.pasivos,
                'Pasivos Var. Rel. %': financial_report.pasivos_var_rel,
                'Pasivos Var. Abs. $': financial_report.pasivos_var_abs,
                'Cant. Deudas': financial_report.cant_deudas,
                'Ingresos': financial_report.ingresos,
                'Ingresos Var. Rel. %': financial_report.ingresos_var_rel,
                'Ingresos Var. Abs. $': financial_report.ingresos_var_abs,
                'Cant. Ingresos': financial_report.cant_ingresos,
                'Bancos Saldo': financial_report.banco_saldo,
                'Bancos Var. Rel. %': financial_report.banco_saldo_var_rel,
                'Bancos Var. $': financial_report.banco_saldo_var_abs,
                'Cant. Cuentas': financial_report.cant_cuentas,
                'Cant. Bancos': financial_report.cant_bancos,
                'Bienes Valor': financial_report.bienes,
                'Bienes Var. Rel. %': financial_report.bienes_var_rel,
                'Bienes Var. $': financial_report.bienes_var_abs,
                'Cant. Bienes': financial_report.cant_bienes,
                'Inversiones Valor': financial_report.inversiones,
                'Inversiones Var. Rel. %': financial_report.inversiones_var_rel,
                'Inversiones Var. $': financial_report.inversiones_var_abs,
                'Cant. Inversiones': financial_report.cant_inversiones,
            })
        data.append(row_data)

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

@login_required
def conflict_list(request):
    search_query = request.GET.get('q', '')
    compania_filter = request.GET.get('compania', '')
    column_filter = request.GET.get('column', '')
    answer_filter = request.GET.get('answer', '')
    missing_details_view = request.GET.get('missing_details', False) # New parameter for missing details

    order_by = request.GET.get('order_by', 'person__nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    conflicts = Conflict.objects.select_related('person').all()

    if search_query:
        conflicts = conflicts.filter(
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query)
        )

    if compania_filter:
        conflicts = conflicts.filter(person__compania=compania_filter)

    if column_filter and answer_filter:
        filter_q = Q()
        if answer_filter == 'yes':
            filter_q = Q(**{column_filter: True})
        elif answer_filter == 'no':
            filter_q = Q(**{column_filter: False})
        elif answer_filter == 'blank': # Filter for blank answers
            filter_q = Q(**{column_filter: None})
        conflicts = conflicts.filter(filter_q)

    # Filtering for conflicts where qX is True but qX_detalle is blank (None or empty string)
    if missing_details_view == 'true': # Check if the parameter is explicitly 'true'
        missing_details_q = Q()
        for i in range(1, 12):
            q_field = f'q{i}'
            detail_field = f'q{i}_detalle'
            # Only apply this if the q_field is True AND the detail_field is either None or an empty string
            missing_details_q |= Q(**{q_field: True, detail_field + '__isnull': True}) | Q(**{q_field: True, detail_field: ''})
        conflicts = conflicts.filter(missing_details_q)

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    conflicts = conflicts.order_by(order_by)

    # Attach the '_detalle' fields as '_answer' for template display
    for conflict in conflicts:
        for i in range(1, 12):
            detail_field_name = f'q{i}_detalle'
            answer_field_name = f'q{i}_answer'
            # Use getattr to safely access the attribute, with a default of None if it doesn't exist
            setattr(conflict, answer_field_name, getattr(conflict, detail_field_name, None))


    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')

    paginator = Paginator(conflicts, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Build all_params for pagination links
    all_params = {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']}

    context = {
        'conflicts': page_obj,
        'page_obj': page_obj,
        'companias': companias,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': all_params,
        'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
        'missing_details_view': missing_details_view, # Pass the boolean to the template
    }
    return render(request, 'conflicts.html', context)

@login_required
def financial_report_list(request):
    # Initialize a Q object to accumulate all filters
    all_filters_q = Q()

    # Handle the main search query 'q' (for person's name or cedula)
    search_query = request.GET.get('q', '')
    if search_query:
        all_filters_q &= (
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query)
        )

    # Handle existing 'compania' filter
    compania_filter = request.GET.get('compania', '')
    if compania_filter:
        all_filters_q &= Q(person__compania=compania_filter)

    # Handle existing 'ano_declaracion' filter
    ano_declaracion_filter = request.GET.get('ano_declaracion', '')
    if ano_declaracion_filter:
        try:
            # Ensure it's an integer for exact match
            ano_declaracion_int = int(ano_declaracion_filter)
            all_filters_q &= Q(ano_declaracion=ano_declaracion_int)
        except ValueError:
            messages.warning(request, "Año de declaración inválido.")

    # Iterate through potential filter indices (e.g., column_0, column_1, etc.)
    i = 0
    while True:
        # The names in the GET request will be like column_0, operator_0, value_0, value2_0
        column = request.GET.get(f'column_{i}')
        operator = request.GET.get(f'operator_{i}')
        value1 = request.GET.get(f'value_{i}')
        value2 = request.GET.get(f'value2_{i}') # For 'between' operator

        # If no column is found for the current index, stop iterating
        if not column:
            break

        # Only apply a filter if column, operator, and at least value1 are present
        if column and operator and value1:
            try:
                if operator == '>':
                    all_filters_q &= Q(**{f"{column}__gt": _clean_numeric_value(value1)})
                elif operator == '<':
                    all_filters_q &= Q(**{f"{column}__lt": _clean_numeric_value(value1)})
                elif operator == '=':
                    # For exact match, for numeric fields use _clean_numeric_value
                    # For text fields, __iexact is often better than just =
                    cleaned_value = _clean_numeric_value(value1)
                    if cleaned_value is not None: # It's a number
                        all_filters_q &= Q(**{f"{column}": cleaned_value})
                    else: # Treat as text
                        all_filters_q &= Q(**{f"{column}__iexact": value1})
                elif operator == '>=':
                    all_filters_q &= Q(**{f"{column}__gte": _clean_numeric_value(value1)})
                elif operator == '<=':
                    all_filters_q &= Q(**{f"{column}__lte": _clean_numeric_value(value1)})
                elif operator == 'between' and value2:
                    val1_cleaned = _clean_numeric_value(value1)
                    val2_cleaned = _clean_numeric_value(value2)
                    if val1_cleaned is not None and val2_cleaned is not None:
                        # Ensure min/max for correct range
                        all_filters_q &= Q(**{f"{column}__range": (min(val1_cleaned, val2_cleaned), max(val1_cleaned, val2_cleaned))})
                    else:
                        messages.warning(request, f"Valores inválidos para el filtro 'entre' en columna {column}.")
                elif operator == 'contains':
                    # 'contains' is typically for text fields. Use icontains for case-insensitivity.
                    all_filters_q &= Q(**{f"{column}__icontains": value1})
                else:
                    messages.warning(request, f"Operador inválido '{operator}' para la columna {column}.")
            except ValueError:
                messages.error(request, f"Error al convertir valor para el filtro en {column}. Verifique el formato numérico.")
            except Exception as e:
                messages.error(request, f"Error inesperado al aplicar filtro en {column}: {e}")
        
        i += 1 # Move to the next potential filter index

    # Apply all accumulated filters to the queryset
    financial_reports = FinancialReport.objects.select_related('person').filter(all_filters_q)

    # Ordering logic
    order_by = request.GET.get('order_by', 'person__nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    
    financial_reports = financial_reports.order_by(order_by)

    # Get distinct values for existing filters (Companias and Anos Declaracion)
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')
    anos_declaracion = FinancialReport.objects.exclude(ano_declaracion__isnull=True).values_list('ano_declaracion', flat=True).distinct().order_by('ano_declaracion')

    # Pagination
    paginator = Paginator(financial_reports, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Prepare all_params for pagination links to persist all GET parameters
    all_params = request.GET.copy()
    if 'page' in all_params:
        del all_params['page']
    
    # Alert count
    alerts_count = Person.objects.filter(revisar=True).count()

    context = {
        'financial_reports': page_obj,
        'page_obj': page_obj,
        'companias': companias,
        'anos_declaracion': anos_declaracion,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': all_params, # This will now correctly include all column_X, operator_X, value_X params
        'alerts_count': alerts_count,
    }

    return render(request, 'finances.html', context)

@login_required
def import_finances(request):
    """View for importing protected Excel files and running analysis"""
    if request.method == 'POST' and request.FILES.get('finances_file'):
        excel_file = request.FILES['finances_file']
        password = request.POST.get('excel_password', '')

        try:
            # Save the original file temporarily
            temp_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'temp_protected.xlsx')
            with open(temp_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            # Try to decrypt the file if password is provided
            decrypted_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'data.xlsx')

            if password:
                try:
                    with open(temp_path, 'rb') as f:
                        file = msoffcrypto.OfficeFile(f)
                        file.load_key(password=password)
                        decrypted = io.BytesIO()
                        file.decrypt(decrypted)

                        with open(decrypted_path, 'wb') as out:
                            out.write(decrypted.getvalue())
                except Exception as e:
                    messages.error(request, f'Error al desproteger el archivo: {str(e)}')
                    return HttpResponseRedirect('/import/')
            else:
                # If no password, just copy the file
                import shutil
                shutil.copyfile(temp_path, decrypted_path)

            # Remove the temporary file
            os.remove(temp_path)

            # Run the analysis scripts in sequence
            try:
                # Run cats.py analysis
                subprocess.run(['python3', 'core/cats.py'], check=True, cwd=settings.BASE_DIR)

                # Run nets.py analysis
                subprocess.run(['python3', 'core/nets.py'], check=True, cwd=settings.BASE_DIR)

                # Run trends.py analysis
                subprocess.run(['python3', 'core/trends.py'], check=True, cwd=settings.BASE_DIR)

                # Run idTrends.py analysis
                subprocess.run(['python3', 'core/idTrends.py'], check=True, cwd=settings.BASE_DIR)

                # After idTrends.py generates idTrends.xlsx, import the data into the FinancialReport model
                import_financial_reports(request) # Call the new import function

                # Remove the data.xlsx file after processing
                os.remove(decrypted_path)

                messages.success(request, 'Archivo procesado exitosamente y análisis completado!')
            except subprocess.CalledProcessError as e:
                messages.error(request, f'Error ejecutando análisis: {str(e)}')
            except Exception as e:
                messages.error(request, f'Error durante el análisis: {str(e)}')

        except Exception as e:
            messages.error(request, f'Error procesando archivo protegido: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
def person_details(request, cedula):
    """
    View to display the details of a specific person, including related financial reports, conflicts, and credit card transactions.
    """
    myperson = get_object_or_404(Person, cedula=cedula)
    financial_reports = FinancialReport.objects.filter(person=myperson).order_by('-ano_declaracion')
    conflicts = myperson.conflicts.all().order_by('-fecha_inicio')
    
    # Retrieve all CreditCard objects associated with the person
    credit_card_transactions = CreditCard.objects.filter(person=myperson)
    
    context = {
        'myperson': myperson,
        'financial_reports': financial_reports,
        'conflicts': conflicts,
        'alerts_count': Person.objects.filter(revisar=True).count(),
        'credit_card_transactions': credit_card_transactions, # Add the credit card transactions to the context
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
        pdf_files = request.FILES.getlist('visa_pdf_files')
        pdf_password = request.POST.get('visa_pdf_password', '')

        if not pdf_files:
            messages.error(request, 'No se seleccionaron archivos PDF.', extra_tags='import_tcs')
            return redirect('import') # Updated from 'import_page'

        input_pdf_dir = os.path.join(settings.BASE_DIR, 'core', 'src', 'extractos')
        output_excel_dir = os.path.join(settings.BASE_DIR, 'core', 'src')
        tcs_excel_path = os.path.join(output_excel_dir, "tcs.xlsx") # Path to the output Excel

        os.makedirs(input_pdf_dir, exist_ok=True) # Ensure input directory exists

        # Clear existing PDFs in the input_pdf_dir before saving new ones
        for filename in os.listdir(input_pdf_dir):
            if filename.endswith(".pdf"):
                os.remove(os.path.join(input_pdf_dir, filename))

        files_saved = 0
        for pdf_file in pdf_files:
            file_path = os.path.join(input_pdf_dir, pdf_file.name)
            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)
                files_saved += 1
            except Exception as e:
                messages.error(request, f"Error saving PDF '{pdf_file.name}': {e}", extra_tags='import_tcs')

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

    return redirect('import') # Updated from 'import_page'


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
def export_financial_reports_excel(request):
    """
    Exports the filtered list of FinancialReport objects to an Excel file.
    """
    
    # --- Start Filter Logic (Use the same logic as financial_report_list) ---
    queryset = FinancialReport.objects.all()
    # Apply filtering based on request.GET if necessary...
    # --- End Filter Logic ---

    # 1. Convert the QuerySet to a Pandas DataFrame
    # Use list(queryset.values()) to get a list of dictionaries with field values
    data = list(queryset.values()) 
    df = pd.DataFrame(data)

    # 2. ⚡️ FIX: Remove Timezone Information from Datetime Columns ⚡️
    # Iterate over columns and convert timezone-aware columns to timezone-naive
    for col in df.columns:
        # Check if the column is of datetime type and if it has timezone information
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            # Use 'dt.tz_localize(None)' to remove the timezone information
            # We assume the data is already in the desired local time zone due to Django settings.
            df[col] = df[col].dt.tz_localize(None)

    # 3. Prepare Excel workbook and response headers
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    # Set filename
    response['Content-Disposition'] = 'attachment; filename=reporte_bienes_y_rentas.xlsx'

    # 4. Create a workbook and sheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Bienes y Rentas"

    # 5. Write the DataFrame to the worksheet
    for row in dataframe_to_rows(df, header=True, index=False):
        worksheet.append(row)
    
    # 6. Save the workbook to the response
    workbook.save(response)
    
    return response

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
