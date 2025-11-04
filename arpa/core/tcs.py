import os
import re
import fitz
import pdfplumber
import pandas as pd
from datetime import datetime

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
            "Valor COP", "Número de Autorización", "Fecha de Transacción", "Día", 
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
