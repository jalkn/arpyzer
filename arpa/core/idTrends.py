import pandas as pd
import re
import numpy as np

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"
        elif value_float > 0.1:  # more than 10% increase
            return "📈"
        elif value_float < -0.1:  # more than 10% decrease
            return "📉"
        else:
            return "➡️"  # relatively stable
    except Exception:
        return "➡️"

def clean_and_convert(value, keep_trend=False):
    """Clean and convert to float, optionally preserving trend symbol."""
    if pd.isna(value):
        return value
    
    str_value = str(value)
    
    # Handle "N/A ➡️" case specifically
    if "N/A ➡️" in str_value:
        return np.nan
    
    if keep_trend:
        # Extract numeric part (including percentages)
        numeric_part = re.sub(r'[^\d.%\-]', '', str_value)
        try:
            numeric_value = float(numeric_part.strip('%')) / 100 if '%' in numeric_part else float(numeric_part)
            trend_symbol = get_trend_symbol(str_value)
            return f"{numeric_value:.2%}"[:-1] + trend_symbol  # Format as percentage without % and add symbol
        except:
            return None
    else:
        # For absolute values, just clean numbers
        cleaned = re.sub(r'[^\d.-]', '', str_value)
        try:
            return float(cleaned) if cleaned else None
        except:
            return None

def remove_trend_symbol(value):
    """Remove the trend symbol from a string value."""
    if pd.isna(value):
        return value
    str_value = str(value)
    # Remove any known trend symbols
    cleaned_value = str_value.replace("📈", "").replace("📉", "").replace("➡️", "").strip()
    return cleaned_value


# Read the Excel file
file_path_trends = 'core/src/trends.xlsx'
df_trends = pd.read_excel(file_path_trends)

# Ensure all specified columns exist (create empty ones if they don't)
required_columns = [
    'Id', 'Nombre', 'Compania', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 
    'Año Creación', 'Activos', 'Cant_Bienes', 'Cant_Bancos', 'Cant_Cuentas', 
    'Cant_Inversiones', 'Pasivos', 'Cant_Deudas', 'Patrimonio', 'Apalancamiento', 
    'Endeudamiento', 'Capital', 'Aum. Pat. Subito', 'Activos Var. Abs.', 
    'Activos Var. Rel.', 'Pasivos Var. Abs.', 'Pasivos Var. Rel.', 
    'Patrimonio Var. Abs.', 'Patrimonio Var. Rel.', 'Apalancamiento Var. Abs.', 
    'Apalancamiento Var. Rel.', 'Endeudamiento Var. Abs.', 'Endeudamiento Var. Rel.', 
    'Banco_Saldo', 'Bienes', 'Inversiones', 'Banco_Saldo Var. Abs.', 
    'Banco_Saldo Var. Rel.', 'Bienes Var. Abs.', 'Bienes Var. Rel.', 
    'Inversiones Var. Abs.', 'Inversiones Var. Rel.', 'Ingresos', 
    'Cant_Ingresos', 'Ingresos Var. Abs.', 'Ingresos Var. Rel.'
]

# Add any missing columns with NaN values
for col in required_columns:
    if col not in df_trends.columns:
        df_trends[col] = None

# List of columns to convert to float (absolute variation columns)
float_columns = [
    'Activos Var. Abs.', 
    'Pasivos Var. Abs.', 
    'Patrimonio Var. Abs.', 
    'Apalancamiento Var. Abs.', 
    'Endeudamiento Var. Abs.',  
    'Banco_Saldo Var. Abs.', 
    'Bienes Var. Abs.', 
    'Inversiones Var. Abs.', 
    'Ingresos Var. Abs.'
]

# List of columns to clean infinity values and keep trend symbols
trend_columns = [
    'Apalancamiento', 
    'Endeudamiento', 
    'Activos Var. Rel.', 
    'Pasivos Var. Rel.', 
    'Patrimonio Var. Rel.', 
    'Apalancamiento Var. Rel.', 
    'Endeudamiento Var. Rel.', 
    'Banco_Saldo Var. Rel.', 
    'Bienes Var. Rel.', 
    'Inversiones Var. Rel.', 
    'Ingresos Var. Rel.'
]

# Convert absolute variation columns to float
for col in float_columns:
    if col in df_trends.columns:
        df_trends[col] = df_trends[col].apply(lambda x: clean_and_convert(x, keep_trend=False))

# Process trend columns (handle infinity and preserve trend symbols)
for col in trend_columns:
    if col in df_trends.columns:
        df_trends[col] = df_trends[col].apply(lambda x: clean_and_convert(x, keep_trend=True) 
                          if not pd.isna(x) and str(x).lower() not in ['inf', '-inf', 'inf%'] 
                          else np.nan)

# Special handling for 'Aum. Pat. Subito' column
if 'Aum. Pat. Subito' in df_trends.columns:
    df_trends['Aum. Pat. Subito'] = df_trends['Aum. Pat. Subito'].apply(
        lambda x: np.nan if pd.isna(x) or "N/A ➡️" in str(x) else x
    )

# Reorder columns to match the specified order
df_trends = df_trends[required_columns]

# Read the Personas.xlsx file
file_path_personas = 'core/src/Personas.xlsx'
try:
    df_personas = pd.read_excel(file_path_personas)
except FileNotFoundError:
    print(f"Error: {file_path_personas} not found. Please ensure the file exists.")
    exit()

# You can change 'how' to 'inner', 'right', or 'outer' depending on your desired merge behavior
df_merged = pd.merge(df_trends, df_personas, on='Id', how='left')

# Fill null values in 'Cant_Ingresos' with 0
if 'Cant_Ingresos' in df_merged.columns:
    df_merged['Cant_Ingresos'] = df_merged['Cant_Ingresos'].fillna(0)

# Fill null values in 'Ingresos' with 0
if 'Ingresos' in df_merged.columns:
    df_merged['Ingresos'] = df_merged['Ingresos'].fillna(0)

# Define columns to remove from the output
columns_to_remove = ["Id", "Nombre", "Cargo", "Compania_x", "correo"]

# Drop the specified columns
df_merged = df_merged.drop(columns=columns_to_remove, errors='ignore') # 'errors=ignore' prevents an error if a column isn't found

# Define the desired order of columns
desired_start_columns = ["Cedula", "NOMBRE COMPLETO", "Estado", "Compania_y", "CARGO"]

# Then, get all other columns that are not in the desired_start_columns
remaining_columns = [col for col in df_merged.columns if col not in desired_start_columns]

# Concatenate the two lists to form the final column order
final_column_order = desired_start_columns + remaining_columns

# Reindex the DataFrame with the new column order
df_merged = df_merged[final_column_order]

# Save the modified and merged dataframe back to Excel (idTrends.xlsx)
output_path_idtrends = 'core/src/trendSym.xlsx'
df_merged.to_excel(output_path_idtrends, index=False)
print(f"File has been modified and saved as {output_path_idtrends}")

# --- New section for idTrends.xlsx (without trend symbols) ---
df_idTrends = df_merged.copy() # Create a copy to modify without affecting trendSym.xlsx

# List of columns from which to remove trend symbols (these are the 'Rel.' columns and Apalancamiento/Endeudamiento)
columns_to_clean_symbols = [
    'Apalancamiento', 
    'Endeudamiento', 
    'Activos Var. Rel.', 
    'Pasivos Var. Rel.', 
    'Patrimonio Var. Rel.', 
    'Apalancamiento Var. Rel.', 
    'Endeudamiento Var. Rel.', 
    'Banco_Saldo Var. Rel.', 
    'Bienes Var. Rel.', 
    'Inversiones Var. Rel.', 
    'Ingresos Var. Rel.'
]

for col in columns_to_clean_symbols:
    if col in df_idTrends.columns:
        df_idTrends[col] = df_idTrends[col].apply(remove_trend_symbol)

# Save the dataframe without trend symbols to idTrends.xlsx
output_path_idtrends = 'core/src/idTrends.xlsx'
df_idTrends.to_excel(output_path_idtrends, index=False)
print(f"File without trend symbols has been saved as {output_path_idtrends}")
