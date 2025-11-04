import pandas as pd

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        # Check if the value is "N/A" or empty, indicating no trend for the first year
        if value in ["N/A", "0.00%", ""]:
            return "" # Return empty string for no trend
            
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

def calculate_variation(df, column):
    """Calculate absolute and relative variations for a specific column."""
    df = df.sort_values(by=['Usuario', 'Año Declaración'])
    
    absolute_col = f'{column} Var. Abs.'
    relative_col = f'{column} Var. Rel.'
    
    df[absolute_col] = df.groupby('Usuario')[column].diff()
    
    # Calculate percentage change
    pct_change = df.groupby('Usuario')[column].pct_change(fill_method=None) * 100
    
    # Apply formatting: "0.00%" for non-NaN values, empty string for NaN (first year)
    df[relative_col] = pct_change.apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
    
    return df

def embed_trend_symbols(df, columns):
    """Add trend symbols to variation columns."""
    for col in columns:
        absolute_col = f'{col} Var. Abs.'
        relative_col = f'{col} Var. Rel.'
        
        if absolute_col in df.columns:
            df[absolute_col] = df.apply(
                lambda row: f"{row[absolute_col]:.2f} {get_trend_symbol(row[relative_col])}" 
                if pd.notna(row[f'{col} Var. Abs. No_Symbol']) else "N/A", # Use a temporary column to check for original NaN
                axis=1
            )
        
        if relative_col in df.columns:
            # Check if the underlying relative change was NaN before formatting
            df[relative_col] = df.apply(
                lambda row: f"{row[relative_col]} {get_trend_symbol(row[relative_col])}" if row[relative_col] != "" else "",
                axis=1
            )
    
    return df

def calculate_leverage(df):
    """Calculate financial leverage."""
    df['Apalancamiento'] = (df['Patrimonio'] / df['Activos']) * 100
    return df

def calculate_debt_level(df):
    """Calculate debt level."""
    df['Endeudamiento'] = (df['Pasivos'] / df['Activos']) * 100
    return df

def process_asset_data(df_assets):
    """Process asset data with variations and trends."""
    df_assets_grouped = df_assets.groupby(['Usuario', 'Año Declaración']).agg(
        Banco_Saldo=('Banco - Saldo COP', 'sum'),
        Bienes=('Total Bienes', 'sum'),
        Inversiones=('Total Inversiones', 'sum')
    ).reset_index()

    for column in ['Banco_Saldo', 'Bienes', 'Inversiones']:
        df_assets_grouped = calculate_variation(df_assets_grouped, column)
        # Create a temporary column to hold the absolute variation without symbol for the N/A check
        df_assets_grouped[f'{column} Var. Abs. No_Symbol'] = df_assets_grouped[f'{column} Var. Abs.']
    
    df_assets_grouped = embed_trend_symbols(df_assets_grouped, ['Banco_Saldo', 'Bienes', 'Inversiones'])
    return df_assets_grouped

def process_income_data(df_income):
    """Process income data with variations and trends."""
    df_income_grouped = df_income.groupby(['Usuario', 'Año Declaración']).agg(
        Ingresos=('Total Ingresos', 'sum'),
        Cant_Ingresos=('Cant_Ingresos', 'sum')
    ).reset_index()

    df_income_grouped = calculate_variation(df_income_grouped, 'Ingresos')
    df_income_grouped[f'Ingresos Var. Abs. No_Symbol'] = df_income_grouped[f'Ingresos Var. Abs.']
    df_income_grouped = embed_trend_symbols(df_income_grouped, ['Ingresos'])
    return df_income_grouped

def calculate_yearly_variations(df):
    """Calculate yearly variations for all columns."""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    columns_to_analyze = [
        'Activos', 'Pasivos', 'Patrimonio', 
        'Apalancamiento', 'Endeudamiento',
        'Banco_Saldo', 'Bienes', 'Inversiones', 'Ingresos',
        'Cant_Ingresos'
    ]
    
    # Store original values of absolute changes before formatting
    temp_abs_cols = {}
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        grouped = df.groupby('Usuario')[column]
        
        for year in [2021, 2022, 2023, 2024]:
            abs_col_name = f'{year} {column} Var. Abs.'
            rel_col_name = f'{year} {column} Var. Rel.'
            
            # Calculate absolute variation (diff)
            abs_variation = grouped.diff()
            df[abs_col_name] = abs_variation
            
            # Store original (unformatted) absolute variation for trend symbol logic
            temp_abs_cols[abs_col_name] = abs_variation
            
            # Calculate relative variation (pct_change)
            pct_change = grouped.pct_change(fill_method=None) * 100
            df[rel_col_name] = pct_change.apply(
                lambda x: f"{x:.2f}%" if pd.notna(x) else ""
            )
            
    # Apply formatting and symbols after all calculations
    for column in [col for col in columns_to_analyze if col in df.columns]:
        for year in [2021, 2022, 2023, 2024]:
            abs_col_name = f'{year} {column} Var. Abs.'
            rel_col_name = f'{year} {column} Var. Rel.'
            
            if abs_col_name in df.columns:
                df[abs_col_name] = df.apply(
                    lambda row: (
                        f"{temp_abs_cols[abs_col_name].loc[row.name]:.2f} {get_trend_symbol(row[rel_col_name])}" 
                        if pd.notna(temp_abs_cols[abs_col_name].loc[row.name]) else "N/A"
                    ),
                    axis=1
                )
            if rel_col_name in df.columns:
                df[rel_col_name] = df.apply(
                    lambda row: (
                        f"{row[rel_col_name]} {get_trend_symbol(row[rel_col_name])}" 
                        if row[rel_col_name] != "" else ""
                    ), 
                    axis=1
                )
    
    return df

def calculate_sudden_wealth_increase(df):
    """Calculate sudden wealth increase rate (Aum. Pat. Subito) as decimal with 1 decimal place"""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    # Calculate total wealth (Activo + Patrimonio)
    df['Capital'] = df['Activos'] + df['Patrimonio']
    
    # Calculate year-to-year change as decimal
    df['Aum. Pat. Subito_No_Symbol'] = df.groupby('Usuario')['Capital'].pct_change(fill_method=None)
    
    # Format as decimal (1 place) with trend symbol
    df['Aum. Pat. Subito'] = df['Aum. Pat. Subito_No_Symbol'].apply(
        lambda x: f"{x:.1f} {get_trend_symbol(f'{x*100:.1f}%')}" if pd.notna(x) else "N/A"
    )
    
    return df

def save_results(df, excel_filename="core/src/trends.xlsx"):
    """Save results to Excel with modified column names."""
    try:
        # Create a copy of the dataframe to avoid modifying the original
        df_output = df.copy()
        
        # Convert Usuario to string if it exists (before renaming)
        if 'Usuario' in df_output.columns:
            df_output['Usuario'] = df_output['Usuario'].astype(str)
        
        # Rename columns for output and drop temporary columns
        cols_to_drop = [col for col in df_output.columns if 'No_Symbol' in col]
        df_output = df_output.drop(columns=cols_to_drop, errors='ignore')

        df_output.columns = [col.replace('Usuario', 'Id').replace('Compañía', 'Compania') 
                           for col in df_output.columns]
        
        # Ensure Id is string after renaming
        if 'Id' in df_output.columns:
            df_output['Id'] = df_output['Id'].astype(str)
        
        df_output.to_excel(excel_filename, index=False)
        print(f"Data saved to {excel_filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    """Main function to process all data and generate analysis files."""
    try:
        # Process worth data
        df_worth = pd.read_excel("core/src/worthNets.xlsx")
        df_worth = df_worth.rename(columns={
            'Total Activos': 'Activos',
            'Total Pasivos': 'Pasivos',
            'Total Patrimonio': 'Patrimonio'
        })
        
        df_worth = calculate_leverage(df_worth)
        df_worth = calculate_debt_level(df_worth)
        df_worth = calculate_sudden_wealth_increase(df_worth)
        
        for column in ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento']:
            df_worth = calculate_variation(df_worth, column)
            # Create a temporary column to hold the absolute variation without symbol for the N/A check
            df_worth[f'{column} Var. Abs. No_Symbol'] = df_worth[f'{column} Var. Abs.']

        df_worth = embed_trend_symbols(df_worth, ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento'])
        
        # Process asset data
        df_assets = pd.read_excel("core/src/assetNets.xlsx")
        df_assets_processed = process_asset_data(df_assets)
        
        # Process income data
        df_income = pd.read_excel("core/src/incomeNets.xlsx")
        df_income_processed = process_income_data(df_income)
        
        # Merge all data
        df_combined = pd.merge(df_worth, df_assets_processed, on=['Usuario', 'Año Declaración'], how='left')
        df_combined = pd.merge(df_combined, df_income_processed, on=['Usuario', 'Año Declaración'], how='left')
        
        # Calculate yearly variations for the combined dataframe
        df_combined = calculate_yearly_variations(df_combined)

        # Save basic trends
        save_results(df_combined, "core/src/trends.xlsx")
        
    except FileNotFoundError as e:
        print(f"Error: Required file not found - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
