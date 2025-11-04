import PyPDF2
import pandas as pd
import re
import datetime
import locale
from PyPDF2.errors import PdfReadError

def extract_and_parse_data(pdf_path):
    """
    Extracts and parses transaction data from a PDF file, grouped by card.

    Args:
        pdf_path (str): The file path to the PDF.

    Returns:
        list: A list of lists, where each inner list is a row of parsed data
              with card details included.
    """
    parsed_rows = []

    try:
        # Set the locale to Spanish for date formatting
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except locale.Error:
            print("Warning: Spanish locale not found. Falling back to default.")
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

                # Find all card blocks in the text of the current page
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
                            r'(\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2})\s+'
                            r'(.*?)\s+'
                            r'(\d{6})\s*'
                            r'(\True[\d\.,]+)\s*(?:([\d\.,]+)?\s*(\bUSD\b|\bEUR\b|\bPEN\b)?)?',
                            re.IGNORECASE | re.MULTILINE
                        )
                        
                        # Find all transactions within this specific card block
                        transactions = transaction_line_pattern.findall(block_text)
                        
                        if not transactions:
                            print(f"No transactions found for {cardholder_name} ({card_number}) on page {page_number}.")
                        
                        for transaction in transactions:
                            date = transaction[0].strip()
                            description = transaction[1].strip()
                            auth_num = transaction[2].strip()
                            
                            primary_value = transaction[3].strip()
                            secondary_value = transaction[4].strip()
                            moneda = transaction[5].strip() if transaction[5] else "COP"

                            valor_cop = primary_value

                            if secondary_value:
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
                                date,
                                day_of_week,
                                year,
                                description,
                                auth_num,
                                valor_original,
                                moneda,
                                valor_cop,
                                formatted_card_type,
                                card_number,
                                cardholder_name,
                                pdf_path,
                                page_number
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
    Reads the Personas.xlsx file into a DataFrame.
    """
    try:
        personas_df = pd.read_excel(excel_path)
        return personas_df
    except FileNotFoundError:
        print(f"Error: The file '{excel_path}' was not found.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred while reading '{excel_path}': {e}")
        return pd.DataFrame()

if __name__ == "__main__":
    pdf_file_name = "clara.pdf"
    excel_file_name = "Clara.xlsx"
    personas_file_name = "Personas.xlsx"
    
    column_headers = ["Fecha Transacción", "Día", "Año", "Descripción", "Número de Autorización", "Valor Original", "Moneda", "Valor COP", "Tipo de Tarjeta", "Número de Tarjeta", "Tarjetahabiente", "Archivo", "Página"]

    extracted_data = extract_and_parse_data(pdf_file_name)

    if extracted_data:
        df_clara = pd.DataFrame(extracted_data, columns=column_headers)

        df_personas = read_personas_excel(personas_file_name)

        if not df_personas.empty:
            final_df = pd.merge(df_clara, df_personas[['NOMBRE COMPLETO', 'Cedula', 'Compania', 'CARGO', 'AREA']], 
                                how='left', left_on='Tarjetahabiente', right_on='NOMBRE COMPLETO')
            
            final_df.drop('NOMBRE COMPLETO', axis=1, inplace=True)

            save_data_to_excel(final_df, excel_file_name)
        else:
            print("Could not read 'Personas.xlsx'. Saving only PDF data.")
            save_data_to_excel(df_clara, excel_file_name)
    else:
        print("No transaction data was found in the PDF. Please check the file and the data format.")
