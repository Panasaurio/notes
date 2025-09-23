import fitz  # PyMuPDF
import polars as pl
import pandas as pd
from typing import List, Dict, Any, Tuple, Optional

# --- Configuration ---
EXCEL_PATH = "input/EXEMPLE EXCEL.xlsx"
PDF_PATH = "input/CIC.pdf"

# Excel column mappings
EXCEL_ACCOUNT_CELL = {"skiprows": 0, "nrows": 1, "usecols": "I"}
EXCEL_DATA_RANGE = {"skiprows": 3, "nrows": 1, "usecols": "L:M"}
EXCEL_VALUE_COL = "Coût d'acquisition"
EXCEL_CHANGE_COL = "Valorisation "

# PDF rectangle definitions (x0, y0, x1, y1)
PDF_RECTANGLES = {
    "value": fitz.Rect(370, 120, 470, 140),
    "change": fitz.Rect(370, 150, 444, 165),
    "account_num": fitz.Rect(100, 150, 200, 170),
}

# Lookup table
# This now contains a list of 13 elements for each account
LOOKUP_TABLE = {
    "193": ["Juandi", "Debit account", "193000", "CIC", "Funds", "A123", "", "EUR", "Medellin", "Fund XYZ", None, None, "Finance"],
    "100": ["Juanse", "Credit account", "100123", "EXEX", "Stocks", "B456", "", "EUR", "Tolima", "Stock ABC", None, None, "Technology"],
}

# --- Helper Functions ---

def read_excel_data(path: str, **kwargs) -> Optional[Any]:
    """Reads a specific part of an Excel file and returns the first value or a Polars DataFrame."""
    try:
        df_pandas = pd.read_excel(path, **kwargs)
        if kwargs.get("nrows") == 1 and kwargs.get("usecols") == "I":
            return df_pandas.iloc[0, 0] if not df_pandas.empty else None
        return pl.from_pandas(df_pandas)
    except FileNotFoundError:
        print(f"Error: Excel file not found at {path}")
        return None
    except Exception as e:
        print(f"Error reading Excel data: {e}")
        return None

def lookup_account_info(account_num: Optional[Any], lookup_table: Dict[str, List[Any]]) -> List[Optional[Any]]:
    """
    Look up account information from the lookup table.
    Returns a list of 13 elements: [Client, Type de compte, Nom et N. de compte, Nom de banque, Classe d'actifs, Code ISIN,
    Cours, Devise, Zone géographique, Libellé, Quantité (placeholder), Prix de revient (placeholder), Secteur]
    """
    default_info = [None] * 13 # Initialize with 13 Nones
    if account_num is None:
        return default_info

    account_str = str(account_num).strip()
    first_three_digits = account_str[:3] if len(account_str) >= 3 else account_str

    if first_three_digits in lookup_table:
        return lookup_table[first_three_digits]
    
    # If not found, return a default structure with the original account number
    # This assumes "Nom et N. de compte" is at index 2
    default_info[2] = account_str 
    return default_info

def clean_extracted_text(text: str, is_account: bool = False) -> List[str]:
    """Cleans extracted text from PDF rectangles."""
    if not text:
        return []

    lines = [entry.strip() for entry in text.splitlines() if entry.strip()]
    
    if is_account:
        return lines # For accounts, just remove leading/trailing spaces and empty lines
    
    cleaned = []
    for entry in lines:
        entry = entry.replace(" ", "").replace(",", ".")
        # Remove any letter within strings (keep only digits, periods, and currency symbols)
        entry = ''.join(char for char in entry if char.isdigit() or char in {'.', '€', 'EUR'})
        if any(char.isdigit() for char in entry): # Only keep if it still contains digits
            cleaned.append(entry)
    return cleaned

def extract_pdf_data(pdf_path: str, rectangles: Dict[str, fitz.Rect], num_pages: int = 1) -> Dict[str, List[str]]:
    """Extracts text from specified rectangles across given number of PDF pages."""
    extracted_data = {key: [] for key in rectangles}
    
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(min(num_pages, len(doc))):
            page = doc[page_num]
            for key, rect in rectangles.items():
                text = page.get_textbox(rect).strip()
                extracted_data[key].extend(text.splitlines())
        doc.close()
    except FileNotFoundError:
        print(f"Error: PDF file not found at {pdf_path}")
    except Exception as e:
        print(f"Error processing PDF: {e}")
    
    return extracted_data

def combine_and_pad_lists(lists: List[List[Any]]) -> List[List[Any]]:
    """Combines lists and pads them to the maximum length with None."""
    if not lists:
        return []
    
    max_length = max(len(lst) for lst in lists)
    padded_lists = []
    for lst in lists:
        padded_lists.append(lst + [None] * (max_length - len(lst)))
    return padded_lists

# --- Main Logic ---

def main():
    # 1. Read Excel Data
    excel_account = read_excel_data(EXCEL_PATH, **EXCEL_ACCOUNT_CELL)
    df_excel = read_excel_data(EXCEL_PATH, **EXCEL_DATA_RANGE)

    print(f"Account from Excel: {excel_account}")
    # Now excel_account_info will contain the full list of 13 elements
    excel_account_info_list = lookup_account_info(excel_account, LOOKUP_TABLE)
    print(f"Excel account lookup (partial): Full Account: {excel_account_info_list[2]}, Name: {excel_account_info_list[0]}, City: {excel_account_info_list[8]}")

    # 2. Extract PDF Data
    pdf_raw_data = extract_pdf_data(PDF_PATH, PDF_RECTANGLES, num_pages=1)
    
    cleaned_pdf_value = clean_extracted_text(" ".join(pdf_raw_data["value"]))
    cleaned_pdf_change = clean_extracted_text(" ".join(pdf_raw_data["change"]))
    cleaned_pdf_accounts = clean_extracted_text(" ".join(pdf_raw_data["account_num"]), is_account=True)

    # 3. Combine PDF and Excel Data for 'Quantité' and 'Prix de revient'
    combined_quantity = []
    if df_excel is not None and len(df_excel) > 0 and EXCEL_VALUE_COL in df_excel.columns:
        combined_quantity.append(df_excel[EXCEL_VALUE_COL][0])
    combined_quantity.extend(cleaned_pdf_value)

    combined_prix_revient = []
    if df_excel is not None and len(df_excel) > 0 and EXCEL_CHANGE_COL in df_excel.columns:
        combined_prix_revient.append(df_excel[EXCEL_CHANGE_COL][0])
    combined_prix_revient.extend(cleaned_pdf_change)

    combined_accounts_identifiers = []
    if excel_account is not None:
        combined_accounts_identifiers.append(excel_account)
    combined_accounts_identifiers.extend(cleaned_pdf_accounts)

    # Pad lists to ensure equal length for "dynamic" fields
    padded_dynamic_lists = combine_and_pad_lists([combined_quantity, combined_prix_revient, combined_accounts_identifiers])
    combined_quantity, combined_prix_revient, combined_accounts_identifiers = padded_dynamic_lists[0], padded_dynamic_lists[1], padded_dynamic_lists[2]

    # 4. Apply Account Lookup to Combined Accounts and fill the full template
    final_data_rows: List[List[Any]] = []

    # Get the maximum length of any combined list to iterate correctly
    max_rows = len(combined_accounts_identifiers)
    
    for i in range(max_rows):
        account_id = combined_accounts_identifiers[i]
        
        # Get the 13-element template from lookup, filling in defaults if not found
        account_template = lookup_account_info(account_id, LOOKUP_TABLE)
        
        # Override Quantity (index 10) and Prix de revient (index 11) with extracted/excel values
        # Ensure we don't go out of bounds for combined_quantity and combined_prix_revient
        if i < len(combined_quantity):
            account_template[10] = combined_quantity[i] # Quantité
        if i < len(combined_prix_revient):
            account_template[11] = combined_prix_revient[i] # Prix de revient
            
        final_data_rows.append(account_template)

    # 5. Create and Save DataFrame
    if final_data_rows:
        df = pl.DataFrame(
            final_data_rows,
            schema=[
                "Client",
                "Type de compte",
                "Nom et N. de compte",
                "Nom de banque",
                "Classe d'actifs",
                "Code ISIN",
                "Cours",
                "Devise",
                "Zone géographique",
                "Libellé",
                "Quantité",
                "Prix de revient",
                "Secteur"
            ],
            strict=False
        )
        print("\nCombined PDF and Excel Data:")
        print(df)

        df.write_csv("extracted_data_with_full_details.csv")
        print("\nData saved to extracted_data_with_full_details.csv")
    else:
        print("No meaningful data extracted from PDF or Excel file.")

if __name__ == "__main__":
    main()