import openpyxl
import os

# --- 1. CONFIGURATION: The Maps ---
# This tells the code which cells to look at for each customer.
CUSTOMER_MAPS = {
    "Kinnex": ["A3:A23", "G7:G21"],
    "Quattro": ["A3:A32", "G7:G8", "G12:G22"]
}

def get_links_from_excel(file_path, customer_name):
    print(f"--- PROCESSING {customer_name} ---")
    print(f"Looking for file: {file_path}")
    
    # Check if file exists first
    if not os.path.exists(file_path):
        print(f"ERROR: Could not find file '{file_path}'. Check spelling or folder location.")
        return []

    if customer_name not in CUSTOMER_MAPS:
        print(f"ERROR: No map defined for '{customer_name}'")
        return []

    # Load the workbook. data_only=False ensures we get the hyperlinks.
    wb = openpyxl.load_workbook(file_path, data_only=False)
    ws = wb.active
    
    items_found = []
    ranges = CUSTOMER_MAPS[customer_name]

    for range_string in ranges:
        cell_range = ws[range_string]
        
        # Handle single cell vs range (technical adjustment for openpyxl)
        if not isinstance(cell_range, tuple):
            cell_range = (cell_range,)
            
        for row in cell_range:
            if not isinstance(row, tuple):
                row = (row,)
                
            for cell in row:
                if cell.hyperlink:
                    items_found.append({
                        "cell": cell.coordinate,
                        "text": cell.value,
                        "url": cell.hyperlink.target
                    })
                    # Print what we found to prove it works
                    print(f"  [Found] Cell {cell.coordinate} | Text: {cell.value}")

    return items_found

if __name__ == "__main__":
    # --- TEST ZONE ---
    # I have matched these filenames to your screenshot.
    
    # TEST 1: Run Kinnex
    print("\nTESTING KINNEX:")
    get_links_from_excel("Kinnex Revision Checker.xlsm", "Kinnex")
    
    # TEST 2: Run Quattro
    # (This one runs right after, so we test both at once!)
    print("\nTESTING QUATTRO:")
    get_links_from_excel("Quattro Revision Checker.xlsm", "Quattro")