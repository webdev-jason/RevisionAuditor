from playwright.sync_api import sync_playwright
from read_data import get_links_from_excel
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os

# --- CORE FUNCTIONS ---

def scan_links(page, links, customer_name):
    """
    Takes an active browser page and a list of links.
    Returns a list of cell addresses that are 'Dead'.
    """
    if not links:
        print(f"Skipping {customer_name} (No links found).")
        return []

    print(f"\n--- SCANNING {customer_name} ({len(links)} docs) ---")
    broken_cells = []

    for item in links:
        print(f"Checking {item['text']}...", end="", flush=True)
        
        is_dead = False
        try:
            page.goto(item['url'])
            # Fast check: wait for content to load
            page.wait_for_load_state("domcontentloaded")
            title = page.title()
            
            # CHECK FOR DEATH SIGNALS
            if "Entry not found" in title or "Application Error" in title or "404" in title:
                    is_dead = True
                    print(f" -> DEAD LINK (Highlighting)")
            else:
                print(f" -> OK")
            
        except Exception as e:
            is_dead = True
            print(f" -> Failed to load: {e}")
        
        if is_dead:
            broken_cells.append(item['cell'])
            
    return broken_cells

def generate_report(source_filename, customer_name, broken_cells):
    """
    Opens the source file, highlights broken links in Yellow, saves new report.
    """
    print(f"Generating Report for {customer_name}...")
    
    if not os.path.exists(source_filename):
        print(f"Error: Source file {source_filename} not found.")
        return

    # Load Workbook
    wb = load_workbook(source_filename, data_only=False)
    ws = wb.active
    
    # Highlight Style
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    count = 0
    for cell_addr in broken_cells:
        # Highlight the Revision Letter cell (Column + 1)
        link_cell = ws[cell_addr]
        rev_cell = link_cell.offset(column=1)
        rev_cell.fill = yellow_fill
        count += 1
        
    # Save
    output_filename = f"{customer_name}_Audit_Report.xlsx"
    wb.save(output_filename)
    print(f" -> Saved: {output_filename} ({count} flags)")


# --- MAIN EXECUTION FLOW ---

def run_daily_audit():
    print("="*50)
    print("   DAILY REVISION AUDIT: KINNEX & QUATTRO")
    print("="*50)

    # 1. PREPARE DATA
    file_kinnex = "Kinnex Revision Checker.xlsm"
    file_quattro = "Quattro Revision Checker.xlsm"
    
    links_kinnex = get_links_from_excel(file_kinnex, "Kinnex")
    links_quattro = get_links_from_excel(file_quattro, "Quattro")
    
    if not links_kinnex and not links_quattro:
        print("No data found in either file. Exiting.")
        return

    # 2. BROWSER AUTOMATION (Single Session)
    kinnex_broken = []
    quattro_broken = []

    with sync_playwright() as p:
        print("\nLaunching Browser...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # --- SINGLE LOGIN SEQUENCE ---
        # We pick the first available link just to get to the login screen
        first_url = links_kinnex[0]['url'] if links_kinnex else links_quattro[0]['url']
        
        print(f"\nOPENING LOGIN PAGE...")
        try:
            page.goto(first_url)
        except:
            pass

        input("\n>>> PLEASE LOG IN, THEN PRESS ENTER TO START SCANS <<<")
        print("Starting batch processing...")

        # --- RUN SCANS ---
        kinnex_broken = scan_links(page, links_kinnex, "Kinnex")
        quattro_broken = scan_links(page, links_quattro, "Quattro")

        print("\nClosing Browser...")
        browser.close()

    # 3. GENERATE REPORTS
    print("\n" + "-"*30)
    if links_kinnex:
        generate_report(file_kinnex, "Kinnex", kinnex_broken)
    if links_quattro:
        generate_report(file_quattro, "Quattro", quattro_broken)
    print("-"*30)
    print("ALL AUDITS COMPLETE.")

if __name__ == "__main__":
    run_daily_audit()