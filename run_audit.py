from playwright.sync_api import sync_playwright
from read_data import get_links_from_excel
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
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

def generate_report(source_filename, customer_name, all_links, broken_cells):
    """
    Opens the source file:
    1. Removes hyperlinks from all scanned cells and makes text black.
    2. Highlights broken links in Yellow and clears their text.
    3. Saves as 'Customer Revisions.xlsx'.
    """
    print(f"Generating Report for {customer_name}...")
    
    if not os.path.exists(source_filename):
        print(f"Error: Source file {source_filename} not found.")
        return

    # Load Workbook
    wb = load_workbook(source_filename, data_only=False)
    ws = wb.active
    
    # Styles
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    black_font = Font(color="000000")
    
    # 1. CLEANUP: Remove hyperlinks and set text to black for ALL links found
    for item in all_links:
        cell = ws[item['cell']]
        cell.hyperlink = None  # Remove the link
        cell.font = black_font # Set text color to black
    
    # 2. PROCESS BROKEN LINKS: Highlight and Clear Text
    count = 0
    for cell_addr in broken_cells:
        # Get the Revision Letter cell (Column + 1)
        link_cell = ws[cell_addr]
        rev_cell = link_cell.offset(column=1)
        
        # Clear the text
        rev_cell.value = ""
        # Highlight Yellow
        rev_cell.fill = yellow_fill
        
        count += 1
        
    # Save
    output_filename = f"{customer_name} Revisions.xlsx"
    wb.save(output_filename)
    print(f" -> Saved: {output_filename} ({count} broken links cleared & highlighted)")


# --- MAIN EXECUTION FLOW ---

def run_daily_audit():
    print("="*50)
    print("   DAILY REVISION AUDIT: KINNEX & QUATTRO")
    print("="*50)

    # 1. PREPARE DATA
    # UPDATED FILENAMES TO MATCH YOUR SCREENSHOT:
    file_kinnex = "Kinnex Revision Source.xlsx"
    file_quattro = "Quattro Revision Source.xlsx"
    
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
        generate_report(file_kinnex, "Kinnex", links_kinnex, kinnex_broken)
    if links_quattro:
        generate_report(file_quattro, "Quattro", links_quattro, quattro_broken)
    print("-"*30)
    print("ALL AUDITS COMPLETE.")

if __name__ == "__main__":
    run_daily_audit()