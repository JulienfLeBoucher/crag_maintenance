# Note :
# The file template.ods must be open before using this macro.

import os
from scriptforge import CreateScriptService

CRAGS_FOLDER = '/home/louberehc/veille_aouste/voies'
TEMPLATE_DOC_URL = '/home/louberehc/veille_aouste/template.ods'

#### FUNCTIONS
def get_info_type(text: str) -> str:
    """Qualify the text information."""
    if text.startswith("Secteur"):
        return "Secteur"
    elif text.startswith('-'):
        return "Voie"
    elif text.startswith('\n'):
        return "Saut de ligne"
    else:
        return "Inconnu"

def fill_sector(doc, text_input, current_row):
    # Write the sector name in the A column.
    doc.setValue(
        f"A{current_row}",
        text_input.removeprefix("Secteur ")
    )
    current_row += 1
    return current_row

def fill_route(doc, route_range, text_input, current_row):
    # Copy the table to track route observations.
    doc.copyToCell(route_range, f"B{current_row}")
    # Rectify the route name in the B column.
    doc.setValue(
        f"B{current_row}",
        text_input.removeprefix("- ")
    )
    current_row += 10
    return current_row  
    
def fill_spreadsheet(
    doc, 
    sheet_name,
    route_range,
    text_input: str,
    current_row: int
):
    """ 
    - Fill the spreadsheet according the text_input type.
    - Group and hide rows where each route maintenance will be made.
    - Track the row number of the 'cursor'.
    
    # Args :
        - doc : the target document, a libreoffice calc.
        - sheet_name : a string (name of the crag).
        - route_range : a template range to be copied. 
        - text_input : a text line.
        - current row : the line number where to write information in
        the sheet.
        
    # Return :
        The next row to write to.
    """
    match get_info_type(text_input):
        case "Secteur":
            current_row = fill_sector(doc, text_input, current_row)
        case "Voie":
            initial_current_row = current_row
            current_row = fill_route(
                doc,
                route_range,
                text_input,
                current_row
            )
            # Group and hide
            range_str = f'{sheet_name}.A{initial_current_row + 1}:B{current_row-1}'
            range_add = doc.XCellRange(range_str).RangeAddress
            doc.XSpreadsheet(f'{sheet_name}').group(range_add, 'ROWS')
            doc.XSpreadsheet(f'{sheet_name}').hideDetail(range_add)                
        case "Saut de ligne":
            current_row += 1 
        case _:
            pass
    return current_row


#### MACROS
def create_crags_document(args=None):
    # Get the open spreadsheet
    doc = CreateScriptService("Calc")
    # Get 2 templates which will be copied many times in the spreadsheet
    svc = CreateScriptService("UI")
    source_doc = svc.getDocument(TEMPLATE_DOC_URL)
    header_range = source_doc.Range("Feuille1.A1:G2")
    route_range = source_doc.Range("Feuille1.B4:O13")
    
    # Loop on crags
    crag_files_local = os.listdir("/home/louberehc/veille_aouste/voies")

    for crag_file in crag_files_local:
        CRAG_URL = os.path.join(CRAGS_FOLDER, crag_file)
        crag_name = crag_file.removesuffix('.txt')
        # Generate crag sheet
        doc.InsertSheet(f"{crag_name}")
        doc.Activate(f"{crag_name}")
        # Fill the header
        doc.copyToCell(header_range, "A1") 
        current_row = 3
        # Loop on the information about the crag.
        with open(CRAG_URL, 'r') as crag_info:
            for line in crag_info:
                # Fill the spreadsheet according to the text input type and 
                # update the row position to write next info
                current_row = fill_spreadsheet(
                    doc,
                    crag_name,
                    route_range,
                    line,
                    current_row
                )
            

g_exportedScripts = (create_crags_document,)
