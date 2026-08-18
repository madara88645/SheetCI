import os
import openpyxl
from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace
import openpyxl.cell._writer as _cell_writer
import openpyxl.worksheet._writer as _ws_writer

# Dictionary to store custom cached values: (sheet_name, cell_coordinate) -> (cached_value_str, data_type_str)
CACHED_VALUES_MAPPING = {}

# Monkey-patch cell writing to support writing cached values
original_write_cell = _cell_writer.etree_write_cell

def patched_write_cell(xf, worksheet, cell, styled=None):
    key = (worksheet.title, cell.coordinate)
    if cell.data_type == 'f' and key in CACHED_VALUES_MAPPING:
        from openpyxl.cell._writer import _set_attributes
        _, attributes = _set_attributes(cell, styled)
        
        cached_val, data_type = CACHED_VALUES_MAPPING[key]
        if data_type:
            attributes['t'] = data_type
            
        el = Element("c", attributes)
        
        formula = SubElement(el, 'f')
        formula.text = cell.value[1:]  # strip leading '='
        
        cell_content = SubElement(el, 'v')
        cell_content.text = safe_string(cached_val)
        
        xf.write(el)
    else:
        original_write_cell(xf, worksheet, cell, styled)

_cell_writer.etree_write_cell = patched_write_cell
_ws_writer.write_cell = patched_write_cell

def create_broken_model(filepath: str):
    wb = openpyxl.Workbook()
    
    # 1. Main Sheet
    ws = wb.active
    ws.title = "Commissions"
    
    # Setup some data for inconsistent calculations
    ws["B2"] = 1000
    ws["C2"] = 0.05
    ws["D2"] = "=B2*C2"
    
    ws["B3"] = 2000
    ws["C3"] = 0.05
    ws["D3"] = "=B3*C3"
    
    ws["B4"] = 1500
    ws["C4"] = 0.05
    ws["D4"] = "=B4*C4"
    
    ws["B5"] = 3000
    ws["C5"] = 0.05
    ws["D5"] = "=B5*C5"
    
    ws["B6"] = 2500
    ws["C6"] = 0.05
    # Inconsistent formula in repeated column (D)
    ws["D6"] = "=B6*C6+100"
    
    ws["B7"] = 4000
    ws["C7"] = 0.05
    ws["D7"] = "=B7*C7"
    
    # Broken reference formula
    ws["E2"] = "=B2+#REF!"
    
    # Hardcoded number in formula
    ws["F2"] = "=B2*0.08"
    
    # External link-like formula
    ws["G2"] = "=[ExternalSource.xlsx]Sheet1!A1"
    
    # Direct self-reference
    ws["A1"] = "=A1+1"
    
    # Cached formula error (requires patch to write cached error string and set type 'e')
    ws["H2"] = "=B2/0"
    CACHED_VALUES_MAPPING[("Commissions", "H2")] = ("#DIV/0!", "e")
    
    # 2. Hidden Sheet
    hidden_ws = wb.create_sheet(title="HiddenArchive")
    hidden_ws.sheet_state = "hidden"
    hidden_ws["A1"] = "Secret Data"
    
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    print(f"Created broken commission model: {filepath}")

def create_clean_model(filepath: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Commissions"
    
    ws["B2"] = 1000
    ws["C2"] = 0.05
    ws["D2"] = "=B2*C2"
    
    ws["B3"] = 2000
    ws["C3"] = 0.05
    ws["D3"] = "=B3*C3"
    
    ws["B4"] = 1500
    ws["C4"] = 0.05
    ws["D4"] = "=B4*C4"
    
    ws["B5"] = 3000
    ws["C5"] = 0.05
    ws["D5"] = "=B5*C5"
    
    ws["B6"] = 2500
    ws["C6"] = 0.05
    ws["D6"] = "=B6*C6"
    
    ws["B7"] = 4000
    ws["C7"] = 0.05
    ws["D7"] = "=B7*C7"
    
    # Harmless constant formulas (0, 1, 2)
    ws["E2"] = "=D2+0"
    ws["E3"] = "=D3*1"
    ws["E4"] = "=D4-1"
    ws["E5"] = "=D5/2"
    ws["E6"] = "=D6+0"
    ws["E7"] = "=D7+0"
    
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    print(f"Created clean model: {filepath}")

def create_realistic_model(filepath: str):
    """A healthy model written the way a competent analyst writes one.

    Nothing here is a defect: rates live on a dedicated assumptions sheet and are
    referenced absolutely, currency is rounded, and each numeric column has a
    totals row. SheetCI must pass this file.
    """
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Assumptions"
    ws["A1"] = "Parameter"
    ws["B1"] = "Value"
    ws["A2"] = "VAT rate"
    ws["B2"] = 0.20
    ws["A3"] = "Commission rate"
    ws["B3"] = 0.075
    ws["A4"] = "Discount tier"
    ws["B4"] = 0.05

    ws2 = wb.create_sheet("Sales")
    headers = ["Rep", "Region", "Units", "Unit Price", "Gross", "VAT", "Net", "Commission", "Band"]
    for i, header in enumerate(headers, start=1):
        ws2.cell(row=1, column=i, value=header)

    for r in range(2, 42):
        ws2.cell(row=r, column=1, value=f"Rep {r - 1}")
        ws2.cell(row=r, column=2, value="EMEA" if r % 2 else "AMER")
        ws2.cell(row=r, column=3, value=100 + r)
        ws2.cell(row=r, column=4, value=19.99)
        ws2.cell(row=r, column=5, value=f"=C{r}*D{r}")
        ws2.cell(row=r, column=6, value=f"=E{r}*Assumptions!$B$2")
        ws2.cell(row=r, column=7, value=f"=E{r}-F{r}")
        ws2.cell(row=r, column=8, value=f"=ROUND(G{r}*Assumptions!$B$3,2)")
        ws2.cell(row=r, column=9, value=f'=IF(C{r}>150,"High",IF(C{r}>120,"Mid","Low"))')

    ws2["E42"] = "=SUM(E2:E41)"
    ws2["G42"] = "=SUM(G2:G41)"
    ws2["H42"] = "=SUM(H2:H41)"

    ws3 = wb.create_sheet("Summary")
    ws3["A1"] = "Metric"
    ws3["B1"] = "Value"
    ws3["A2"] = "Total Gross"
    ws3["B2"] = "=Sales!E42"
    ws3["A3"] = "Total Commission"
    ws3["B3"] = "=Sales!H42"
    ws3["A4"] = "Avg Unit Price"
    ws3["B4"] = "=AVERAGE(Sales!D2:D41)"
    ws3["B5"] = '=VLOOKUP("Rep 1",Sales!A2:I41,5,FALSE)'
    ws3["B6"] = '=INDEX(Sales!G2:G41,MATCH("Rep 3",Sales!A2:A41,0))'
    ws3["B7"] = "=B3/B2*100"
    ws3["B8"] = "=EOMONTH(TODAY(),12)"

    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    print(f"Created realistic model: {filepath}")

if __name__ == "__main__":
    create_broken_model("examples/broken-commission-model.xlsx")
    create_clean_model("examples/clean-model.xlsx")
    create_realistic_model("examples/realistic-model.xlsx")
