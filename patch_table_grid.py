def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()
        
    old_func = """def _apply_table_grid(table):
    \"\"\"Apply Table Grid when the template contains that built-in style.\"\"\"
    try:
        table.style = 'Table Grid'
    except KeyError:
        pass"""
        
    new_func = """from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def _apply_table_grid(table):
    \"\"\"Apply Table Grid when the template contains that built-in style, fallback to oxml borders.\"\"\"
    try:
        table.style = 'Table Grid'
    except KeyError:
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'auto')
            tblBorders.append(border)
        table._tbl.tblPr.append(tblBorders)"""
        
    if old_func in code:
        code = code.replace(old_func, new_func)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched _apply_table_grid successfully!")
    else:
        print("Could not find _apply_table_grid block.")

patch()
