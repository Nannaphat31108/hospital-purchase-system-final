def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_total = """    cells = table.add_row().cells
    _set_cell_text(cells[0].merge(cells[3]), "รวมเป็นเงินทั้งสิ้น", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _set_cell_text(cells[4], f"{purchase.total_amount:,.2f}", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _paragraph(doc, f"({baht_text(purchase.total_amount)})", WD_ALIGN_PARAGRAPH.CENTER, True)"""
    
    new_total = """    cells = table.add_row().cells
    _set_cell_text(cells[1], f"({baht_text(purchase.total_amount)})", True, WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(cells[4], f"{purchase.total_amount:,.2f}", True, WD_ALIGN_PARAGRAPH.RIGHT)"""

    if old_total in code:
        code = code.replace(old_total, new_total)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched total row!")
    else:
        print("Could not find total row block!")
        
patch()
