def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()
    
    bad_code = """    table.autofit = False
    widths = [Cm(1.3), Cm(4.5), Cm(2.9), Cm(2.9), Cm(3.1), Cm(2.9)]
    for i in range(6):
        table.columns[i].width = widths[i]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
        table.rows[0].cells[i].width = widths[i]"""
    
    good_code = """    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)"""
    
    if bad_code in code:
        # replace only the second occurrence which is in _add_spec
        # Wait, the string might occur in _add_purchase_order and _add_spec.
        # Let's count them
        occurrences = code.count(bad_code)
        if occurrences == 2:
            parts = code.split(bad_code)
            new_code = parts[0] + bad_code + parts[1] + good_code + parts[2]
            with open("app/routes.py", "w") as f:
                f.write(new_code)
            print("Fixed!")
        else:
            print(f"Found {occurrences} occurrences, expected 2.")
    else:
        print("Bad code not found.")

patch()
