def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # 1. Remove the line from its current position
    old_line_current = '    _paragraph(doc, f"เลขที่โครงการ {purchase.project_number or \'........................\'}    เลขคุมสัญญา {purchase.contract_control_number or \'........................\'}")\n'
    if old_line_current in code:
        code = code.replace(old_line_current, "")
    else:
        print("Could not find the old line to remove.")
        return

    # 2. Append the new lines at the end of the signature blocks
    # I need to find the end of the signature blocks.
    old_sig2 = '    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)\n'
    new_sig2 = old_sig2 + """    _paragraph(doc, "")
    _paragraph(doc, f"เลขที่โครงการ {purchase.project_number or '........................'}")
    _paragraph(doc, f"เลขคุมสัญญา {purchase.contract_control_number or '........................'}")
"""
    if old_sig2 in code:
        code = code.replace(old_sig2, new_sig2)
    else:
        print("Could not find the second signature block.")
        return

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched!")

patch()
