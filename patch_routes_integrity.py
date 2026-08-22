def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_block = """    _paragraph(doc, f"หมายเหตุ : สำหรับใบสั่งซื้อเลขที่ {purchase.po_number} ลงวันที่ {purchase.document_date.strftime('%d/%m/%Y')}")
    for name, role in persons:
        _paragraph(doc, f"ลงนาม .................................................................\\n({name})\\n({role})", WD_ALIGN_PARAGRAPH.CENTER)"""

    new_block = """    for name, role in persons:
        _paragraph(doc, f"ลงนาม .................................................................\\n({name})\\n({role})", WD_ALIGN_PARAGRAPH.CENTER)
        _paragraph(doc, "")
    _paragraph(doc, f"หมายเหตุ : สำหรับใบสั่งซื้อเลขที่ {purchase.po_number} ลงวันที่ {purchase.document_date.strftime('%d/%m/%Y')}")"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched integrity form!")
    else:
        print("Could not find block!")
        
patch()
