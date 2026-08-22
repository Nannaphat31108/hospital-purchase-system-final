def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_block = """    elif form_type == "all":
        # Build the corrected master document first.
        with _DisableDigitFontSplit():
            doc = _build_exact_procurement_template(purchase)

        # Add the Excel-style specification table as a NEW LAST PAGE.
        # This final page contains only this single table.
        _add_final_spec_excel_page(doc, purchase)"""

    new_block = """    elif form_type == "all":
        # เรียง 8, 2, 3, 7, 6 ตามที่ผู้ใช้ต้องการ
        # 8. ชุดรายงานจัดซื้อ (procurement_pack)
        doc = _build_procurement_pack(purchase)
        
        # Configure styles just in case for the appended documents
        _configure_document(doc)
        
        doc.add_page_break()
        # 2. ใบสั่งซื้อ (po)
        _add_purchase_order(doc, purchase)
        
        doc.add_page_break()
        # 3. แบบกำหนด Spec (spec)
        _add_spec(doc, purchase)
        
        doc.add_page_break()
        # 7. แบบแสดงความบริสุทธิ์ใจ (integrity)
        _add_integrity_form(doc, purchase)
        
        doc.add_page_break()
        # 6. ใบตรวจรับ (receipt)
        _add_acceptance_receipt(doc, purchase)"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched export_word 'all' form successfully!")
    else:
        print("Could not find the old block.")
        
patch()
