def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    helper = """def _remove_trailing_empty_paragraphs(doc):
    while doc.paragraphs and not doc.paragraphs[-1].text.strip():
        # Make sure we don't delete a page break
        has_break = False
        for run in doc.paragraphs[-1].runs:
            if run._element.xpath(".//w:br[@w:type='page']"):
                has_break = True
                break
        if has_break:
            break
        
        p = doc.paragraphs[-1]._element
        p.getparent().remove(p)

"""

    old_block = """    elif form_type == "all":
        # เรียง 8, 2, 3, 7, 6 ตามที่ผู้ใช้ต้องการ
        # 8. ชุดรายงานจัดซื้อ (procurement_pack)
        doc = _build_procurement_pack(purchase)
        
        # Configure styles just in case for the appended documents
        _configure_document(doc)
        
        doc.add_page_break()
        # 2. ใบสั่งซื้อ (po)"""

    new_block = """    elif form_type == "all":
        # เรียง 8, 2, 3, 7, 6 ตามที่ผู้ใช้ต้องการ
        # 8. ชุดรายงานจัดซื้อ (procurement_pack)
        doc = _build_procurement_pack(purchase)
        
        _remove_trailing_empty_paragraphs(doc)
        
        # Configure styles just in case for the appended documents
        _configure_document(doc)
        
        doc.add_page_break()
        # 2. ใบสั่งซื้อ (po)"""

    if old_block in code:
        code = code.replace("def _build_word(purchase, form_type):", helper + "def _build_word(purchase, form_type):")
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched trailing blank pages successfully!")
    else:
        print("Could not find block.")
        
patch()
