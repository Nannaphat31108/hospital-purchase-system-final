def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # Change _add_garuda signature and implementation
    old_def = """def _add_garuda(doc):
    path = Path(__file__).resolve().parent / "static" / "img" / "garuda.png"
    if path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        run.add_picture(str(path), width=Emu(635726), height=Emu(696026))"""
    
    new_def = """def _add_garuda(doc, width=None, height=None):
    path = Path(__file__).resolve().parent / "static" / "img" / "garuda.png"
    if path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        if width is not None and height is not None:
            run.add_picture(str(path), width=width, height=height)
        else:
            run.add_picture(str(path), width=Emu(635726), height=Emu(696026))"""
            
    code = code.replace(old_def, new_def)

    # Update _add_purchase_order call
    old_call = """def _add_purchase_order(doc, purchase):
    _add_garuda(doc)"""
    new_call = """def _add_purchase_order(doc, purchase):
    _add_garuda(doc, width=Cm(1.27), height=Cm(1.4))"""
    code = code.replace(old_call, new_call)

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Garuda patched!")
    
patch()
