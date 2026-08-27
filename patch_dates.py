def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()
    
    # Change date format in integrity form (since user mentioned doc 5 and 6)
    old = "purchase.document_date.strftime('%d/%m/%Y')"
    new = "format_thai_date_full(purchase.document_date)"
    
    if old in code:
        code = code.replace(old, new)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Replaced all old date formats with full Thai dates!")
    else:
        print("Could not find the old date format.")
        
patch()
