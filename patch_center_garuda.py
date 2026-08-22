def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_code = """    # Floating pictures: force horizontal position to the page center.
    for anchor in doc.element.xpath(".//wp:anchor"):
        position_h = anchor.find(qn("wp:positionH"))
        if position_h is None:
            position_h = OxmlElement("wp:positionH")
            anchor.insert(0, position_h)

        position_h.set("relativeFrom", "page")

        # Remove old horizontal offset/alignment.
        for child in list(position_h):
            position_h.remove(child)

        align = OxmlElement("wp:align")
        align.text = "center"
        position_h.append(align)"""

    new_code = """    # Floating pictures: convert to inline so they are centered by paragraph alignment
    for anchor in doc.element.xpath(".//wp:anchor"):
        anchor.tag = qn("wp:inline")
        for attr in ['behindDoc', 'locked', 'layoutInCell', 'allowOverlap', 'simplePos', 'relativeHeight']:
            anchor.attrib.pop(attr, None)
        
        for child_tag in ['wp:simplePos', 'wp:positionH', 'wp:positionV', 'wp:wrapNone', 'wp:wrapSquare', 'wp:wrapTight', 'wp:wrapThrough', 'wp:wrapTopAndBottom']:
            for child in anchor.findall(qn(child_tag)):
                anchor.remove(child)"""

    if old_code in code:
        code = code.replace(old_code, new_code)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched _center_all_garuda_paragraphs to convert anchor to inline!")
    else:
        print("Could not find the old code block!")

patch()
