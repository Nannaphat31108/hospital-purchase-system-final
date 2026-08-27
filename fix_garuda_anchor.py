import re

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # The bad loop for finding where to insert wrapNone:
    old_loop = """                # find extent
                extent_idx = 3
                for i, child in enumerate(inline):
                    if child.tag == qn('wp:effectExtent'):
                        extent_idx = i + 1
                        break
                inline.insert(extent_idx, wrapNone)"""
                
    new_loop = """                # insert wrapNone before docPr
                docpr_idx = 3
                for i, child in enumerate(inline):
                    if child.tag == qn('wp:docPr'):
                        docpr_idx = i
                        break
                inline.insert(docpr_idx, wrapNone)
                
                # Also we need to make sure effectExtent exists if required, but wrapNone before docPr is safe.
                # Wait! effectExtent is actually required by the schema for wp:anchor!
                # Let's add an empty effectExtent before wrapNone if it does not exist
                has_effect = any(child.tag == qn('wp:effectExtent') for child in inline)
                if not has_effect:
                    effectExtent = etree.Element(qn('wp:effectExtent'), l="0", t="0", r="0", b="0")
                    inline.insert(docpr_idx, effectExtent)
                    # increment docpr_idx since we just inserted something before it
                    docpr_idx += 1"""
                    
    if old_loop in code:
        code = code.replace(old_loop, new_loop)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Fixed garuda anchor order!")
    else:
        print("Could not find the loop to fix.")

patch()
