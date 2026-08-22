def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    helper = """def _remove_blank_line_above_memo(doc):
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    for p in doc.element.xpath(".//w:p"):
        text = "".join(t.text for t in p.xpath(".//w:t") if t.text)
        if "บันทึก" in text and "ข้อความ" in text:
            parent = p.getparent()
            
            # Remove space before/after for the memo paragraph itself
            para_obj = Paragraph(p, parent)
            para_obj.paragraph_format.space_before = Pt(0)
            para_obj.paragraph_format.space_after = Pt(0)
            
            paragraphs = parent.xpath("./w:p")
            p_idx = paragraphs.index(p)
            if p_idx > 0:
                prev = paragraphs[p_idx-1]
                prev_text = "".join(t.text for t in prev.xpath(".//w:t") if t.text)
                if not prev_text.strip():
                    parent.remove(prev)
            elif parent.tag.endswith("tc"):
                tc = parent
                tr = tc.getparent()
                tbl = tr.getparent()
                tr_idx = tbl.index(tr)
                
                prev_tr = None
                for i in range(tr_idx-1, -1, -1):
                    if tbl[i].tag.endswith("tr"):
                        prev_tr = tbl[i]
                        break
                        
                if prev_tr is not None:
                    tcs = prev_tr.xpath("./w:tc")
                    if len(tcs) > 0:
                        prev_tc = tcs[0]
                        prev_paragraphs = prev_tc.xpath("./w:p")
                        if prev_paragraphs:
                            last_p = prev_paragraphs[-1]
                            last_p_text = "".join(t.text for t in last_p.xpath(".//w:t") if t.text)
                            if not last_p_text.strip():
                                prev_tc.remove(last_p)
                                # After removing, the new last paragraph is the one with the Garuda
                                new_prev_paragraphs = prev_tc.xpath("./w:p")
                                if new_prev_paragraphs:
                                    garuda_p = new_prev_paragraphs[-1]
                                    garuda_para_obj = Paragraph(garuda_p, prev_tc)
                                    garuda_para_obj.paragraph_format.space_before = Pt(0)
                                    garuda_para_obj.paragraph_format.space_after = Pt(0)

"""

    if "def _remove_blank_line_above_memo(doc):" in code:
        start_idx = code.find("def _remove_blank_line_above_memo(doc):")
        end_idx = code.find("def _center_all_garuda_paragraphs(doc):")
        code = code[:start_idx] + helper + "\n" + code[end_idx:]
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched _remove_blank_line_above_memo with python-docx properties!")
    else:
        print("Could not find block.")

patch()
