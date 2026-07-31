from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path

from flask import Blueprint, flash, jsonify, redirect, render_template, request, send_file, url_for
from sqlalchemy import or_

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .models import Company, DropdownOption, GovernmentProfile, Item, Purchase, PurchaseLine, Unit, db
from .utils import baht_text, safe_filename, to_decimal

main_bp = Blueprint("main", __name__)


@main_bp.app_template_filter("money")
def money(value):
    return f"{to_decimal(value):,.2f}"


@main_bp.app_context_processor
def inject_helpers():
    return {"baht_text": baht_text}


@main_bp.route("/")
def dashboard():
    return render_template(
        "dashboard.html",
        company_count=Company.query.filter_by(active=True).count(),
        item_count=Item.query.filter_by(active=True).count(),
        unit_count=Unit.query.filter_by(active=True).count(),
        purchase_count=Purchase.query.filter_by(active=True).count(),
        recent=Purchase.query.order_by(Purchase.id.desc()).limit(5).all(),
    )


# Companies
@main_bp.route("/companies")
def companies():
    q = request.args.get("q", "").strip()
    query = Company.query
    if q:
        query = query.filter(or_(Company.name.ilike(f"%{q}%"), Company.tax_id.ilike(f"%{q}%"), Company.phone.ilike(f"%{q}%")))
    return render_template("companies/list.html", companies=query.order_by(Company.name).all(), q=q)


@main_bp.route("/companies/new", methods=["GET", "POST"])
@main_bp.route("/companies/<int:company_id>/edit", methods=["GET", "POST"])
def company_form(company_id=None):
    company = Company.query.get_or_404(company_id) if company_id else Company()
    if request.method == "POST":
        company.name = request.form.get("name", "").strip()
        if not company.name:
            flash("กรุณากรอกชื่อบริษัท", "danger")
            return render_template("companies/form.html", company=company)
        for field in ["address", "phone", "tax_id", "bank_name", "bank_branch", "account_no", "account_name"]:
            setattr(company, field, request.form.get(field, "").strip())
        if not company_id:
            db.session.add(company)
        db.session.commit()
        flash("บันทึกข้อมูลบริษัทแล้ว ข้อมูลจะอยู่ถาวรในฐานข้อมูล", "success")
        return redirect(url_for("main.companies"))
    return render_template("companies/form.html", company=company)


@main_bp.post("/companies/<int:company_id>/toggle")
def company_toggle(company_id):
    company = Company.query.get_or_404(company_id)
    company.active = not company.active
    db.session.commit()
    flash("เปลี่ยนสถานะบริษัทแล้ว", "success")
    return redirect(url_for("main.companies"))


# Units
@main_bp.route("/units")
def units():
    q = request.args.get("q", "").strip()
    query = Unit.query
    if q:
        query = query.filter(Unit.name.ilike(f"%{q}%"))
    return render_template("units/list.html", units=query.order_by(Unit.name).all(), q=q)


@main_bp.route("/units/new", methods=["GET", "POST"])
@main_bp.route("/units/<int:unit_id>/edit", methods=["GET", "POST"])
def unit_form(unit_id=None):
    unit = Unit.query.get_or_404(unit_id) if unit_id else Unit()
    if request.method == "POST":
        unit.name = request.form.get("name", "").strip()
        if not unit.name:
            flash("กรุณากรอกชื่อหน่วย", "danger")
            return render_template("units/form.html", unit=unit)
        duplicate = Unit.query.filter(Unit.name == unit.name, Unit.id != (unit.id or 0)).first()
        if duplicate:
            flash("มีหน่วยนี้อยู่แล้ว", "danger")
            return render_template("units/form.html", unit=unit)
        if not unit_id:
            db.session.add(unit)
        db.session.commit()
        flash("บันทึกข้อมูลหน่วยแล้ว ข้อมูลจะอยู่ถาวรในฐานข้อมูล", "success")
        return redirect(url_for("main.units"))
    return render_template("units/form.html", unit=unit)


@main_bp.post("/units/<int:unit_id>/toggle")
def unit_toggle(unit_id):
    unit = Unit.query.get_or_404(unit_id)
    unit.active = not unit.active
    db.session.commit()
    flash("เปลี่ยนสถานะหน่วยแล้ว", "success")
    return redirect(url_for("main.units"))


# Items
@main_bp.route("/items")
def items():
    q = request.args.get("q", "").strip()
    query = Item.query
    if q:
        query = query.filter(or_(Item.name.ilike(f"%{q}%"), Item.code.ilike(f"%{q}%")))
    return render_template("items/list.html", items=query.order_by(Item.name).all(), q=q)


@main_bp.route("/items/new", methods=["GET", "POST"])
@main_bp.route("/items/<int:item_id>/edit", methods=["GET", "POST"])
def item_form(item_id=None):
    item = Item.query.get_or_404(item_id) if item_id else Item()
    units = Unit.query.filter_by(active=True).order_by(Unit.name).all()
    if request.method == "POST":
        item.code = request.form.get("code", "").strip()
        item.name = request.form.get("name", "").strip()
        item.specification = request.form.get("specification", "").strip()
        item.default_price = to_decimal(request.form.get("default_price"))
        item.unit_id = request.form.get("unit_id", type=int)
        if not item.name or not item.unit_id:
            flash("กรุณากรอกชื่อเวชภัณฑ์และหน่วย", "danger")
            return render_template("items/form.html", item=item, units=units)
        if not item_id:
            db.session.add(item)
        db.session.commit()
        flash("บันทึกข้อมูลเวชภัณฑ์แล้ว ข้อมูลจะอยู่ถาวรในฐานข้อมูล", "success")
        return redirect(url_for("main.items"))
    return render_template("items/form.html", item=item, units=units)


@main_bp.post("/items/<int:item_id>/toggle")
def item_toggle(item_id):
    item = Item.query.get_or_404(item_id)
    item.active = not item.active
    db.session.commit()
    flash("เปลี่ยนสถานะเวชภัณฑ์แล้ว", "success")
    return redirect(url_for("main.items"))


@main_bp.get("/api/items/<int:item_id>")
def item_api(item_id):
    item = Item.query.get_or_404(item_id)
    return jsonify({
        "id": item.id,
        "name": item.name,
        "price": f"{item.default_price:.2f}",
        "unit_id": item.unit_id,
        "unit_name": item.unit.name,
        "specification": item.specification or "",
    })


# Master dropdown data
@main_bp.route("/masters")
def masters():
    profiles = GovernmentProfile.query.order_by(GovernmentProfile.id).all()
    options = DropdownOption.query.order_by(DropdownOption.category, DropdownOption.id).all()
    return render_template("masters/list.html", profiles=profiles, options=options)


@main_bp.route("/masters/profile/new", methods=["GET", "POST"])
@main_bp.route("/masters/profile/<int:profile_id>/edit", methods=["GET", "POST"])
def profile_form(profile_id=None):
    profile = GovernmentProfile.query.get_or_404(profile_id) if profile_id else GovernmentProfile()
    if request.method == "POST":
        for field in ["name","department","letter_prefix","recipient","officer_name","officer_position","chief_name","chief_position","approver_name","approver_position","inspector1_name","inspector1_position","inspector2_name","inspector2_position","inspector3_name","inspector3_position","specifier_name","specifier_position","receipt1_name","receipt2_name","receipt3_name"]:
            setattr(profile, field, request.form.get(field, "").strip())
        if not profile_id:
            db.session.add(profile)
        db.session.commit()
        flash("บันทึกข้อมูลราชการแล้ว", "success")
        return redirect(url_for("main.masters"))
    return render_template("masters/profile_form.html", profile=profile)


@main_bp.route("/masters/option/new", methods=["GET", "POST"])
@main_bp.route("/masters/option/<int:option_id>/edit", methods=["GET", "POST"])
def option_form(option_id=None):
    option = DropdownOption.query.get_or_404(option_id) if option_id else DropdownOption()
    if request.method == "POST":
        option.category = request.form.get("category", "").strip()
        option.value = request.form.get("value", "").strip()
        if not option.category or not option.value:
            flash("กรุณากรอกหมวดและค่า", "danger")
        else:
            if not option_id:
                db.session.add(option)
            db.session.commit()
            flash("บันทึกตัวเลือก Dropdown แล้ว", "success")
            return redirect(url_for("main.masters"))
    return render_template("masters/option_form.html", option=option)


@main_bp.post("/masters/option/<int:option_id>/toggle")
def option_toggle(option_id):
    option = DropdownOption.query.get_or_404(option_id)
    option.active = not option.active
    db.session.commit()
    return redirect(url_for("main.masters"))


# Purchases
@main_bp.route("/purchases")
def purchases():
    q = request.args.get("q", "").strip()
    query = Purchase.query.join(Company)
    if q:
        query = query.filter(or_(Purchase.po_number.ilike(f"%{q}%"), Purchase.project_number.ilike(f"%{q}%"), Company.name.ilike(f"%{q}%")))
    return render_template("purchases/list.html", purchases=query.order_by(Purchase.document_date.desc(), Purchase.id.desc()).all(), q=q)


def purchase_context(purchase):
    categories = {}
    for option in DropdownOption.query.filter_by(active=True).order_by(DropdownOption.category, DropdownOption.id).all():
        categories.setdefault(option.category, []).append(option)
    return {
        "purchase": purchase,
        "companies": Company.query.filter_by(active=True).order_by(Company.name).all(),
        "items": Item.query.filter_by(active=True).order_by(Item.name).all(),
        "units": Unit.query.filter_by(active=True).order_by(Unit.name).all(),
        "government_profiles": GovernmentProfile.query.filter_by(active=True).order_by(GovernmentProfile.name).all(),
        "dropdowns": categories,
        "today": date.today().isoformat(),
    }


@main_bp.route("/purchases/new", methods=["GET", "POST"])
@main_bp.route("/purchases/<int:purchase_id>/edit", methods=["GET", "POST"])
def purchase_form(purchase_id=None):
    purchase = Purchase.query.get_or_404(purchase_id) if purchase_id else Purchase(document_date=date.today())
    if request.method == "POST":
        po_number = request.form.get("po_number", "").strip()
        duplicate = Purchase.query.filter(Purchase.po_number == po_number, Purchase.id != (purchase.id or 0)).first()
        if not po_number or duplicate:
            flash("กรุณากรอกเลขที่ใบสั่งซื้อที่ไม่ซ้ำ", "danger")
            return render_template("purchases/form.html", **purchase_context(purchase))

        purchase.po_number = po_number
        try:
            purchase.document_date = datetime.strptime(request.form.get("document_date", ""), "%Y-%m-%d").date()
        except ValueError:
            purchase.document_date = date.today()
        purchase.company_id = request.form.get("company_id", type=int)
        purchase.government_profile_id = request.form.get("government_profile_id", type=int)
        purchase.procurement_type = request.form.get("procurement_type", "").strip() or "เวชภัณฑ์มิใช่ยา"
        purchase.necessity_reason = request.form.get("necessity_reason", "").strip() or "ใช้ในการรักษาผู้ป่วย"
        purchase.project_number = request.form.get("project_number", "").strip()
        purchase.contract_control_number = request.form.get("contract_control_number", "").strip()
        purchase.delivery_days = request.form.get("delivery_days", type=int) or 30
        purchase.delivery_place = request.form.get("delivery_place", "").strip()
        purchase.budget_source = request.form.get("budget_source", "").strip()
        purchase.budget_allocated = to_decimal(request.form.get("budget_allocated"))
        purchase.budget_previously_used = to_decimal(request.form.get("budget_previously_used"))
        purchase.note = request.form.get("note", "").strip()

        lines = []
        for index in range(1, 7):
            item_id = request.form.get(f"item_id_{index}", type=int)
            if not item_id:
                continue
            item = db.session.get(Item, item_id)
            if item is None:
                continue
            quantity = to_decimal(request.form.get(f"quantity_{index}"))
            unit_price = to_decimal(request.form.get(f"unit_price_{index}"))
            unit_id = request.form.get(f"unit_id_{index}", type=int) or item.unit_id
            if quantity <= 0:
                continue
            lines.append(PurchaseLine(
                line_no=len(lines) + 1,
                item_id=item.id,
                description=request.form.get(f"description_{index}", "").strip() or item.name,
                specification=request.form.get(f"specification_{index}", "").strip() or item.specification,
                quantity=quantity,
                unit_id=unit_id,
                unit_price=unit_price,
                amount=(quantity * unit_price).quantize(Decimal("0.01")),
            ))

        if not purchase.company_id or not lines:
            flash("กรุณาเลือกบริษัทและเพิ่มอย่างน้อย 1 รายการ", "danger")
            return render_template("purchases/form.html", **purchase_context(purchase))

        if not purchase_id:
            db.session.add(purchase)
        purchase.lines.clear()
        purchase.lines.extend(lines)
        db.session.commit()
        flash("บันทึกใบสั่งซื้อแล้ว ข้อมูลถูกเก็บในฐานข้อมูลถาวร สามารถส่งออก Word ได้ด้านล่าง", "success")
        return redirect(url_for("main.purchase_detail", purchase_id=purchase.id))

    return render_template("purchases/form.html", **purchase_context(purchase))


@main_bp.route("/purchases/<int:purchase_id>")
def purchase_detail(purchase_id):
    return render_template("purchases/detail.html", purchase=Purchase.query.get_or_404(purchase_id))


@main_bp.post("/purchases/<int:purchase_id>/toggle")
def purchase_toggle(purchase_id):
    purchase = Purchase.query.get_or_404(purchase_id)
    purchase.active = not purchase.active
    db.session.commit()
    flash("เปลี่ยนสถานะเอกสารแล้ว", "success")
    return redirect(url_for("main.purchases"))


def _vat_values(purchase):
    total = to_decimal(purchase.total_amount)
    subtotal = (total / Decimal("1.07")).quantize(Decimal("0.01")) if total else Decimal("0.00")
    vat = (total - subtotal).quantize(Decimal("0.01"))
    return subtotal, vat


@main_bp.route("/purchases/<int:purchase_id>/print/<form_type>")
def print_form(purchase_id, form_type):
    purchase = Purchase.query.get_or_404(purchase_id)
    if form_type == "po":
        template = "print/purchase_single.html" if len(purchase.lines) == 1 else "print/purchase_multiple.html"
    elif form_type == "spec":
        template = "print/spec.html"
    elif form_type == "all":
        template = "print/all.html"
    else:
        return "ไม่พบแบบฟอร์ม", 404
    subtotal, vat = _vat_values(purchase)
    return render_template(template, purchase=purchase, subtotal_before_vat=subtotal, vat_amount=vat, embedded=False)


# ---------------- Word export ----------------
def _set_cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=16):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text or ""))
    run.bold = bold
    _set_run_font(run, size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return p


def _set_run_font(run, size=16, bold=None):
    run.font.name = "TH Sarabun New"
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), "TH Sarabun New")
    rfonts.set(qn("w:hAnsi"), "TH Sarabun New")
    rfonts.set(qn("w:eastAsia"), "TH Sarabun New")
    rfonts.set(qn("w:cs"), "TH Sarabun New")


def _paragraph(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size=16, first_line=False, space_after=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    if first_line:
        p.paragraph_format.first_line_indent = Cm(1.25)
    run = p.add_run(str(text))
    _set_run_font(run, size, bold)
    return p


def _configure_document(doc):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "TH Sarabun New"
    normal.font.size = Pt(16)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "TH Sarabun New")


def _add_garuda(doc):
    path = Path(__file__).resolve().parent / "static" / "img" / "garuda.png"
    if path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(str(path), width=Cm(2.8))

def _apply_table_grid(table):
    """Apply Table Grid when the template contains that built-in style."""
    try:
        table.style = "Table Grid"
    except KeyError:
        pass


def _add_purchase_order(doc, purchase):
    _add_garuda(doc)
    _paragraph(doc, "ใบสั่งซื้อ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)

    meta = doc.add_table(rows=1, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta.autofit = True
    left = (
        f"ผู้ขาย {purchase.company.name}\n"
        f"ที่อยู่ {purchase.company.address or '-'}\n"
        f"โทรศัพท์ {purchase.company.phone or '-'}\n"
        f"เลขประจำตัวผู้เสียภาษี {purchase.company.tax_id or '-'}\n"
        f"เลขที่บัญชีเงินฝากธนาคาร {purchase.company.account_no or '-'}\n"
        f"ชื่อบัญชี {purchase.company.account_name or '-'}\n"
        f"ธนาคาร {purchase.company.bank_name or '-'}"
        + (f" สาขา {purchase.company.bank_branch}" if purchase.company.bank_branch else "")
    )
    right = (
        f"เลขที่ {purchase.po_number}\n"
        f"วันที่ {purchase.document_date.strftime('%d/%m/%Y')}\n\n"
        "ส่วนราชการ โรงพยาบาลสิงห์บุรี\n"
        "ที่อยู่ ๙๑๗/๓ ตำบลบางพุทรา อำเภอเมืองสิงห์บุรี จังหวัดสิงห์บุรี ๑๖๐๐๐\n"
        "โทรศัพท์ ๐๓๖-๕๒๒๕๐๗"
    )
    _set_cell_text(meta.cell(0, 0), left)
    _set_cell_text(meta.cell(0, 1), right)

    _paragraph(doc, f"ตามที่ {purchase.company.name} ได้เสนอราคาไว้ต่อโรงพยาบาลสิงห์บุรี ซึ่งได้รับราคาและตกลงซื้อตามรายการดังต่อไปนี้", first_line=True)

    table = doc.add_table(rows=1, cols=6)
    _apply_table_grid(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["ลำดับ", "รายการ", "จำนวน", "หน่วย", "ราคาต่อหน่วย\n(บาท)", "จำนวนเงิน\n(บาท)"]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
    for line in purchase.lines:
        cells = table.add_row().cells
        values = [line.line_no, line.description, f"{line.quantity:g}", line.unit.name, f"{line.unit_price:,.2f}", f"{line.amount:,.2f}"]
        for i, value in enumerate(values):
            align = WD_ALIGN_PARAGRAPH.LEFT if i == 1 else (WD_ALIGN_PARAGRAPH.RIGHT if i >= 4 else WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(cells[i], value, align=align)

    subtotal, vat = _vat_values(purchase)
    for label, value in [("รวมเป็นเงิน", subtotal), ("ภาษีมูลค่าเพิ่ม", vat), ("รวมเป็นเงินทั้งสิ้น", purchase.total_amount)]:
        cells = table.add_row().cells
        merged = cells[0].merge(cells[3])
        _set_cell_text(merged, f"({baht_text(purchase.total_amount)})" if label == "รวมเป็นเงิน" else "", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(cells[4], label, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_cell_text(cells[5], f"{value:,.2f}", bold=(label == "รวมเป็นเงินทั้งสิ้น"), align=WD_ALIGN_PARAGRAPH.RIGHT)

    _paragraph(doc, "การซื้ออยู่ภายใต้เงื่อนไขต่อไปนี้", bold=True)
    conditions = [
        f"๑. กำหนดส่งมอบภายใน {purchase.delivery_days} วัน นับถัดจากวันที่ผู้ขายได้รับใบสั่งซื้อ",
        "๒. ครบกำหนดส่งมอบวันที่ ................................................................",
        f"๓. สถานที่ส่งมอบ {purchase.delivery_place or 'โรงพยาบาลสิงห์บุรี ๙๑๗/๓'}",
        "๔. ระยะเวลารับประกัน -",
        "๕. สงวนสิทธิ์ค่าปรับกรณีส่งมอบเกินกำหนด โดยคิดค่าปรับเป็นรายวันในอัตราร้อยละ ๐.๒๐ ของราคาสิ่งของที่ยังไม่ได้รับมอบ",
        "๖. ส่วนราชการสงวนสิทธิ์ที่จะไม่รับมอบ หากสินค้ามีลักษณะไม่ตรงตามรายการที่ระบุไว้ในใบสั่งซื้อ ผู้ขายจะต้องดำเนินการเปลี่ยนใหม่ให้ถูกต้องทุกประการ",
        "๗. หน่วยงานของรัฐสามารถนำผลการปฏิบัติงานตามสัญญาหรือข้อตกลงมาประเมินผลการปฏิบัติงานของผู้ประกอบการ",
    ]
    for text in conditions:
        _paragraph(doc, text)

    _paragraph(doc, "หมายเหตุ", bold=True)
    _paragraph(doc, "๑. การติดอากรแสตมป์ให้เป็นไปตามประมวลกฎหมายรัษฎากร หากต้องการให้ใบสั่งซื้อมีผลตามกฎหมาย")
    _paragraph(doc, f"๒. ใบสั่งซื้อนี้อ้างอิงตามเลขที่โครงการ {purchase.project_number or '........................'} ซื้อพัสดุจำนวน {len(purchase.lines)} รายการ เป็นเงิน {purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)}) โดยวิธีเฉพาะเจาะจง")
    _paragraph(doc, f"เลขที่โครงการ {purchase.project_number or '........................'}    เลขคุมสัญญา {purchase.contract_control_number or '........................'}")
    _paragraph(doc, "ลงชื่อ ................................................ ผู้สั่งซื้อ\n(................................................)\nหัวหน้าเจ้าหน้าที่\nวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\n(................................................)\nวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)

def _add_spec(doc, purchase):
    _add_garuda(doc)
    _paragraph(doc, "บันทึกข้อความ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, "ส่วนราชการ โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม โทร. ๐๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙")
    _paragraph(doc, f"ที่ สห ๐๐๓๓.๒๐๕.๑๒/........................ วันที่ {purchase.document_date.strftime('%d/%m/%Y')}")
    _paragraph(doc, f"เรื่อง ขออนุมัติแต่งตั้งผู้กำหนดรายละเอียดคุณลักษณะเฉพาะของพัสดุ จำนวน {len(purchase.lines)} รายการ")
    _paragraph(doc, "เรียน ผู้ว่าราชการจังหวัดสิงห์บุรี")
    _paragraph(doc, "ด้วยกลุ่มงานเภสัชกรรม โรงพยาบาลสิงห์บุรี จะดำเนินการจัดซื้อพัสดุ ดังนี้", first_line=True)

    table = doc.add_table(rows=1, cols=5)
    _apply_table_grid(table)
    headers = ["ลำดับ", "รายการ", "จำนวน", "หน่วยละ", "เป็นเงิน"]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
    for line in purchase.lines:
        cells = table.add_row().cells
        values = [line.line_no, line.description, f"{line.quantity:g} {line.unit.name}", f"{line.unit_price:,.2f}", f"{line.amount:,.2f}"]
        for i, value in enumerate(values):
            _set_cell_text(cells[i], value, align=WD_ALIGN_PARAGRAPH.LEFT if i == 1 else WD_ALIGN_PARAGRAPH.CENTER)
    cells = table.add_row().cells
    _set_cell_text(cells[0].merge(cells[3]), "รวมเป็นเงินทั้งสิ้น", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _set_cell_text(cells[4], f"{purchase.total_amount:,.2f}", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _paragraph(doc, f"({baht_text(purchase.total_amount)})", WD_ALIGN_PARAGRAPH.CENTER, True)
    _paragraph(doc, "เพื่อให้ได้ร่างขอบเขตของงานหรือรายละเอียดคุณลักษณะเฉพาะของพัสดุดังกล่าว รวมทั้งกำหนดหลักเกณฑ์การพิจารณาคัดเลือกข้อเสนอ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ จึงขออนุมัติแต่งตั้งผู้กำหนดรายละเอียดคุณลักษณะเฉพาะ", first_line=True)
    _paragraph(doc, "ลงชื่อ ................................................ เจ้าหน้าที่                         อนุมัติ ลงชื่อ ................................................", WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    _paragraph(doc, "รายละเอียดคุณลักษณะเฉพาะ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, f"พัสดุจำนวน {len(purchase.lines)} รายการ", WD_ALIGN_PARAGRAPH.CENTER, True)
    for line in purchase.lines:
        _paragraph(doc, f"{line.line_no}. {line.description}", bold=True)
        _paragraph(doc, line.specification or "รายละเอียดตามที่หน่วยงานกำหนด")
        _paragraph(doc, f"จำนวน {line.quantity:g} {line.unit.name}")
    _paragraph(doc, "ลงชื่อ ................................................ ผู้กำหนดรายละเอียด\n(................................................)", WD_ALIGN_PARAGRAPH.CENTER)



def _add_acceptance_receipt(doc, purchase):
    _paragraph(doc, "ใบตรวจรับการจัดซื้อ/จัดจ้าง", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, f"วันที่ {purchase.document_date.strftime('%d/%m/%Y')}", WD_ALIGN_PARAGRAPH.RIGHT)
    _paragraph(
        doc,
        f"ตามใบสั่งซื้อ เลขที่ {purchase.po_number} ลงวันที่ {purchase.document_date.strftime('%d/%m/%Y')} "
        f"โรงพยาบาลสิงห์บุรีได้ตกลงซื้อกับ {purchase.company.name} สำหรับโครงการซื้อพัสดุ "
        f"จำนวน {len(purchase.lines)} รายการ โดยวิธีเฉพาะเจาะจง เป็นจำนวนเงินทั้งสิ้น "
        f"{purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)})",
        first_line=True,
    )
    _paragraph(doc, "คณะกรรมการตรวจรับพัสดุ ได้ตรวจรับงานแล้ว ผลปรากฏดังนี้", first_line=True)
    _paragraph(doc, "1. ผลการตรวจรับ")
    _paragraph(doc, "     ☐ ถูกต้อง      ☐ ครบถ้วนตามสัญญา      ☐ ไม่ครบถ้วนตามสัญญา")
    _paragraph(doc, "2. ค่าปรับ")
    _paragraph(doc, "     ☐ มีค่าปรับ      ☐ ไม่มีค่าปรับ")
    _paragraph(doc, "3. การเบิกจ่ายเงิน")
    if len(purchase.lines) == 1:
        line = purchase.lines[0]
        _paragraph(doc, f"     เบิกจ่ายเงิน เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")
    else:
        for line in purchase.lines:
            _paragraph(doc, f"     - รายการที่ {line.line_no} {line.description}")
            _paragraph(doc, f"       เบิกจ่ายเงิน งวดที่ 1 เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")

    _paragraph(doc, "", space_after=0)
    sig = doc.add_table(rows=3, cols=2)
    sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    names = [
        "(นางสาวธัญรดา ใจเสน)",
        "(นางสาวจุฑามาศ อาคมสรรเสริญ)",
        "(นางสาวนริศรา ม่วงงาม)",
    ]
    for i, name in enumerate(names):
        _set_cell_text(sig.cell(i, 1), f"(ลงชื่อ) .................................................. ผู้ตรวจรับ\n{name}", align=WD_ALIGN_PARAGRAPH.CENTER)
        if i == 1:
            _set_cell_text(sig.cell(i, 0), "เรียน ผู้ว่าราชการจังหวัดสิงห์บุรี\n- เพื่อทราบผลการตรวจรับ\n\n(นางพิณนภา ศริพันธุ์)\nเจ้าหน้าที่", align=WD_ALIGN_PARAGRAPH.CENTER)
        elif i == 2:
            _set_cell_text(sig.cell(i, 0), "(นายชัชวาล บุญญฤทธิ์)\nหัวหน้าเจ้าหน้าที่", align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text(sig.cell(i, 0), "", align=WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, f"หมายเหตุ : เลขที่โครงการ {purchase.project_number or '........................'}", space_after=0)
    _paragraph(doc, f"            เลขคุมสัญญา {purchase.contract_control_number or '........................'}", space_after=0)


def _add_integrity_form(doc, purchase):
    _paragraph(doc, "แบบแสดงความบริสุทธิ์ใจในการจัดซื้อจัดจ้างทุกวิธีของหน่วยงาน", WD_ALIGN_PARAGRAPH.CENTER, True, 18)
    _paragraph(doc, "ในการเปิดเผยข้อมูลความขัดแย้งทางผลประโยชน์", WD_ALIGN_PARAGRAPH.CENTER, True, 18)
    _paragraph(doc, "ของหัวหน้าพัสดุ เจ้าหน้าที่พัสดุ และคณะกรรมการตรวจรับพัสดุ", WD_ALIGN_PARAGRAPH.CENTER, True, 18)
    _paragraph(doc, "-----------------------------------", WD_ALIGN_PARAGRAPH.CENTER)
    persons = [
        ("นายชัชวาลย์ บุญญฤทธิ์", "หัวหน้าเจ้าหน้าที่"),
        ("นางพิณนภา ศริพันธุ์", "เจ้าหน้าที่พัสดุ"),
        ("นางสาวกัญญพัชร ธรกิจการค้า", "คณะกรรมการตรวจรับพัสดุ"),
        ("นางสาวชุลีพร สุขมี", "คณะกรรมการตรวจรับพัสดุ"),
        ("นางสาวกัญญาพัชร เลิศอนันตกูล", "คณะกรรมการตรวจรับพัสดุ"),
    ]
    for name, role in persons:
        _paragraph(doc, f"ข้าพเจ้า {name} ({role})")
    _paragraph(
        doc,
        "ขอให้คำรับรองว่าไม่มีความเกี่ยวข้องหรือมีส่วนได้ส่วนเสียไม่ว่าโดยตรงหรือโดยอ้อม "
        "หรือผลประโยชน์ใด ๆ ที่ก่อให้เกิดความขัดแย้งทางผลประโยชน์กับผู้ขาย ผู้รับจ้าง "
        "ผู้เสนองาน หรือผู้ชนะประมูล หรือผู้มีส่วนเกี่ยวข้องที่เข้ามามีนิติสัมพันธ์ "
        "และวางตัวเป็นกลางในการดำเนินการเกี่ยวกับการพัสดุ ปฏิบัติหน้าที่ด้วยจิตสำนึก "
        "ด้วยความโปร่งใส สามารถให้ผู้เกี่ยวข้องตรวจสอบได้ทุกเวลา และมุ่งประโยชน์ส่วนรวมเป็นสำคัญ",
        first_line=True,
    )
    _paragraph(
        doc,
        "หากปรากฏว่าเกิดความขัดแย้งทางผลประโยชน์ระหว่างข้าพเจ้ากับผู้ขาย ผู้รับจ้าง "
        "ผู้เสนองาน หรือผู้ชนะประมูล หรือผู้มีส่วนเกี่ยวข้องที่เข้ามามีนิติสัมพันธ์ "
        "ข้าพเจ้าจะรายงานให้ทราบโดยทันที",
        first_line=True,
    )
    _paragraph(doc, f"หมายเหตุ : สำหรับใบสั่งซื้อเลขที่ {purchase.po_number} ลงวันที่ {purchase.document_date.strftime('%d/%m/%Y')}")
    for name, role in persons:
        _paragraph(doc, f"ลงนาม .................................................................\n({name})\n({role})", WD_ALIGN_PARAGRAPH.CENTER)


def _add_procurement_pack(doc, purchase):
    """Create the government procurement document pack from the supplied examples."""
    # 1) Purchase request memorandum
    _add_garuda(doc)
    _paragraph(doc, "บันทึกข้อความ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, "ส่วนราชการ โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม โทร. ๐๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙")
    _paragraph(doc, f"ที่ สห ๐๐๓๓.๒๐๕.๑๒/........................ วันที่ {purchase.document_date.strftime('%d/%m/%Y')}")
    subject = purchase.lines[0].description if len(purchase.lines) == 1 else f"เวชภัณฑ์มิใช่ยา จำนวน {len(purchase.lines)} รายการ"
    _paragraph(doc, f"เรื่อง รายงานขอซื้อ {subject}")
    _paragraph(doc, "เรียน ผู้ว่าราชการจังหวัดสิงห์บุรี")
    _paragraph(doc, f"ด้วยโรงพยาบาลสิงห์บุรีมีความประสงค์จะซื้อ {subject} โดยวิธีเฉพาะเจาะจง ซึ่งมีรายละเอียดดังต่อไปนี้", first_line=True)
    _paragraph(doc, "๑. เหตุผลความจำเป็นที่ต้องซื้อ\n    ใช้ในการรักษาผู้ป่วย")
    _paragraph(doc, f"๒. รายละเอียดของพัสดุ\n    {subject}")
    _paragraph(doc, f"๓. ราคากลางของพัสดุ เป็นเงิน {purchase.total_amount:,.2f} บาท")
    _paragraph(doc, f"๔. วงเงินที่จะซื้อ {purchase.budget_source or 'เงินบำรุงโรงพยาบาลสิงห์บุรี'} จำนวน {purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)})")
    _paragraph(doc, f"๕. กำหนดเวลาส่งมอบภายใน {purchase.delivery_days} วัน นับถัดจากวันลงนามในสัญญา")
    _paragraph(doc, "๖. ดำเนินการโดยวิธีเฉพาะเจาะจง")
    _paragraph(doc, "๗. หลักเกณฑ์การพิจารณาคัดเลือกข้อเสนอโดยใช้เกณฑ์ราคา")
    _paragraph(doc, "๘. ผู้ตรวจรับพัสดุ\n    ๑. นางสาวกัญญพัชร ธนกิจการค้า ประธานกรรมการ\n    ๒. นางสาวชุลีพร สุขมี กรรมการ\n    ๓. นางสาวกัญญาพัชร เลิศอนันตกูล กรรมการ")
    _paragraph(doc, "จึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ", first_line=True)
    _paragraph(doc, "(นางพิณนภา ศริพันธุ์)\nเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)

    # 2) Consideration memorandum
    doc.add_page_break()
    _add_garuda(doc)
    _paragraph(doc, "บันทึกข้อความ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, f"เรื่อง รายงานผลการพิจารณาและขออนุมัติสั่งซื้อ {subject}")
    _paragraph(doc, "เรียน ผู้ว่าราชการจังหวัดสิงห์บุรี")
    _paragraph(doc, f"ขอรายงานผลการพิจารณาซื้อ {subject} โดยวิธีเฉพาะเจาะจง ดังนี้", first_line=True)
    table = doc.add_table(rows=2, cols=4)
    _apply_table_grid(table)
    headers = ["รายการพิจารณา", "รายชื่อผู้ยื่นข้อเสนอ", "ราคาที่เสนอ*", "ราคาที่ตกลงซื้อหรือจ้าง*"]
    for i, h in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], h, True, WD_ALIGN_PARAGRAPH.CENTER)
    vals = [subject, purchase.company.name, f"{purchase.total_amount:,.2f}", f"{purchase.total_amount:,.2f}"]
    for i, v in enumerate(vals):
        _set_cell_text(table.rows[1].cells[i], v, align=WD_ALIGN_PARAGRAPH.CENTER if i != 0 else WD_ALIGN_PARAGRAPH.LEFT)
    _paragraph(doc, "*ราคาที่เสนอและราคาที่ตกลงซื้อหรือจ้างเป็นราคารวมภาษีมูลค่าเพิ่ม ภาษีอื่น ค่าขนส่ง ค่าจดทะเบียน และค่าใช้จ่ายอื่นทั้งปวง", size=14)
    _paragraph(doc, "โรงพยาบาลสิงห์บุรีพิจารณาแล้ว เห็นสมควรจัดซื้อจากผู้เสนอราคาดังกล่าว จึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ", first_line=True)
    _paragraph(doc, "(นางพิณนภา ศริพันธุ์)\nเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)

    # 3) Winner announcement
    doc.add_page_break()
    _add_garuda(doc)
    _paragraph(doc, "ประกาศจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, f"เรื่อง ประกาศผู้ชนะการเสนอราคา ซื้อ {subject}", WD_ALIGN_PARAGRAPH.CENTER, True, 18)
    _paragraph(doc, "โดยวิธีเฉพาะเจาะจง", WD_ALIGN_PARAGRAPH.CENTER, True, 18)
    _paragraph(doc, "-----------------------------------", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, f"ตามที่โรงพยาบาลสิงห์บุรีได้มีโครงการซื้อ {subject} โดยวิธีเฉพาะเจาะจงนั้น ผู้ได้รับการคัดเลือก ได้แก่ {purchase.company.name} โดยเสนอราคาเป็นเงินทั้งสิ้น {purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)}) รวมภาษีมูลค่าเพิ่มและค่าใช้จ่ายอื่นทั้งปวง", first_line=True)
    _paragraph(doc, f"ประกาศ ณ วันที่ {purchase.document_date.strftime('%d/%m/%Y')}", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายพิรุณ ปิตะหงษ์นันท์)\nผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติราชการแทน\nผู้ว่าราชการจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER)

def _replace_text_once_in_paragraph(paragraph, old, new):
    """Replace text across runs while preserving surrounding formatting."""
    old = str(old or "")
    new = str(new or "")
    if not old or not paragraph.runs:
        return False

    full_text = "".join(run.text for run in paragraph.runs)
    start_index = full_text.find(old)
    if start_index < 0:
        return False
    end_index = start_index + len(old)

    positions = []
    cursor = 0
    for index, run in enumerate(paragraph.runs):
        positions.append((index, cursor, cursor + len(run.text)))
        cursor += len(run.text)

    first_index = next((i for i, s, e in positions if e > start_index), None)
    last_index = None
    for i, s, e in positions:
        if s < end_index:
            last_index = i

    if first_index is None or last_index is None:
        return False

    first_run = paragraph.runs[first_index]
    last_run = paragraph.runs[last_index]
    first_start = positions[first_index][1]
    last_start = positions[last_index][1]

    prefix = first_run.text[:max(0, start_index - first_start)]
    suffix = last_run.text[max(0, end_index - last_start):]
    first_run.text = prefix + new + suffix
    _set_run_font(first_run, 16)

    for index in range(first_index + 1, last_index + 1):
        paragraph.runs[index].text = ""
    return True


def _replace_paragraph_text(paragraph, replacements):
    # Replace each source phrase at most once per paragraph.
    # Using a while-loop can repeat forever when the replacement text
    # still contains the source phrase.
    for old, new in replacements:
        _replace_text_once_in_paragraph(paragraph, old, new)


def _replace_in_document(doc, replacements):
    for paragraph in doc.paragraphs:
        _replace_paragraph_text(paragraph, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_paragraph_text(paragraph, replacements)


def _thai_document_date(value):
    months = [
        "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม",
        "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม",
        "พฤศจิกายน", "ธันวาคม",
    ]
    return f"{value.day} {months[value.month]} {value.year + 543}"


def _set_table_value(table, row_index, column_index, value, align=WD_ALIGN_PARAGRAPH.CENTER):
    if row_index < len(table.rows) and column_index < len(table.columns):
        _set_cell_text(table.cell(row_index, column_index), value, align=align, size=16)


def _find_table(doc, required_text):
    for table in doc.tables:
        table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
        if str(required_text) in table_text:
            return table
    return None


def _build_exact_procurement_template(purchase):
    """Fill the manually corrected Word master without adding duplicate pages."""
    template_path = Path(__file__).resolve().parent / "templates" / "word" / "purchase_master.docx"
    if not template_path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ต้นแบบ Word: {template_path}")

    doc = Document(str(template_path))
    profile = (
        purchase.government_profile
        or GovernmentProfile.query.filter_by(active=True).first()
        or GovernmentProfile()
    )

    lines = list(purchase.lines)
    item_count = len(lines)
    total = to_decimal(purchase.total_amount)
    subtotal, vat = _vat_values(purchase)
    procurement_type = purchase.procurement_type or "เวชภัณฑ์มิใช่ยา"
    procurement_label = f"{procurement_type} จำนวน {item_count} รายการ"
    thai_date = _thai_document_date(purchase.document_date)
    short_date = purchase.document_date.strftime("%d/%m/%Y")

    budget_allocated = to_decimal(purchase.budget_allocated)
    budget_used = to_decimal(purchase.budget_previously_used)
    budget_this_time = total
    budget_remaining = (budget_allocated - budget_used - budget_this_time).quantize(Decimal("0.01"))

    company = purchase.company

    replacements = [
        ("โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรมโทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙", profile.department),
        ("โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม โทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙", profile.department),
        ("โรงพยาบาลสิงห์บุรี   กลุ่มงานเภสัชกรรม โทร. 03652 2508 ต่อ 1129", profile.department),
        ("สห ๐๐๓๓.๒๐๕.๑๒/", profile.letter_prefix),
        ("กรกฎาคม   2569", thai_date),
        ("กรกฎาคม ๒๕๖๙", thai_date),
        ("กรกฎาคม 2569", thai_date),

        ("เวชภัณฑ์มิใช่ยา จำนวน 10 รายการ", procurement_label),
        ("เวชภัณฑ์มิใช่ยา จำวน 10 รายการ", procurement_label),
        ("เวชภัณฑ์มิใช่ยา จำนวน 2 รายการ", procurement_label),
        ("เวชภัณฑ์มิใช่ยา จำนวน2 รายการ", procurement_label),
        ("เวชภัณฑ์มิใช่ยา จำนวน 1 รายการ", procurement_label),
        ("เวชภัณฑ์มิใช่ยา จำวน 1 รายการ", procurement_label),

        ("บริษัท พี.เอ็น.โปรดักส์ นครสวรรค์ จำกัด", company.name),
        ("บริษัท พี.เอ็น.โปรดักส์ นครสวรรค์จำกัด", company.name),
        ("๐๕๖๒๒๒๑๑๒", company.phone or "-"),
        ("๐๖๐๕๕๒๒๐๐๐๗๙๗", company.tax_id or "-"),
        ("ธนาคารกรุงไทยจำกัด (มหาชน)", company.bank_name or "-"),
        ("ปากน้ำโพ", company.bank_branch or "-"),
        ("๖๒๘๑๒๘๐๙๘๑", company.account_no or "-"),
        ("พี.เอ็น.โปรดักส์ นครสวรรค์ หจก.", company.account_name or "-"),

        ("PO-๖๙-0๒00๗๘", purchase.po_number),
        ("25/07/2026", short_date),
        ("69079275357", purchase.project_number or "........................"),
        ("690714258286", purchase.contract_control_number or "........................"),

        ("30,000.00", f"{total:,.2f}"),
        ("30000", f"{total:,.2f}"),
        ("สามหมื่นบาทถ้วน", baht_text(total)),
        ("28,037.38", f"{subtotal:,.2f}"),
        ("1,962.62", f"{vat:,.2f}"),

        ("1,000.00", f"{budget_allocated:,.2f}"),
        ("1000.00", f"{budget_allocated:,.2f}"),
        ("คำนวณอัตโนมัติ", f"{budget_this_time:,.2f}"),
        ("คำนวนอัตโนมัติ", f"{budget_this_time:,.2f}"),

        ("ใช้ในการรักษาผู้ป่วย", purchase.necessity_reason or "ใช้ในการรักษาผู้ป่วย"),
        ("เงินนอกงบประมาณจาก เงินบำรุงโรงพยาบาลสิงห์บุรี ปี ๒๕๖๙", purchase.budget_source or ""),
        ("โรงพยาบาลสิงห์บุรี ๙๑๗/๓", purchase.delivery_place or "โรงพยาบาลสิงห์บุรี ๙๑๗/๓"),

        ("นางพิณนภา ศริพันธุ์", profile.officer_name),
        ("นายชัชวาลย์ บุญญฤทธิ์", profile.chief_name),
        ("นายชัชวาล บุญญฤทธิ์", profile.chief_name),
        ("นายพิรุณ ปิตะหงษ์นันท์", profile.approver_name),
        ("นางสาวกัญญพัชร ธนกิจการค้า", profile.inspector1_name),
        ("นางสาวชุลีพร สุขมี", profile.inspector2_name),
        ("นางสาวกัญญาพัชร เลิศอนันตกูล", profile.inspector3_name),
        ("นางสาวนลินี เครือทิวา", profile.specifier_name),
    ]
    _replace_in_document(doc, replacements)

    # Update all product tables while retaining the template's layout.
    for table in doc.tables:
        header = " | ".join(cell.text.strip() for cell in table.rows[0].cells)
        if not ("รายการ" in header and "จำนวน" in header and ("หน่วยละ" in header or "ราคาต่อหน่วย" in header)):
            continue

        product_row_count = min(6, max(0, len(table.rows) - 1))
        for index in range(product_row_count):
            line = lines[index] if index < item_count else None
            row_index = index + 1

            if len(table.columns) == 5:
                values = [
                    str(index + 1),
                    line.description if line else "",
                    f"{line.quantity:g} {line.unit.name}" if line else "",
                    f"{line.unit_price:,.2f}" if line else "",
                    f"{line.amount:,.2f}" if line else "",
                ]
            else:
                values = [
                    str(index + 1),
                    line.description if line else "",
                    f"{line.quantity:g}" if line else "",
                    line.unit.name if line else "",
                    f"{line.unit_price:,.2f}" if line else "",
                    f"{line.amount:,.2f}" if line else "",
                ]

            for column_index, value in enumerate(values):
                if column_index < len(table.columns):
                    align = WD_ALIGN_PARAGRAPH.LEFT if column_index == 1 else WD_ALIGN_PARAGRAPH.CENTER
                    _set_table_value(table, row_index, column_index, value, align)

        for row_index, row in enumerate(table.rows):
            row_text = " ".join(cell.text for cell in row.cells)
            if "รวมเป็นเงินทั้งสิ้น" in row_text:
                _set_table_value(table, row_index, len(table.columns) - 1, f"{total:,.2f}", WD_ALIGN_PARAGRAPH.RIGHT)
                if len(table.columns) >= 5:
                    _set_table_value(table, row_index, 1, f"({baht_text(total)})", WD_ALIGN_PARAGRAPH.CENTER)

    # Replace seller information in the purchase-order box.
    seller_table = _find_table(doc, "ผู้ขาย")
    if seller_table is not None and seller_table.rows:
        left_text = (
            f"ผู้ขาย {company.name}\n"
            f"ที่อยู่ {company.address or '-'}\n"
            f"โทรศัพท์ {company.phone or '-'}\n"
            f"เลขประจำตัวผู้เสียภาษี {company.tax_id or '-'}\n"
            f"เลขที่บัญชีเงินฝากธนาคาร {company.account_no or '-'}\n"
            f"ชื่อบัญชี {company.account_name or '-'}\n"
            f"ธนาคาร {company.bank_name or '-'}"
            + (f" สาขา {company.bank_branch}" if company.bank_branch else "")
        )
        right_text = (
            f"เลขที่ {purchase.po_number}\n"
            f"วันที่ {short_date}\n\n"
            "ส่วนราชการ โรงพยาบาลสิงห์บุรี\n"
            "ที่อยู่ ๙๑๗/๓ ตำบลบางพุทรา อำเภอเมืองสิงห์บุรี จังหวัดสิงห์บุรี ๑๖๐๐๐\n"
            "โทรศัพท์ ๐๓๖-๕๒๒๕๐๗"
        )
        _set_cell_text(seller_table.cell(0, 0), left_text, size=16)
        if len(seller_table.columns) > 1:
            _set_cell_text(seller_table.cell(0, 1), right_text, size=16)

    # Fill the budget table with calculated values.
    budget_table = _find_table(doc, "ยอดที่ได้รับ")
    if budget_table is not None and len(budget_table.rows) >= 2:
        values = [budget_allocated, budget_used, budget_this_time, budget_remaining]
        for column_index, value in enumerate(values):
            if column_index < len(budget_table.columns):
                _set_table_value(
                    budget_table,
                    1,
                    column_index,
                    f"{value:,.2f}",
                    WD_ALIGN_PARAGRAPH.CENTER,
                )

    return doc


def _build_word(purchase, form_type):
    doc = Document()
    _configure_document(doc)

    if form_type in {"po_single", "po_multiple", "po"}:
        _add_purchase_order(doc, purchase)
    elif form_type == "spec":
        _add_spec(doc, purchase)
    elif form_type in {"receipt_single", "receipt_two", "receipt_multiple", "receipt"}:
        _add_acceptance_receipt(doc, purchase)
    elif form_type == "integrity":
        _add_integrity_form(doc, purchase)
    elif form_type == "procurement_pack":
        doc = _build_exact_procurement_template(purchase)
    elif form_type == "all":
        # The corrected master already contains the complete document pack.
        # Do not append duplicate pages.
        doc = _build_exact_procurement_template(purchase)
    else:
        raise ValueError("ไม่พบรูปแบบ Word")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


@main_bp.route("/purchases/<int:purchase_id>/word/<form_type>")
def export_word(purchase_id, form_type):
    purchase = Purchase.query.get_or_404(purchase_id)
    allowed = {
        "po_single", "po_multiple", "spec",
        "receipt_single", "receipt_two", "receipt_multiple",
        "integrity", "procurement_pack", "all",
        # backwards-compatible names
        "po", "receipt",
    }
    if form_type not in allowed:
        return "ไม่พบแบบฟอร์ม", 404

    output = _build_word(purchase, form_type)
    suffixes = {
        "po_single": "ใบสั่งซื้อ_1รายการ",
        "po_multiple": "ใบสั่งซื้อ_หลายรายการ",
        "po": "ใบสั่งซื้อ",
        "spec": "แบบกำหนดสเปค",
        "receipt_single": "ใบตรวจรับ_1รายการ",
        "receipt_two": "ใบตรวจรับ_2รายการ",
        "receipt_multiple": "ใบตรวจรับ_มากกว่า2รายการ",
        "receipt": "ใบตรวจรับ",
        "integrity": "แบบแสดงความบริสุทธิ์ใจ",
        "procurement_pack": "ชุดรายงานจัดซื้อ",
        "all": "เอกสารทั้งหมด",
    }
    filename = safe_filename(f"{purchase.po_number}_{suffixes[form_type]}.docx")
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        max_age=0,
    )
