import os
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from datetime import date
import io
import requests

st.set_page_config(page_title="Packing List Generator", layout="wide")

# ─────────────────────────────────────────────
# SUPABASE
# ─────────────────────────────────────────────
SUPABASE_URL = "https://lztrggttkgvgjouofibd.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx6dHJnZ3R0a2d2Z2pvdW9maWJkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDAwNzEsImV4cCI6MjA4ODgxNjA3MX0.tbHCQtGW21C2fXCEu2FGwlsXn4kGUWOGoOqjuYyiC7A"
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json"
}

# ─────────────────────────────────────────────
# FORMATTERS
# ─────────────────────────────────────────────
def fmt_weight(value):
    """Italian format: 1.000,– / 1,25 / 136,–"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return ""
    cents = round((f % 1) * 100)
    int_str = f"{int(f):,}".replace(",", ".")
    return f"{int_str},–" if cents == 0 else f"{int_str},{cents:02d}"

# ─────────────────────────────────────────────
# SUPABASE LOADERS
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_fatture():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/fatture",
        headers=HEADERS,
        params={
            "select": "id,invoice_number,client_company,created_at",
            "order": "created_at.desc"
        }
    )
    try:
        d = r.json()
        return d if isinstance(d, list) else []
    except:
        return []

@st.cache_data(ttl=300)
def load_customers():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/customers",
        headers=HEADERS,
        params={
            "select": "company_name,contact_name,salutation,address,city,zip,region,country",
            "order": "company_name.asc"
        }
    )
    try:
        d = r.json()
        return d if isinstance(d, list) else []
    except:
        return []

@st.cache_data(ttl=300)
def load_products():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/products",
        headers=HEADERS,
        params={
            "select": "id,description,description_eng,net_weight_kg,dimensions,category",
            "order": "category.asc,description.asc"
        }
    )
    try:
        d = r.json()
        return d if isinstance(d, list) else []
    except:
        return []

# ─────────────────────────────────────────────
# DOCX HELPERS
# ─────────────────────────────────────────────
def replace_para_text(para, old, new):
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return False
    full = full.replace(old, new)
    if para.runs:
        para.runs[0].text = full
        for r in para.runs[1:]:
            r.text = ""
    return True

def set_para_bold(para, bold):
    for run in para.runs:
        if run.text.strip():
            run.bold = bold

def delete_para(para):
    p = para._p
    p.getparent().remove(p)

def set_cell_text(cell, text, bold=False, font_name="Verdana", font_size=10):
    tc = cell._tc
    paras = tc.findall(qn("w:p"))
    for extra in paras[1:]:
        tc.remove(extra)
    first_p = cell.paragraphs[0]
    for run in first_p.runs:
        run.text = ""
        rPr = run._r.find(qn("w:rPr"))
        if rPr is not None:
            run._r.remove(rPr)
    run = first_p.add_run(text)
    run.bold = bold
    run.font.name = font_name
    run.font.size = Pt(font_size)

def collapse_row(row):
    """Make unused table row height = 1 twip (invisible)."""
    trPr = row._tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        row._tr.insert(0, trPr)
    existing = trPr.find(qn("w:trHeight"))
    if existing is not None:
        trPr.remove(existing)
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), "1")
    trH.set(qn("w:hRule"), "exact")
    trPr.append(trH)

# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
fatture   = load_fatture()
customers = load_customers()
products  = load_products()

# Build lookup maps
customer_map = {c["company_name"]: c for c in customers if c.get("company_name")}
product_map  = {p["id"]: p for p in products}

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📦 Packing List Generator")

# ── 1. SELECT FATTURA ──────────────────────────
st.subheader("1. Link to Fattura")

if not fatture:
    st.warning("No fatture found in Supabase.")
    st.stop()

fattura_labels = [
    f"{f['invoice_number']} — {f['client_company']} ({f['created_at'][:10]})"
    for f in fatture
]
sel_fattura_idx = st.selectbox(
    "Select Fattura",
    range(len(fattura_labels)),
    format_func=lambda i: fattura_labels[i]
)
sel_fattura = fatture[sel_fattura_idx]

invoice_number = sel_fattura.get("invoice_number", "")
client_company = sel_fattura.get("client_company", "")
fattura_date_raw = sel_fattura.get("created_at", "")
try:
    fattura_date = date.fromisoformat(fattura_date_raw[:10]).strftime("%d/%m/%Y")
except:
    fattura_date = fattura_date_raw[:10]

st.info(f"📄 Invoice: **{invoice_number}** | Client: **{client_company}** | Date: **{fattura_date}**")

# ── 2. CLIENT DETAILS ──────────────────────────
st.subheader("2. Client Details")

# Autofill from customers table by company name
cust = customer_map.get(client_company, {})

col1, col2 = st.columns(2)
with col1:
    company  = st.text_input("Company *",       value=cust.get("company_name", client_company))
    address  = st.text_input("Address",         value=cust.get("address", ""))
    zip_code = st.text_input("ZIP",             value=cust.get("zip", ""))
with col2:
    city    = st.text_input("City",             value=cust.get("city", ""))
    region  = st.text_input("Region",           value=cust.get("region", "") or "")
    country = st.text_input("Country",          value=cust.get("country", ""))

include_attn = st.checkbox("Include 'To the attn. of' line?", value=False)
salutation = ""
full_name  = ""
if include_attn:
    a1, a2 = st.columns([1, 3])
    with a1:
        salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Dr.", "Messrs."])
    with a2:
        full_name = st.text_input(
            "Full Name (optional)",
            value=cust.get("contact_name", "") or ""
        )

# ── 3. DIMENSIONS ──────────────────────────────
st.subheader("3. Crate Dimensions")
crate_dimensions = st.text_input(
    "Crate dimensions (cm)",
    value="",
    placeholder="e.g. 120 x 80 x 90"
)

# ── 4. PRODUCTS ────────────────────────────────
st.subheader("4. Products")
st.caption("Select products from catalogue. Net weight auto-fills from database.")

MAX_ROWS = 15
product_labels = ["— empty —"] + [
    f"{p.get('description_eng') or p.get('description', '')} "
    f"({'%.3f' % float(p['net_weight_kg']) if p.get('net_weight_kg') else 'no weight'} kg)"
    for p in products
]
product_list = [None] + products  # index 0 = empty

if "pl_rows" not in st.session_state:
    st.session_state.pl_rows = [
        {"prod_idx": 0, "qty": 1.0, "net_weight": 0.0, "gross_weight": 0.0, "description": "", "dimensions": ""}
    ]

def add_pl_row():
    st.session_state.pl_rows.append(
        {"prod_idx": 0, "qty": 1.0, "net_weight": 0.0, "gross_weight": 0.0, "description": "", "dimensions": ""}
    )

rows_to_remove = []
needs_rerun = False

for i, row in enumerate(st.session_state.pl_rows):
    with st.expander(f"Row {(i+1)*10}", expanded=(i < 4)):
        c1, c2, c3, c4, c5 = st.columns([4, 1, 2, 2, 2])

        with c1:
            prod_idx = st.selectbox(
                "Product", range(len(product_labels)),
                format_func=lambda x: product_labels[x],
                key=f"pl_prod_{i}", index=row["prod_idx"]
            )
            if prod_idx != row["prod_idx"]:
                row["prod_idx"] = prod_idx
                p = product_list[prod_idx]
                if p:
                    row["description"] = p.get("description_eng") or p.get("description", "")
                    nw = p.get("net_weight_kg")
                    row["net_weight"]   = float(nw) if nw is not None else 0.0
                    row["gross_weight"] = row["net_weight"]
                    row["dimensions"]   = p.get("dimensions") or ""
                else:
                    row["description"]  = ""
                    row["net_weight"]   = 0.0
                    row["gross_weight"] = 0.0
                    row["dimensions"]   = ""
                needs_rerun = True

        with c2:
            row["qty"] = st.number_input(
                "Qty", min_value=0.0, value=float(row["qty"]),
                step=1.0, format="%.1f", key=f"pl_qty_{i}"
            )
        with c3:
            row["net_weight"] = st.number_input(
                "Net Weight (kg)", min_value=0.0,
                value=float(row["net_weight"]),
                step=0.01, format="%.3f", key=f"pl_nw_{i}"
            )
        with c4:
            row["gross_weight"] = st.number_input(
                "Gross Weight (kg)", min_value=0.0,
                value=float(row["gross_weight"]),
                step=0.01, format="%.3f", key=f"pl_gw_{i}"
            )
        with c5:
            st.write("")
            st.write("")
            if st.button("🗑", key=f"pl_del_{i}"):
                rows_to_remove.append(i)

        # Show dimensions as info
        p = product_list[prod_idx]
        if p and p.get("dimensions"):
            st.caption(f"📐 Dimensions: {p['dimensions']}")

for i in sorted(rows_to_remove, reverse=True):
    st.session_state.pl_rows.pop(i)
if rows_to_remove or needs_rerun:
    st.rerun()

st.button("➕ Add Row", on_click=add_pl_row)

active_rows = [r for r in st.session_state.pl_rows if r["prod_idx"] > 0 and r["qty"] > 0]
total_net_weight   = sum(r["net_weight"]   for r in active_rows)
total_gross_weight = sum(r["gross_weight"] for r in active_rows)

col_nw, col_gw = st.columns(2)
with col_nw:
    st.metric("Total Net Weight", f"{fmt_weight(total_net_weight)} kg")
with col_gw:
    st.metric("Total Gross Weight", f"{fmt_weight(total_gross_weight)} kg")

# ── GENERATE ───────────────────────────────────
st.divider()

if not company.strip():
    st.warning("⚠️ Company is mandatory.")

if st.button("📦 Generate Packing List", type="primary",
             disabled=not company.strip() or len(active_rows) == 0,
             use_container_width=True):

    try:
        template_path = os.path.join(os.path.dirname(__file__), "Packing_list_template.docx")
        doc = Document(template_path)
        paras = doc.paragraphs

        # ── Build header strings ──
        zip_city = f"{zip_code} {city}".strip()
        if region:
            zip_city += f", {region}"

        # Para 2: Company (bold)
        replace_para_text(paras[2], "[COMPANY NAME]", company.upper())
        set_para_bold(paras[2], True)

        # Para 3: Address (not bold)
        replace_para_text(paras[3], "[Address]", address)
        set_para_bold(paras[3], False)

        # Para 4: Zip City Region (not bold)
        replace_para_text(paras[4], "[Zip] [City], [Region]", zip_city)
        set_para_bold(paras[4], False)

        # Para 5: Country (not bold)
        replace_para_text(paras[5], "[Country]", country)
        set_para_bold(paras[5], False)

        # Para 7: Attn line — delete if not needed
        attn_para = paras[7]
        if include_attn and (salutation or full_name):
            attn_text = f"To the attn. of {salutation} {full_name}".strip()
            attn_text = attn_text.replace("To the attn. of  ", "To the attn. of ")
            replace_para_text(attn_para, "To the attn. of [Sal.] [Full Name]", attn_text)
            set_para_bold(attn_para, False)
        else:
            delete_para(attn_para)

        # Para 10: Invoice ref (re-fetch paras after possible deletion)
        paras = doc.paragraphs
        for para in paras:
            if "[NNN/YY]" in para.text or "[DD/MM/YYYY]" in para.text:
                replace_para_text(para, "[NNN/YY]",      invoice_number)
                replace_para_text(para, "[DD/MM/YYYY]",  fattura_date)
                set_para_bold(para, False)
                break

        # Para with dimensions
        for para in paras:
            if "[dimensions]" in para.text:
                if crate_dimensions.strip():
                    replace_para_text(para, "[dimensions]", crate_dimensions.strip())
                # else leave [dimensions] as-is per spec
                break

        # Para with net weight sum
        for para in paras:
            if "[sum of Net Weight]" in para.text:
                replace_para_text(para, "[sum of Net Weight]", fmt_weight(total_net_weight))
                break

        # GROSS WEIGHT paragraph — leave hardcoded, do not touch

        # ── Product table ──
        table = doc.tables[0]

        for idx in range(MAX_ROWS):
            row_obj = table.rows[idx + 1]  # row 0 = header
            cells = row_obj.cells

            if idx < len(active_rows):
                r = active_rows[idx]
                p = product_list[r["prod_idx"]]

                # Description: use eng name + dimensions if available
                desc = r["description"]
                if p and p.get("dimensions"):
                    desc += f"\n{p['dimensions']}"

                # Qty formatting
                qty_val = r["qty"]
                qty_str = f"{int(qty_val)},0" if qty_val == int(qty_val) else f"{qty_val:.1f}".replace(".", ",")

                set_cell_text(cells[0], qty_str)
                set_cell_text(cells[1], desc)
                set_cell_text(cells[2], "Kg")
                set_cell_text(cells[3], fmt_weight(r["net_weight"]))
                set_cell_text(cells[4], "Kg")
                set_cell_text(cells[5], fmt_weight(r["gross_weight"]))
            else:
                for cell in cells:
                    set_cell_text(cell, "")
                collapse_row(row_obj)

        # ── Save ──
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        default_name = f"PackingList_{invoice_number.replace('/', '-')}_{company.replace(' ', '_')}"
        st.success(f"✅ Packing list generated!")
        st.download_button(
            label="⬇️ Download Packing List",
            data=buf,
            file_name=f"{default_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"❌ Error generating document: {e}")
        st.exception(e)
