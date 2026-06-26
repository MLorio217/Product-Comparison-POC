import re
import io
import shutil
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Product Comparison Generator",
    page_icon="📊",
    layout="centered"
)

# ─────────────────────────────────────────────
# PDF PARSER
# ─────────────────────────────────────────────
def clean_line(line):
    return re.sub(r"\s+", " ", str(line).strip())

def is_noise(line):
    lowered = line.lower()
    noise = [
        "demo notice", "this document is", "does not represent",
        "fictional", "not affiliated", "testing only", "software testing",
        "proposal for ", "benefit schedule", "accident benefits",
        "critical illness benefits", "important notice", "page ",
        "table of contents", "in $", "employee",
    ]
    return any(t in lowered for t in noise)

def looks_like_amount(text):
    text = text.strip()
    return bool(re.match(r"^\$?\d{1,3}(,\d{3})*(\.\d+)?(%|/\w+)?$", text))

def parse_amount(text):
    text = text.strip().replace("$", "").replace(",", "")
    text = re.sub(r"/\w+$", "", text)
    try:
        val = float(text)
        return int(val) if val == int(val) else val
    except ValueError:
        return text

def extract_benefits(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    all_lines = []
    for page in doc:
        for raw in page.get_text().splitlines():
            line = clean_line(raw)
            if line and not is_noise(line):
                all_lines.append(line)
    doc.close()

    benefits = {}
    i = 0
    while i < len(all_lines):
        current = all_lines[i]
        if i + 1 < len(all_lines):
            nxt = all_lines[i + 1]
            if looks_like_amount(nxt):
                benefits[current] = parse_amount(nxt)
                i += 2
                continue
        m = re.match(
            r"^(.+?)\s+(\$?\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:%|/\w+)?|\d+%)$",
            current
        )
        if m:
            name = clean_line(m.group(1))
            raw_val = m.group(2).strip()
            if name and not is_noise(name) and len(name) > 3:
                benefits[name] = raw_val if "%" in raw_val else parse_amount(raw_val)
                i += 1
                continue
        i += 1
    return benefits

# ─────────────────────────────────────────────
# PRODUCT ROUTER
# ─────────────────────────────────────────────
ACC_TERMS = [
    "ambulance", "emergency room", "urgent care", "hospital admission",
    "hospital confinement", "icu", "intensive care unit", "critical care unit",
    "surgery", "coma", "transport", "transportation", "lodging", "x-ray",
    "xray", "x ray", "medical equipment", "speech therapy", "physical therapy",
    "occupational therapy", "rehabilitation therapy", "physical rehabilitation",
    "diagnostic imaging", "diagnostic exams", "major diagnostic",
    "blood", "plasma", "platelets", "doctor visit", "physician visit",
    "initial doctor", "initial physician", "prescription",
    "prosthetic", "ground emergency", "air emergency",
    "ground ambulance", "air ambulance",
]

CI_TERMS = [
    "heart attack", "stroke", "cancer", "invasive cancer",
    "non-invasive cancer", "non invasive cancer", "coronary artery bypass",
    "coronary bypass", "major organ transplant", "organ transplant",
    "kidney failure", "renal failure", "alzheimer", "hiv", "hepatitis",
    "recurrence", "carcinoma", "myocardial infarction",
]

def classify(name):
    text = str(name).lower()
    for t in CI_TERMS:
        if t in text:
            return "CI"
    for t in ACC_TERMS:
        if t in text:
            return "ACC"
    return "UNKNOWN"

def route(benefits):
    result = {"ACC": {}, "CI": {}, "UNKNOWN": {}}
    for name, value in benefits.items():
        result[classify(name)][name] = value
    return result

# ─────────────────────────────────────────────
# ACC CELL MAP
# ─────────────────────────────────────────────
ACC_CELL_MAP = {
    "surgery (open abdominal, thoracic)":                    (7,  True, True),
    "surgery (exploratory or without repair)":               (8,  True, True),
    "blood, plasma,  platelets":                             (10, True, True),
    "blood, plasma, platelets":                              (10, True, True),
    "hospital admission":                                    (11, True, True),
    "hospital confinement (per day up to 365 days)":         (12, True, False),
    "critical care unit (ccu) admission":                    (13, True, True),
    "critical care unit confinement (per day up to 30 days)":(14, True, False),
    "induced coma (up to 14 days)":                          (17, True, False),
    "non-induced coma (duration of 14 or more days)":        (18, True, False),
    "transportation (per trip up to 3 per accident)":        (19, True, False),
    "lodging (per day up to 30 days)":                       (20, True, False),
    "initial doctor visit":                                  (26, True, True),
    "urgent care facility treatment":                        (27, True, True),
    "emergency room treatment":                              (28, True, True),
    "ground ambulance":                                      (29, True, True),
    "air ambulance":                                         (30, True, True),
    "prescription medicine":                                 (34, True, True),
    "medical equipment":                                     (35, True, True),
    "physical or occupational therapy (per treatment up to 10)": (36, True, True),
    "speech therapy (per treatment up to 10)":               (37, True, True),
    "prosthetic device (one)":                               (39, True, False),
    "prosthetic device  (two or more)":                      (40, True, False),
    "major diagnostic exams":                                (41, True, True),
    "x-ray":                                                 (44, True, True),
}

SYNONYMS = {
    "initial physician visit": "initial doctor visit",
    "physician visit": "initial doctor visit",
    "urgent care facility": "urgent care facility treatment",
    "urgent care treatment": "urgent care facility treatment",
    "emergency room": "emergency room treatment",
    "ground emergency transport": "ground ambulance",
    "air emergency transport": "air ambulance",
    "icu admission": "critical care unit (ccu) admission",
    "icu confinement": "critical care unit confinement (per day up to 30 days)",
    "critical care unit admission": "critical care unit (ccu) admission",
    "critical care unit confinement": "critical care unit confinement (per day up to 30 days)",
    "open abdominal surgery": "surgery (open abdominal, thoracic)",
    "exploratory surgery": "surgery (exploratory or without repair)",
    "physical rehabilitation therapy": "physical or occupational therapy (per treatment up to 10)",
    "physical therapy": "physical or occupational therapy (per treatment up to 10)",
    "speech therapy": "speech therapy (per treatment up to 10)",
    "prescription medication": "prescription medicine",
    "major diagnostic imaging": "major diagnostic exams",
    "blood / plasma / platelets": "blood, plasma, platelets",
    "transportation": "transportation (per trip up to 3 per accident)",
    "lodging": "lodging (per day up to 30 days)",
    "non-induced coma": "non-induced coma (duration of 14 or more days)",
    "induced coma": "induced coma (up to 14 days)",
    "hospital confinement": "hospital confinement (per day up to 365 days)",
    "prosthetic device (two or more)": "prosthetic device  (two or more)",
}

def resolve_key(name):
    norm = str(name).lower().strip()
    if norm in ACC_CELL_MAP:
        return norm
    if norm in SYNONYMS and SYNONYMS[norm] in ACC_CELL_MAP:
        return SYNONYMS[norm]
    for syn_key, syn_val in SYNONYMS.items():
        if syn_key in norm and syn_val in ACC_CELL_MAP:
            return syn_val
    return None

# ─────────────────────────────────────────────
# WORKBOOK WRITER
# ─────────────────────────────────────────────
def write_workbook(template_bytes, voya_acc, competitor_acc, review_items):
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(template_bytes)
    tmp.flush()
    tmp_path = Path(tmp.name)
    tmp.close()

    wb = load_workbook(str(tmp_path))
    ws = wb["ACC 2.3 Comparison"]

    writes = []
    skipped = []

    def write_acc(benefits, col):
        for name, value in benefits.items():
            key = resolve_key(name)
            if key is None:
                skipped.append((name, value, "No cell map match"))
                continue
            row, d_ok, h_ok = ACC_CELL_MAP[key]
            ok = d_ok if col == "D" else h_ok
            if not ok:
                skipped.append((name, value, f"{col} not writable for row {row}"))
                continue
            ws[f"{col}{row}"] = value
            writes.append((row, col, name, value))

    write_acc(voya_acc, "D")
    write_acc(competitor_acc, "H")

    # Review Needed sheet
    review_sheet_name = "Review Needed"
    if review_sheet_name in wb.sheetnames:
        ws_r = wb[review_sheet_name]
        for row in ws_r.iter_rows():
            for cell in row:
                cell.value = None
    else:
        ws_r = wb.create_sheet(review_sheet_name)

    headers = ["Source", "Product", "PDF Benefit", "PDF Value", "Reason"]
    for col, h in enumerate(headers, 1):
        c = ws_r.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)

    for idx, item in enumerate(review_items, 2):
        ws_r.cell(row=idx, column=1, value=item.get("source", ""))
        ws_r.cell(row=idx, column=2, value=item.get("product", ""))
        ws_r.cell(row=idx, column=3, value=item.get("pdf_benefit", ""))
        ws_r.cell(row=idx, column=4, value=str(item.get("pdf_value", "")))
        ws_r.cell(row=idx, column=5, value=item.get("reason", ""))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    tmp_path.unlink(missing_ok=True)
    return out.getvalue(), writes, skipped

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📊 Product Comparison Generator")
st.caption("SAFE MODE · ACC 2.3 tab only · CI and Hospital tabs are never modified")

st.markdown("---")

voya_file = st.file_uploader("1. Current / Voya Proposal (PDF)", type=["pdf"])
competitor_file = st.file_uploader("2. Competitor Proposal (PDF)", type=["pdf"])
template_file = st.file_uploader("3. Product Comparison Template (xlsx)", type=["xlsx"])

st.markdown("---")

if st.button("Generate Completed Comparison", type="primary", use_container_width=True):
    if not voya_file:
        st.error("Upload the Current / Voya Proposal PDF.")
    elif not competitor_file:
        st.error("Upload the Competitor Proposal PDF.")
    elif not template_file:
        st.error("Upload the Product Comparison Template.")
    else:
        with st.spinner("Processing..."):
            log = st.container()

            with log:
                st.write("📄 Reading Voya PDF...")
                voya_raw = extract_benefits(voya_file.read())
                st.write(f"✅ Voya: {len(voya_raw)} items found")

                st.write("📄 Reading Competitor PDF...")
                comp_raw = extract_benefits(competitor_file.read())
                st.write(f"✅ Competitor: {len(comp_raw)} items found")

                voya_routed = route(voya_raw)
                comp_routed = route(comp_raw)

                st.write(
                    f"📦 Voya routed — ACC: {len(voya_routed['ACC'])}  "
                    f"CI: {len(voya_routed['CI'])}  "
                    f"Unknown: {len(voya_routed['UNKNOWN'])}"
                )
                st.write(
                    f"📦 Competitor routed — ACC: {len(comp_routed['ACC'])}  "
                    f"CI: {len(comp_routed['CI'])}  "
                    f"Unknown: {len(comp_routed['UNKNOWN'])}"
                )

                review_items = []
                for name, val in voya_routed["CI"].items():
                    review_items.append({"source": "Voya", "product": "CI", "pdf_benefit": name, "pdf_value": val, "reason": "CI not mapped yet"})
                for name, val in voya_routed["UNKNOWN"].items():
                    review_items.append({"source": "Voya", "product": "UNKNOWN", "pdf_benefit": name, "pdf_value": val, "reason": "Could not classify"})
                for name, val in comp_routed["CI"].items():
                    review_items.append({"source": "Competitor", "product": "CI", "pdf_benefit": name, "pdf_value": val, "reason": "CI not mapped yet"})
                for name, val in comp_routed["UNKNOWN"].items():
                    review_items.append({"source": "Competitor", "product": "UNKNOWN", "pdf_benefit": name, "pdf_value": val, "reason": "Could not classify"})

                st.write("✍️ Writing to workbook...")
                out_bytes, writes, skipped = write_workbook(
                    template_file.read(),
                    voya_routed["ACC"],
                    comp_routed["ACC"],
                    review_items
                )

                st.write(f"✅ ACC values written: {len(writes)}")
                st.write(f"⚠️ Skipped: {len(skipped)}")
                st.write(f"📋 Review Needed rows: {len(review_items)}")

        st.success("Done! Download your completed workbook below.")
        st.download_button(
            label="⬇️ Download Completed Comparison Workbook",
            data=out_bytes,
            file_name="Completed_Product_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
