import re
import io
import tempfile
from pathlib import Path

import fitz
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

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
        "table of contents", "in $", "employee", "testing notice",
        "plan highlights", "human review", "administrative notes",
        "testing notes", "proposal id", "effective date", "prepared for",
    ]
    return any(t in lowered for t in noise)

def looks_like_amount(text):
    text = text.strip()
    return bool(re.match(r"^\$?\d{1,3}(,\d{3})*(\.(\d+))?(%|/\w+)?$", text))

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
class ProductRouter:

    def __init__(self):
        self._acc_terms = [
            "ambulance", "emergency room", "urgent care", "hospital admission",
            "hospital confinement", "icu", "intensive care unit", "critical care unit",
            "surgery", "coma", "transport", "transportation", "lodging", "x-ray",
            "xray", "x ray", "medical equipment", "durable medical equipment",
            "speech therapy", "physical therapy", "occupational therapy",
            "rehabilitation therapy", "physical rehabilitation", "rehabilitation facility",
            "diagnostic imaging", "diagnostic exams", "major diagnostic", "advanced imaging",
            "blood", "plasma", "platelets", "doctor visit", "physician visit",
            "initial doctor", "initial physician", "prescription",
            "prosthetic", "ground emergency", "air emergency",
            "ground ambulance", "air ambulance",
            # Common injuries - all ACC
            "fracture", "laceration", "burn", "accidental death",
            "dismemberment", "concussion", "traumatic brain",
            "dislocation", "skin graft", "eye injury", "dental work",
            "puncture wound", "ruptured disk", "tendon", "ligament",
            "rotator cuff", "paralysis", "observation unit",
            "wellness", "home health", "chiropractic", "follow-up doctor",
            "follow-up physician", "physician follow-up",
            "outpatient surgery", "lab service", "pet boarding", "family care",
            "accident medical expense", "accident medical",
        ]

        self._ci_terms = [
            "heart attack", "stroke", "cancer", "invasive cancer",
            "non-invasive cancer", "non invasive cancer", "coronary artery bypass",
            "coronary bypass", "major organ transplant", "organ transplant",
            "kidney failure", "renal failure", "end stage renal",
            "alzheimer", "hiv", "hepatitis", "recurrence",
            "carcinoma", "myocardial infarction", "occupational hiv",
        ]

    def classify(self, benefit_name):
        text = str(benefit_name).lower()
        # CI check first - more specific
        for term in self._ci_terms:
            if term in text:
                return "CI"
        for term in self._acc_terms:
            if term in text:
                return "ACC"
        return "UNKNOWN"

    def split_by_product(self, benefits):
        result = {"ACC": {}, "CI": {}, "UNKNOWN": {}}
        for name, value in benefits.items():
            result[self.classify(name)][name] = value
        return result

# ─────────────────────────────────────────────
# ACC CELL MAP + SYNONYMS + RESOLVER
# ─────────────────────────────────────────────
ACC_CELL_MAP = {
    # Surgery
    "surgery (open abdominal, thoracic)":                       (7,   True, True),
    "surgery (exploratory or without repair)":                  (8,   True, True),
    "general anesthesia":                                       (9,   True, True),

    # Hospital
    "blood, plasma, platelets":                                 (10,  True, True),
    "hospital admission":                                       (11,  True, True),
    "hospital confinement (per day up to 365 days)":            (12,  True, False),
    "critical care unit (ccu) admission":                       (13,  True, True),
    "critical care unit confinement (per day up to 30 days)":   (14,  True, False),
    "rehabilitation facility confinement (per day up to 90 days)": (15, True, True),
    "observation unit stay":                                    (16,  True, True),
    "induced coma (up to 14 days)":                             (17,  True, False),
    "non-induced coma (duration of 14 or more days)":           (18,  True, False),
    "transportation (per trip up to 3 per accident)":           (19,  True, False),
    "lodging (per day up to 30 days)":                          (20,  True, False),
    "pet boarding":                                             (21,  True, True),
    "family care (per child/adult up to 45 days)":              (22,  True, True),

    # Accident Care
    "initial doctor visit":                                     (26,  True, True),
    "urgent care facility treatment":                           (27,  True, True),
    "emergency room treatment":                                 (28,  True, True),
    "ground ambulance":                                         (29,  True, True),
    "air ambulance":                                            (30,  True, True),
    "follow-up doctor treatment":                               (31,  True, True),
    "home health care":                                         (32,  True, True),
    "chiropractic treatment (up to 6 per accident)":            (33,  True, True),
    "prescription medicine":                                    (34,  True, True),
    "medical equipment":                                        (35,  True, True),
    "physical or occupational therapy (per treatment up to 10)":(36,  True, True),
    "speech therapy (per treatment up to 10)":                  (37,  True, True),
    "mental health therapy (per treatment up to 10)":           (38,  True, True),
    "prosthetic device (one)":                                  (39,  True, False),
    "prosthetic device (two or more)":                          (40,  True, False),
    "major diagnostic exams":                                   (41,  True, True),
    "outpatient surgery (once per accident)":                   (42,  True, True),
    "outpatient iv infusion therapy":                           (43,  True, True),
    "x-ray":                                                    (44,  True, True),
    "lab service":                                              (45,  True, True),

    # Burns
    "burns (2nd degree, at least 36% of body)":                 (49,  True, True),
    "burns (3rd degree, at least 2% but less than 4%)":         (50,  True, True),
    "burns (3rd degree, 4% or more)":                           (51,  True, True),
    "skin graft":                                               (52,  True, True),

    # Laceration
    "laceration (treated - no sutures)":                        (59,  True, True),
    "laceration (sutures up to 2 inches)":                      (60,  True, True),
    "laceration (sutures 2 to 6 inches)":                       (61,  True, True),
    "laceration (sutures over 6 inches)":                       (62,  True, True),
    "laceration (sutures)":                                     (63,  True, True),

    # Fractures - generic entry (row 94 = Hip fracture)
    "fracture (hip)":                                           (94,  True, True),

    # Other Common Injuries
    "concussion":                                               (69,  True, True),
    "traumatic brain injury":                                   (70,  True, True),

    # Accidental Death
    "accidental death employee":                                (127, True, True),
    "accidental death spouse":                                  (128, True, True),
    "accidental death children":                                (129, True, True),

    # Additional Benefits
    "wellness benefit":                                         (150, True, True),
}


# -----------------------------------------------------------------------
# SYNONYMS
# Maps incoming PDF benefit names to ACC_CELL_MAP keys
# -----------------------------------------------------------------------
SYNONYMS = {
    # Doctor / physician visits
    "initial physician visit":          "initial doctor visit",
    "physician visit":                  "initial doctor visit",
    "initial doctor visit":             "initial doctor visit",

    # Urgent care
    "urgent care":                      "urgent care facility treatment",
    "urgent care facility":             "urgent care facility treatment",
    "urgent care treatment":            "urgent care facility treatment",

    # Emergency room
    "emergency room":                   "emergency room treatment",
    "er treatment":                     "emergency room treatment",

    # Ambulance
    "ground emergency transport":       "ground ambulance",
    "air emergency transport":          "air ambulance",

    # ICU
    "icu admission":                    "critical care unit (ccu) admission",
    "icu confinement":                  "critical care unit confinement (per day up to 30 days)",
    "critical care unit admission":     "critical care unit (ccu) admission",
    "critical care unit confinement":   "critical care unit confinement (per day up to 30 days)",

    # Surgery
    "open abdominal surgery":           "surgery (open abdominal, thoracic)",
    "exploratory surgery":              "surgery (exploratory or without repair)",
    "outpatient surgery":               "outpatient surgery (once per accident)",

    # Therapy
    "physical rehabilitation therapy":  "physical or occupational therapy (per treatment up to 10)",
    "physical therapy":                 "physical or occupational therapy (per treatment up to 10)",
    "rehabilitation therapy":           "physical or occupational therapy (per treatment up to 10)",
    "occupational therapy":             "physical or occupational therapy (per treatment up to 10)",
    "speech therapy":                   "speech therapy (per treatment up to 10)",

    # Prescription
    "prescription medication":          "prescription medicine",
    "prescription benefit":             "prescription medicine",

    # Imaging / diagnostics
    "major diagnostic imaging":         "major diagnostic exams",
    "advanced imaging":                 "major diagnostic exams",
    "advanced imaging (ct/mri/pet)":    "major diagnostic exams",
    "diagnostic imaging":               "major diagnostic exams",

    # Equipment
    "medical equipment":                "medical equipment",
    "durable medical equipment":        "medical equipment",
    "dme":                              "medical equipment",

    # Blood
    "blood / plasma / platelets":       "blood, plasma, platelets",
    "blood / plasma":                   "blood, plasma, platelets",
    "blood plasma platelets":           "blood, plasma, platelets",

    # Transportation / lodging
    "transportation":                   "transportation (per trip up to 3 per accident)",
    "lodging":                          "lodging (per day up to 30 days)",

    # Coma
    "induced coma":                     "induced coma (up to 14 days)",
    "non-induced coma":                 "non-induced coma (duration of 14 or more days)",
    "coma":                             "induced coma (up to 14 days)",

    # Hospital confinement
    "hospital confinement":             "hospital confinement (per day up to 365 days)",
    "hospital confinement 365 days max":"hospital confinement (per day up to 365 days)",
    "icu confinement 30 days max":      "critical care unit confinement (per day up to 30 days)",

    # Prosthetics
    "prosthetic device (two or more)":  "prosthetic device (two or more)",
    "prosthetic device two or more":    "prosthetic device (two or more)",

    # Burns - generic
    "burn benefit":                     "burns (2nd degree, at least 36% of body)",
    "burns":                            "burns (2nd degree, at least 36% of body)",

    # Laceration - generic
    "laceration benefit":               "laceration (sutures)",
    "laceration":                       "laceration (sutures)",

    # Accidental death - generic
    "accidental death":                 "accidental death employee",

    # Rehab facility
    "rehabilitation facility":          "rehabilitation facility confinement (per day up to 90 days)",

    # Follow-up
    "follow up doctor":                 "follow-up doctor treatment",
    "follow-up doctor":                 "follow-up doctor treatment",

    # Wellness
    "wellness":                         "wellness benefit",

    "physician follow-up visit":         "follow-up doctor treatment",
    "follow-up physician visit":          "follow-up doctor treatment",

    # Fracture - generic maps to hip (first fracture row, row 94)
    "fracture benefit":                 "fracture (hip)",
    "fracture":                         "fracture (hip)",

    # Accident medical expense -> outpatient surgery as closest match
    "accident medical expense":         "outpatient surgery (once per accident)",
    "accident medical":                 "outpatient surgery (once per accident)",
}

def _normalize(text):
    return str(text).lower().strip()


def _resolve_key(name):
    """Resolve a PDF benefit name to an ACC_CELL_MAP key."""
    norm = _normalize(name)

    # Direct hit
    if norm in ACC_CELL_MAP:
        return norm

    # Synonym hit (exact)
    if norm in SYNONYMS:
        resolved = SYNONYMS[norm]
        if resolved in ACC_CELL_MAP:
            return resolved

    # Partial synonym: check if any synonym key appears in the name
    for syn_key, syn_val in SYNONYMS.items():
        if syn_key in norm:
            if syn_val in ACC_CELL_MAP:
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
            key = _resolve_key(name)
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
            st.write("📄 Reading Voya PDF...")
            voya_raw = extract_benefits(voya_file.read())
            st.write(f"✅ Voya: {len(voya_raw)} items found")

            st.write("📄 Reading Competitor PDF...")
            comp_raw = extract_benefits(competitor_file.read())
            st.write(f"✅ Competitor: {len(comp_raw)} items found")

            router = ProductRouter()
            voya_routed = router.split_by_product(voya_raw)
            comp_routed = router.split_by_product(comp_raw)

            st.write(
                f"📦 Voya — ACC: {len(voya_routed['ACC'])}  "
                f"CI: {len(voya_routed['CI'])}  "
                f"Unknown: {len(voya_routed['UNKNOWN'])}"
            )
            st.write(
                f"📦 Competitor — ACC: {len(comp_routed['ACC'])}  "
                f"CI: {len(comp_routed['CI'])}  "
                f"Unknown: {len(comp_routed['UNKNOWN'])}"
            )

            review_items = []
            for name, val in voya_routed["CI"].items():
                review_items.append({"source": "Voya", "product": "CI", "pdf_benefit": name, "pdf_value": val, "reason": "CI tab not mapped yet"})
            for name, val in voya_routed["UNKNOWN"].items():
                review_items.append({"source": "Voya", "product": "UNKNOWN", "pdf_benefit": name, "pdf_value": val, "reason": "Could not classify"})
            for name, val in comp_routed["CI"].items():
                review_items.append({"source": "Competitor", "product": "CI", "pdf_benefit": name, "pdf_value": val, "reason": "CI tab not mapped yet"})
            for name, val in comp_routed["UNKNOWN"].items():
                review_items.append({"source": "Competitor", "product": "UNKNOWN", "pdf_benefit": name, "pdf_value": val, "reason": "Could not classify"})

            st.write("✍️ Writing to workbook...")
            out_bytes, writes, skipped = write_workbook(
                template_file.read(),
                voya_routed["ACC"],
                comp_routed["ACC"],
                review_items
            )

            for name, val, reason in skipped:
                review_items.append({"source": "ACC unmatched", "product": "ACC", "pdf_benefit": name, "pdf_value": val, "reason": reason})

            st.write(f"✅ ACC values written: {len(writes)}")
            if skipped:
                st.write(f"⚠️ Skipped (sent to Review Needed): {len(skipped)}")
            st.write(f"📋 Review Needed rows: {len(review_items)}")

        st.success("Done! Download your completed workbook below.")
        st.download_button(
            label="⬇️ Download Completed Comparison Workbook",
            data=out_bytes,
            file_name="Completed_Product_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
