"""
PDF Parser - Updated
Handles two-line format (name / amount on separate lines)
and same-line format (name $amount).
"""

import re
import fitz


NOISE_TERMS = [
    "fictional", "parser testing", "proposal for",
    "accident hospital care", "accident care", "common injuries",
    "dislocations closed", "fractures closed", "additional benefits",
    "monthly rates", "closed reduction", "open reduction",
    "testing notice", "plan highlights", "human review",
    "administrative notes", "testing notes", "proposal id",
    "effective date", "prepared for", "guaranteed issue",
    "portable coverage", "employee-paid", "rate guarantee",
    "this document", "not intended",
    "summit insurance solutions", "apex benefits group",
    "demo notice", "does not represent",
    "demo notice", "entirely fictional",
]

STANDALONE_SKIP = {"benefit", "amount", "benefit amount", "employee benefit", "employee"}


def _clean(line):
    return re.sub(r"\s+", " ", str(line).strip())


def _is_noise(line):
    low = line.lower().strip()
    if low in STANDALONE_SKIP:
        return True
    return any(t in low for t in NOISE_TERMS) or len(line) < 2


def _looks_like_amount(text):
    """Match dollar amounts with or without commas, with optional /day etc."""
    return bool(re.match(r"^\$?\d+(,\d{3})*(\.\d+)?(%|/\w+)?$", text.strip()))


def _parse_amount(text):
    text = text.strip().replace("$", "").replace(",", "")
    text = re.sub(r"/\w+$", "", text)
    if "%" in text:
        return text + "%"
    try:
        val = float(text)
        return int(val) if val == int(val) else val
    except ValueError:
        return None


def extract_benefits(pdf_input):
    """
    Extract benefit name -> value pairs from a proposal PDF.
    Accepts file path (str/Path) or bytes.
    """
    if isinstance(pdf_input, bytes):
        doc = fitz.open(stream=pdf_input, filetype="pdf")
    else:
        doc = fitz.open(str(pdf_input))

    all_lines = []
    for page in doc:
        for raw in page.get_text().splitlines():
            line = _clean(raw)
            if line and not _is_noise(line):
                all_lines.append(line)
    doc.close()

    benefits = {}
    i = 0

    while i < len(all_lines):
        current = all_lines[i]

        # Pattern 1: next line is standalone amount
        if i + 1 < len(all_lines):
            nxt = all_lines[i + 1]
            if _looks_like_amount(nxt):
                val = _parse_amount(nxt)
                if val is not None and len(current) > 3:
                    benefits[current] = val
                i += 2
                continue

        # Pattern 2: same line ends with amount
        m = re.match(
            r"^(.+?)\s+(\$?\d+(?:,\d{3})*(?:\.\d+)?(?:%|/\w+)?)$",
            current
        )
        if m:
            name = _clean(m.group(1))
            raw_val = m.group(2).strip()
            if name and not _is_noise(name) and len(name) > 3:
                val = _parse_amount(raw_val)
                if val is not None:
                    benefits[name] = val
            i += 1
            continue

        i += 1

    return benefits
