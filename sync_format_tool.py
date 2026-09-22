import re
from docx import Document

# ==============================
# ✅ OFFSET VALIDATION
# ==============================
def validate_offset(offset_str):
    pattern = r"^\d{2}:\d{2}:\d{2}:\d{2}$"
    return re.match(pattern, offset_str)


# ==============================
# ✅ PARSE OFFSET
# ==============================
def parse_offset(offset_str):
    h, m, s, f = map(int, offset_str.split(":"))
    return h, m, s, f


# ==============================
# ✅ TIMECODE FIX (AUTO CORRECTION)
# ==============================
def apply_offset(tc, offset):
    parts = tc.split(":")

    # ✅ Fix missing frames (HH:MM:SS → HH:MM:SS:00)
    if len(parts) == 3:
        parts.append("00")

    h, m, s, f = map(int, parts)

    oh, om, os, of = offset

    total_frames = (h*3600 + m*60 + s)*25 + f
    offset_frames = (oh*3600 + om*60 + os)*25 + of

    new_total = total_frames + offset_frames

    new_s = new_total // 25
    new_f = new_total % 25

    new_h = new_s // 3600
    new_s %= 3600
    new_m = new_s // 60
    new_s %= 60

    return f"{new_h}:{new_m:02}:{new_s:02}:{new_f:02}"


# ==============================
# ✅ SPEAKER AUTO DETECTION
# ==============================
def detect_speaker(line):
    match = re.match(r"^([A-Za-z ]+):", line)
    if match:
        return match.group(1).strip()
    return "UNKNOWN"


# ==============================
# ✅ EXTRACT SEGMENTS
# ==============================
def extract_segments(paragraphs):
    entries = []
    buffer_text = None

    for line in paragraphs:
        text = line.strip()
        if not text:
            continue

        # detect timecode
        tc_match = re.match(r"\d{2}:\d{2}:\d{2}", text)

        if tc_match:
            if buffer_text:
                speaker = detect_speaker(buffer_text)
                entries.append((text, speaker, buffer_text))
                buffer_text = None
        else:
            buffer_text = text

    return entries

# ==============================
# ✅ QC CHECKS
# ==============================
def run_qc_checks(entries):
    errors = []
    prev_time = None

    for i, (tc, speaker, text) in enumerate(entries):
        h, m, s, f = map(int, tc.split(":"))
        current_time = h*3600 + m*60 + s

        if prev_time is not None and current_time < prev_time:
            errors.append(f"Line {i+1}: Time overlap detected")

        if not speaker or speaker == "UNKNOWN":
            errors.append(f"Line {i+1}: Speaker not detected")

        prev_time = current_time

    return errors


# ==============================
# ✅ BUILD OUTPUT (BROADCAST STYLE)
# ==============================
def build_output(entries, media_name, offset):
    doc = Document()

    for tc, speaker, text in entries:
        new_tc = apply_offset(tc, offset)

        doc.add_paragraph(f"[{media_name}]")
        doc.add_paragraph(f"[{new_tc}]")
        doc.add_paragraph(f"[{speaker}] {text}")
    
    return doc


# ==============================
# ✅ MAIN RUN FUNCTION
# ==============================
def run(input_path, output_path, media_name, offset_str, fps=25):
    if not validate_offset(offset_str):
        return {"errors": ["Invalid offset format. Use HH:MM:SS:FF"]}

    offset = parse_offset(offset_str)

    doc = Document(input_path)
    paragraphs = [p.text for p in doc.paragraphs]

    entries = extract_segments(paragraphs)

    if not entries:
        return {"errors": ["No valid transcript entries found"]}

    qc_errors = run_qc_checks(entries)
    if qc_errors:
        return {"errors": qc_errors}

    out_doc = build_output(entries, media_name, offset)
    out_doc.save(output_path)

    return {"success": True}
