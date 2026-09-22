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
    h, m, s, f = map(int, tc.split(":"))
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

    return f"{new_h:02}:{new_m:02}:{new_s:02}:{new_f:02}"


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
    current_tc = None

    for para in paragraphs:
        text = para.strip()
        if not text:
            continue

        # 🔍 Detect timecode anywhere in line
        tc_match = re.search(r"\d{2}:\d{2}:\d{2}:\d{2}", text)

        if tc_match:
            current_tc = tc_match.group()
            # remove timecode from text if same line
            text = text.replace(current_tc, "").strip()

            if text:
                speaker = detect_speaker(text)
                entries.append((current_tc, speaker, text))

        elif current_tc:
            speaker = detect_speaker(text)
            entries.append((current_tc, speaker, text))

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

    for i, (tc, speaker, text) in enumerate(entries, start=1):
        new_tc = apply_offset(tc, offset)

        # 🔥 Broadcast style (matches your format)
        doc.add_paragraph(f"{media_name} {i}")
        doc.add_paragraph(new_tc)
        doc.add_paragraph(f"{speaker}: {text}")
        doc.add_paragraph("")  # spacing

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
