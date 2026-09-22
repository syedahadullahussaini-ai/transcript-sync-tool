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
# ✅ TIMECODE OFFSET
# ==============================
def offset_timecode(tc, offset, fps):
    parts = tc.split(":")

    # Fix missing frames
    if len(parts) == 3:
        parts.append("00")

    try:
        h, m, s, f = map(int, parts)
    except:
        return tc  # return original if bad format

    oh, om, os, of = offset

    total_frames = ((h*3600 + m*60 + s) * fps + f)
    offset_frames = ((oh*3600 + om*60 + os) * fps + of)

    new_total = total_frames + offset_frames

    nh = new_total // (3600 * fps)
    nm = (new_total % (3600 * fps)) // (60 * fps)
    ns = (new_total % (60 * fps)) // fps
    nf = new_total % fps

    return f"{nh:02}:{nm:02}:{ns:02}:{nf:02}"


# ==============================
# ✅ SPEAKER DETECTION
# ==============================
def detect_speaker(text):
    match = re.match(r"^([A-Za-z :]+):", text)
    if match:
        return match.group(1).strip()
    return "UNKNOWN"


# ==============================
# ✅ EXTRACT SEGMENTS
# ==============================
def extract_segments(paragraphs):
    entries = []
    current_tc = None

    for line in paragraphs:
        text = line.strip()
        if not text:
            continue

        # Detect timecode
        if re.match(r"\d{2}:\d{2}:\d{2}", text):
            current_tc = text
            continue

        if current_tc:
            speaker = detect_speaker(text)
            entries.append((current_tc, speaker, text))
            current_tc = None

    return entries


# ==============================
# ✅ QC CHECKS
# ==============================
def run_qc_checks(entries):
    errors = []
    prev_time = None

    for i, (tc, speaker, text) in enumerate(entries):
        parts = tc.split(":")

        if len(parts) == 3:
            parts.append("00")

        try:
            h, m, s, f = map(int, parts)
        except:
            errors.append(f"Line {i+1}: Invalid timecode")
            continue

        current_time = h*3600 + m*60 + s

        if prev_time is not None and current_time < prev_time:
            errors.append(f"Line {i+1}: Time overlap detected")

        prev_time = current_time

        if not speaker or speaker.strip().upper() == "UNKNOWN":
            errors.append(f"Line {i+1}: Speaker not detected")

        if not text or not text.strip():
            errors.append(f"Line {i+1}: Empty text")

    return errors


# ==============================
# ✅ BUILD OUTPUT (EXACT FORMAT)
# ==============================
def build_output_doc(entries, media_name, offset, fps):
    doc = Document()

    # HEADER
    doc.add_paragraph(f"Transcription Media #{media_name}")
    doc.add_paragraph(f"MEDIA #: {media_name}")
    doc.add_paragraph("")

    for raw_tc, speaker, text in entries:

        # Apply offset
        new_tc = offset_timecode(raw_tc, offset, fps)

        # Remove frames
        new_tc = new_tc.rsplit(":", 1)[0]

        speaker_clean = speaker.strip() if speaker else "UNKNOWN"

        # Remove duplicate speaker in text
        if ":" in text:
            parts = text.split(":", 1)
            if len(parts) > 1 and parts[0].strip().upper() == speaker_clean.strip().upper():
                text = parts[1].strip()

        # Clean speaker format
        speaker_clean = speaker_clean.replace(":", " ")

        # FORMAT BLOCK
        doc.add_paragraph(f"[{media_name}]")
        doc.add_paragraph(f"[{new_tc}]")
        doc.add_paragraph(f"[{speaker_clean}]:{text}")

        # spacing
        doc.add_paragraph("")

    return doc


# ==============================
# ✅ MAIN FUNCTION
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

    out_doc = build_output_doc(entries, media_name, offset, fps)
    out_doc.save(output_path)

    return {"success": True}
