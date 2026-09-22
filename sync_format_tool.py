import re
from docx import Document

# ✅ Offset validation
def validate_offset(offset_str):
    pattern = r"^\d{2}:\d{2}:\d{2}:\d{2}$"
    return bool(re.match(pattern, offset_str))


# ✅ QC checks
def run_qc_checks(entries):
    errors = []
    prev_time = None

    for i, (raw_tc, raw_frames, speaker, text) in enumerate(entries):
        try:
            h, m, s = map(int, raw_tc.split(":"))
            current_time = h * 3600 + m * 60 + s

            if prev_time is not None and current_time < prev_time:
                errors.append(f"Line {i+1}: Time overlap detected")

            prev_time = current_time

            if not raw_tc:
                errors.append(f"Line {i+1}: Missing timestamp")

            if not speaker:
                errors.append(f"Line {i+1}: Missing speaker")

        except:
            errors.append(f"Line {i+1}: Invalid timestamp format")

    return errors


# ✅ Dummy extractor (replace later with your real logic if needed)
def extract_segments(paragraphs):
    entries = []
    for para in paragraphs:
        text = para.strip()
        if text:
            entries.append(("00:00:01", 0, "Speaker", text))
    return entries


# ✅ Main function
def run(input_path, output_path, media_name, offset_str, fps=25):

    if not validate_offset(offset_str):
        return {"errors": ["Invalid offset format. Use HH:MM:SS:FF"]}

    doc = Document(input_path)
    paragraphs = [p.text for p in doc.paragraphs]

    entries = extract_segments(paragraphs)

    # QC checks
    qc_errors = run_qc_checks(entries)
    if qc_errors:
        return {"errors": qc_errors}

    # Create output doc
    out_doc = Document()

    for i, (_, _, speaker, text) in enumerate(entries):
        out_doc.add_paragraph(f"{media_name}_{i+1}")
        out_doc.add_paragraph(f"{speaker}: {text}")
        out_doc.add_paragraph("")

    out_doc.save(output_path)

    return {"success": True}
