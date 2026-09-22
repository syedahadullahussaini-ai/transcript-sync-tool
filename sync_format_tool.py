import re
from docx import Document
from docx.shared import Pt, RGBColor

# ==============================
# ✅ OFFSET VALIDATION
# ==============================
def validate_offset(offset_str):
    pattern = r"^\d{1,2}:\d{2}:\d{2}:\d{2}$"
    return re.match(pattern, offset_str) is not None


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
    """tc: 'HH:MM:SS' or 'HH:MM:SS:FF'. fps is coerced to an int — the
    frame math needs whole frames, and non-integer fps (23.98, 29.97...)
    is treated as its rounded whole value (24, 30...)."""
    fps = int(round(float(fps)))
    parts = tc.split(":")

    if len(parts) == 3:
        parts.append("00")

    try:
        h, m, s, f = map(int, parts)
    except ValueError:
        return tc  # return original if bad format

    oh, om, os_, of = offset

    total_frames = (h * 3600 + m * 60 + s) * fps + f
    offset_frames = (oh * 3600 + om * 60 + os_) * fps + of

    new_total = total_frames + offset_frames

    total_seconds, nf = divmod(new_total, fps)
    nh = (total_seconds // 3600) % 24
    nm = (total_seconds % 3600) // 60
    ns = total_seconds % 60

    # hour unpadded, minute/second/frame zero-padded to 2 digits
    return f"{nh}:{nm:02d}:{ns:02d}:{nf:02d}"


# ==============================
# ✅ SPEAKER DETECTION
# ==============================
SPEAKER_RE_BRACKET = re.compile(r"^\[([^\]]+)\]:?\s*(.*)$")
SPEAKER_RE_PLAIN = re.compile(r"^([A-Z][A-Za-z ]{0,30}):\s*(.*)$")


def detect_speaker(line):
    """Returns (speaker_or_None, remaining_text). speaker is None when the
    line has no explicit 'SPEAKER:' or '[SPEAKER]:' prefix — the caller is
    responsible for carrying the previous speaker forward in that case, so
    a continuation line is never mislabeled UNKNOWN."""
    match = SPEAKER_RE_BRACKET.match(line)
    if match:
        return match.group(1).strip(), match.group(2).strip()

    match = SPEAKER_RE_PLAIN.match(line)
    if match:
        return match.group(1).strip(), match.group(2).strip()

    return None, line


# ==============================
# ✅ EXTRACT SEGMENTS
# ==============================
TIMECODE_LINE_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}(:\d{1,2})?$")


def extract_segments(paragraphs, merge_bundled=True):
    """
    IMPORTANT: in these raw docs, a timecode line comes AFTER the dialogue
    it belongs to, not before it — e.g.:

        JACOB LANDRY: Ready? (NON-INTERVIEW)
        00:00:02

    So we buffer dialogue text as we see it, and only stamp it with a
    timecode once the timecode paragraph actually appears. If more than
    one dialogue paragraph piles up before the next timecode (a line with
    no timecode of its own, immediately followed by one that does have
    one), they get merged into a single entry when merge_bundled=True —
    since the source never timestamped them separately, showing them as
    two identical-timecode blocks would be misleading. Lines that already
    each have their own timecode simply produce separate entries, even if
    two of those timecodes happen to carry the same value.
    """
    entries = []
    buffered = []  # (speaker, text) not yet timed
    current_speaker = None

    for line in paragraphs:
        text = line.strip()
        if not text:
            continue

        if TIMECODE_LINE_RE.match(text):
            if buffered:
                if merge_bundled and len(buffered) > 1:
                    speakers = []
                    for spk, _ in buffered:
                        if spk not in speakers:
                            speakers.append(spk)
                    combined_text = " ".join(t for _, t in buffered)
                    entries.append((text, " / ".join(speakers), combined_text))
                else:
                    for spk, txt in buffered:
                        entries.append((text, spk, txt))
                buffered = []
            continue

        speaker, dialogue = detect_speaker(text)
        if speaker is not None:
            current_speaker = speaker
        # If no speaker was ever established (first line has no prefix),
        # fall back to UNKNOWN rather than crashing/erroring.
        buffered.append((current_speaker or "UNKNOWN", dialogue))

    # any text left in `buffered` at end-of-doc had no trailing timecode;
    # it's intentionally dropped from entries (nothing to stamp it with)
    return entries


# ==============================
# ✅ QC CHECKS  (now non-fatal — returns warnings, never blocks output)
# ==============================
def run_qc_checks(entries):
    warnings = []
    prev_time = None

    for i, (tc, speaker, text) in enumerate(entries):
        parts = tc.split(":")
        if len(parts) == 3:
            parts.append("00")
        try:
            h, m, s, f = map(int, parts)
        except ValueError:
            warnings.append(f"Line {i+1}: Invalid timecode")
            continue

        current_time = h * 3600 + m * 60 + s
        if prev_time is not None and current_time < prev_time:
            warnings.append(f"Line {i+1}: Timecode goes backwards")
        prev_time = current_time

        if not speaker or speaker.strip().upper() == "UNKNOWN":
            warnings.append(f"Line {i+1}: Speaker not detected")

        if not text or not text.strip():
            warnings.append(f"Line {i+1}: Empty text")

    return warnings


# ==============================
# ✅ SPEAKER LABEL FORMATTING
# ==============================
def format_speaker_label(speaker):
    """[Q] gets a trailing colon ('[Q]:'), everything else doesn't
    ('[CREW]', '[JACOB LANDRY]'), matching house style. Merged
    multi-speaker labels ('A / B') are upper-cased piece by piece."""
    parts = [p.strip() for p in speaker.split("/")]
    upped = [p.upper() for p in parts]
    label = " / ".join(upped)
    if label == "Q":
        return f"[Q]:"
    return f"[{label}]"


# ==============================
# ✅ BUILD OUTPUT (matches the validated reference layout)
# ==============================
def build_output_doc(entries, media_name, offset, fps,
                      running_header_label="Transcription Media #",
                      body_label="MEDIA #:"):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Pt(612)
    section.page_height = Pt(792)
    section.left_margin = Pt(144)
    section.right_margin = Pt(144)
    section.top_margin = Pt(72)
    section.bottom_margin = Pt(72)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.text = ""
    hr = hp.add_run(f"{running_header_label} {media_name}")
    hr.bold = True
    hr.font.name = "Arial Bold"
    hr.font.size = Pt(12)
    hr.font.color.rgb = RGBColor(0xC0, 0xC0, 0xC0)
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1.0

    def add_para(text, bold_prefix=None):
        p = doc.add_paragraph()
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.bold = True
            r1.font.name = "Arial"
            r1.font.size = Pt(12)
            r2 = p.add_run(" " + text)
            r2.font.name = "Arial"
            r2.font.size = Pt(12)
        else:
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(12)
        p.paragraph_format.line_spacing = 1.0
        return p

    add_para(media_name, bold_prefix=body_label)

    for raw_tc, speaker, text in entries:
        new_tc = offset_timecode(raw_tc, offset, fps)
        speaker_clean = (speaker or "UNKNOWN").strip()

        # strip a duplicated "SPEAKER: " prefix already inside the text
        if ":" in text:
            head, _, rest = text.partition(":")
            if head.strip().upper() == speaker_clean.upper():
                text = rest.strip()

        add_para(f"[{media_name}]")
        add_para(f"[{new_tc}]")
        label = format_speaker_label(speaker_clean)
        sep = " " if not label.endswith(":") else " "
        add_para(f"{label}{sep}{text}")
        doc.add_paragraph("")

    return doc


# ==============================
# ✅ MAIN FUNCTION
# ==============================
def run(input_path, output_path, media_name, offset_str, fps=25):
    if not validate_offset(offset_str):
        return {"success": False, "errors": ["Invalid offset format. Use HH:MM:SS:FF"]}

    offset = parse_offset(offset_str)

    doc = Document(input_path)
    paragraphs = [p.text for p in doc.paragraphs]

    entries = extract_segments(paragraphs)

    if not entries:
        return {"success": False, "errors": ["No valid transcript entries found"]}

    # QC issues are surfaced as warnings, not a hard stop — a raw doc with
    # an odd line or two shouldn't block the whole file from being produced.
    warnings = run_qc_checks(entries)

    out_doc = build_output_doc(entries, media_name, offset, fps)
    out_doc.save(output_path)

    result = {"success": True, "entries": len(entries)}
    if warnings:
        result["warnings"] = warnings
    return result
