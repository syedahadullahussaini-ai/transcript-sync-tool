import re
from docx import Document
from docx.shared import Pt, RGBColor

# Matches "SPEAKER: text" / "0:16:13:14" style timecodes, with or without frames.
TIMECODE_RE = re.compile(r'^(\d{1,2}):(\d{2}):(\d{2})(?::(\d{1,2}))?$')
OFFSET_RE = re.compile(r'^(\d{1,2}):(\d{2}):(\d{2}):(\d{1,2})$')


# ---------------------------------------------------------------------------
# Timecode helpers
# ---------------------------------------------------------------------------

def parse_offset(offset_str):
    """Parse a user-supplied 'HH:MM:SS:FF' offset string into (h, m, s, f)."""
    match = OFFSET_RE.match((offset_str or "").strip())
    if not match:
        raise ValueError(
            f"Offset must be in HH:MM:SS:FF format (e.g. 06:16:11:14), got: {offset_str!r}"
        )
    h, m, s, f = (int(x) for x in match.groups())
    return h, m, s, f


def validate_offset(offset_str):
    """Convenience wrapper for app.py to sanity-check the offset field."""
    try:
        parse_offset(offset_str)
        return True, None
    except ValueError as e:
        return False, str(e)


def timecode_to_frames(h, m, s, f, fps):
    return ((h * 3600 + m * 60 + s) * fps) + f


def frames_to_timecode(frames, fps):
    total_seconds, f = divmod(frames, fps)
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}:{f:02d}"


# ---------------------------------------------------------------------------
# Parsing the raw input transcript
# ---------------------------------------------------------------------------

def _split_into_blocks(paragraph_texts):
    """
    Groups paragraph text into blocks of consecutive non-blank lines,
    separated by one or more blank paragraphs.
    """
    blocks = []
    cur = []
    for raw in paragraph_texts:
        line = raw.strip()
        if not line:
            if cur:
                blocks.append(cur)
                cur = []
            continue
        cur.append(line)
    if cur:
        blocks.append(cur)
    return blocks


def parse_input_entries(doc_path):
    """
    Parses a transcript docx of the form:
        SPEAKER: text
        HH:MM:SS[:FF]
        <blank>
    A block with no trailing timecode line is treated as a continuation of
    the previous entry's dialogue (this happens when the source transcript
    wraps a long speaker turn across a page/paragraph break) and is merged
    into it rather than becoming its own entry.

    Returns a list of (speaker, text, (h, m, s, f)) tuples, in document order.
    """
    doc = Document(doc_path)
    paragraph_texts = [p.text for p in doc.paragraphs]
    blocks = _split_into_blocks(paragraph_texts)

    entries = []
    for block in blocks:
        tc_match = TIMECODE_RE.match(block[-1]) if block else None

        if tc_match and len(block) >= 2:
            speaker_and_text = " ".join(block[:-1])
            speaker, sep, text = speaker_and_text.partition(":")
            if sep:
                speaker = speaker.strip()
                text = text.strip()
            else:
                speaker, text = None, speaker_and_text.strip()

            h, m, s, f = tc_match.groups()
            tc = (int(h), int(m), int(s), int(f) if f is not None else 0)
            entries.append([speaker, text, tc])
        else:
            # No timecode in this block -> continuation of the previous entry.
            continuation_text = " ".join(block)
            if entries:
                entries[-1][1] = (entries[-1][1] + " " + continuation_text).strip()
            # If there is no previous entry yet, the stray text is dropped;
            # run_qc_checks below can't recover it but the file is malformed
            # in that case regardless.

    return [tuple(e) for e in entries]


# ---------------------------------------------------------------------------
# QC
# ---------------------------------------------------------------------------

def run_qc_checks(entries):
    """
    Validates parsed entries before building output. Returns a list of
    human-readable error strings (empty list means the entries are clean).
    """
    errors = []
    prev_seconds = None

    for idx, (speaker, text, tc) in enumerate(entries, start=1):
        if not speaker:
            errors.append(f"Entry {idx}: missing speaker.")
        if not text:
            errors.append(f"Entry {idx}: missing dialogue text.")

        h, m, s, _f = tc
        total_seconds = h * 3600 + m * 60 + s
        if prev_seconds is not None and total_seconds < prev_seconds:
            errors.append(
                f"Entry {idx}: timecode {h}:{m:02d}:{s:02d} goes backwards "
                f"relative to the previous entry."
            )
        prev_seconds = total_seconds

    return errors


# ---------------------------------------------------------------------------
# Output document construction
# ---------------------------------------------------------------------------

def build_output(entries, media_name, offset, fps):
    """Builds the final Document object from parsed + validated entries."""
    off_h, off_m, off_s, off_f = offset
    offset_frames = timecode_to_frames(off_h, off_m, off_s, off_f, fps)

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
    # Split into two runs (label / value) to match the reference doc's run
    # structure exactly, rather than one merged run.
    hr1 = hp.add_run("Transcription Media # ")
    hr2 = hp.add_run(media_name)
    for hr in (hr1, hr2):
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
            # Space and value as separate runs, matching the reference doc.
            r2 = p.add_run(" ")
            r2.font.name = "Arial"
            r2.font.size = Pt(12)
            r3 = p.add_run(text)
            r3.font.name = "Arial"
            r3.font.size = Pt(12)
        else:
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(12)
        p.paragraph_format.line_spacing = 1.0
        return p

    def add_blank():
        # Blank spacer paragraphs get the same line spacing as every other
        # paragraph in the reference doc -- previously these were left at
        # the document default, which is a formatting mismatch.
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = 1.0
        return p

    add_para(media_name, bold_prefix="MEDIA #:")

    last_index = len(entries) - 1
    for i, (speaker, text, tc) in enumerate(entries):
        h, m, s, f = tc
        raw_frames = timecode_to_frames(h, m, s, f, fps)
        out_tc = frames_to_timecode(raw_frames + offset_frames, fps)

        add_para(f"[{media_name}]")
        add_para(f"[{out_tc}]")

        if speaker and speaker.strip().upper() == "Q":
            add_para(f"[Q]: {text}")
        elif speaker:
            add_para(f"[{speaker}] {text}")
        else:
            add_para(text)

        # Reference doc has a blank spacer between entries, but NOT a
        # trailing blank paragraph after the very last entry.
        if i != last_index:
            add_blank()

    return doc


# ---------------------------------------------------------------------------
# Entry point used by app.py
# ---------------------------------------------------------------------------

def run(input_path, output_path, media_name, offset, fps):
    """
    Parses the uploaded transcript, applies the timecode offset/fps, runs QC
    checks, and writes the formatted output docx.

    Returns {"errors": [...]} on failure, or {"output_path": output_path}
    on success.
    """
    try:
        fps = int(fps)
        if fps <= 0:
            raise ValueError
    except (TypeError, ValueError):
        return {"errors": [f"FPS must be a positive whole number, got: {fps!r}"]}

    try:
        off = parse_offset(offset)
    except ValueError as e:
        return {"errors": [str(e)]}

    entries = parse_input_entries(input_path)
    if not entries:
        return {"errors": ["No valid transcript entries found in the uploaded file."]}

    qc_errors = run_qc_checks(entries)
    if qc_errors:
        return {"errors": qc_errors}

    doc = build_output(entries, media_name, off, fps)
    doc.save(output_path)
    return {"output_path": output_path}


if __name__ == "__main__":
    import sys
    in_path = sys.argv[1]
    out_path = sys.argv[2]
    media = sys.argv[3] if len(sys.argv) > 3 else "MEDIA"
    off = sys.argv[4] if len(sys.argv) > 4 else "00:00:00:00"
    fps_arg = int(sys.argv[5]) if len(sys.argv) > 5 else 25

    result = run(in_path, out_path, media, off, fps_arg)
    if "errors" in result:
        print("Errors:")
        for e in result["errors"]:
            print(" -", e)
        sys.exit(1)
    print(f"Wrote {out_path}")
