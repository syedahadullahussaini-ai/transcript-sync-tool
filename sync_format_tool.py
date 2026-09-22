import re
from docx import Document
from docx.shared import Pt, RGBColor

BLOCK_MEDIA_RE = re.compile(r"^\[(.+)\]$")
BLOCK_TC_RE = re.compile(r"^\[(\d{1,2}:\d{2}:\d{2}(?::\d{1,2})?)\]$")


def parse_reference_blocks(lines):
    """
    Parses text already in the target [MEDIA] / [TIMECODE] / [SPEAKER] text
    layout (optionally preceded by a 'MEDIA #: ...' header line) into a
    list of (media_name, timecode, speaker_line_text) tuples. Blank lines
    (or lines that are just whitespace) separate blocks.
    """
    body_label_line = None
    blocks = []
    cur = []

    def flush():
        if not cur:
            return
        if len(cur) >= 3:
            media = cur[0]
            tc = cur[1]
            speaker_line = " ".join(cur[2:])  # in case dialogue wrapped onto >1 line
            blocks.append((media, tc, speaker_line))

    for raw in lines:
        line = raw.strip()
        if line.upper().startswith("MEDIA #:"):
            body_label_line = line
            continue
        if not line:
            flush()
            cur = []
            continue
        cur.append(line)
    flush()

    return body_label_line, blocks


def build_doc_from_reference(lines, media_name=None,
                              running_header_label="Transcription Media #",
                              body_label="MEDIA #:"):
    body_label_line, blocks = parse_reference_blocks(lines)
    if not blocks:
        raise ValueError("No [MEDIA]/[TIMECODE]/[SPEAKER] blocks found in the input")

    # media name: prefer explicit arg, else pull from the first bracketed line
    if media_name is None:
        m = BLOCK_MEDIA_RE.match(blocks[0][0])
        media_name = m.group(1) if m else blocks[0][0]

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

    for media_line, tc_line, speaker_line in blocks:
        add_para(media_line)
        add_para(tc_line)
        add_para(speaker_line)
        doc.add_paragraph("")

    return doc


if __name__ == "__main__":
    import sys
    in_path = sys.argv[1]
    out_path = sys.argv[2]
    with open(in_path, encoding="utf-8") as f:
        lines = f.readlines()
    doc = build_doc_from_reference(lines)
    doc.save(out_path)
    print(f"Wrote {out_path}")
