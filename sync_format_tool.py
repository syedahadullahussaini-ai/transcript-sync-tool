#!/usr/bin/env python3
"""
sync_format_tool.py
--------------------
Converts a raw production transcript (.docx) into the standard synced
transcript format used for finished media files:

    Transcription Media # <MEDIA_NAME>

    [<MEDIA_NAME>]
    [HH:MM:SS:FF]
    [SPEAKER] Dialogue text.

    [<MEDIA_NAME>]
    [HH:MM:SS:FF]
    [SPEAKER] Next line of dialogue.
    ...

It expects the raw doc to look like a Trance-style dump: dialogue lines
(optionally prefixed with "Speaker Name: "), interspersed with bare
relative timecode lines in HH:MM:SS form (no frames), e.g.:

    Crew: Take one. Mark. Roll camera.
    00:00:00

    Yasaswi ard: All right, Beautiful guy. Dynamite.
    00:00:10

Every raw timecode is treated as the point in the *relative* media
timeline where the text immediately above it (back to the previous
timecode) occurs. That relative time is added — at true frame
precision, using the file's frame rate — to the burnt-in start offset
(HH:MM:SS:FF) supplied on the command line, to produce the final
burnt-in timecode. Frames roll over into seconds (and seconds into
minutes/hours) correctly once the fps is known, instead of just being
carried through unchanged. Raw timecodes may be given as HH:MM:SS
(frames assumed 0) or HH:MM:SS:FF if the raw doc already has them.

USAGE
-----
    python3 sync_format_tool.py \
        --input "Test_-_Script_Sync.docx" \
        --output "Conversation_proxy_SYNCED.docx" \
        --media-name "Conversation_proxy" \
        --offset "11:28:14:15" \
        --fps 25

Optional:
    --speakers "Crew,Q,Yasaswi ard,Michael ward"
        Comma-separated whitelist of EXACT speaker labels to recognize.
        If omitted, the tool auto-detects any "Capitalized Word(s):"
        prefix at the start of a paragraph as a new speaker.

    --fps 25
        Frame rate of the source media (whole frames per second — 24,
        25, 30, 50, 60, etc.). Defaults to 30 if omitted. Needed to
        correctly roll frames over into seconds when the offset frame
        plus any raw frame value reaches the frame rate. Drop-frame
        rates (29.97/59.94 NDF/DF) are treated as their rounded whole
        value (30/60); true drop-frame timecode math isn't applied.

    --running-header-label "Transcription Media #"
    --body-label "MEDIA #:"
        Text of the two bold labels (defaults shown above).
"""

import argparse
import re
def validate_offset(offset_str):
    pattern = r"^\d{2}:\d{2}:\d{2}:\d{2}$"
    return re.match(pattern, offset_str)

from docx import Document
from docx.shared import Pt, RGBColor

TIMECODE_RE = re.compile(r'^(\d{1,2}):(\d{2}):(\d{2})(?::(\d{1,2}))?$')
SPEAKER_RE = re.compile(r"^([A-Z][A-Za-z0-9&'.]*(?:\s[A-Za-z0-9&'.]+){0,3}):\s+(.*)$")


def parse_offset(offset_str):
    parts = offset_str.strip().split(':')
    if len(parts) != 4:
        raise ValueError(f"Offset must be HH:MM:SS:FF, got '{offset_str}'")
    h, m, s, f = (int(p) for p in parts)
    return h, m, s, f


def to_frame_count(h, m, s, f, fps):
    return (h * 3600 + m * 60 + s) * fps + f


def frames_to_timecode(total_frames, fps):
    total_seconds, frames = divmod(int(total_frames), fps)
    h = (total_seconds // 3600) % 24
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02d}:{s:02d}:{int(frames):02d}"


def offset_timecode(raw_tc, raw_frames, offset, fps):
    """raw_tc = (h, m, s); raw_frames = frame value from the raw doc (0 if
    the raw doc has no frame column); offset = (h, m, s, f); fps = int."""
    offset_h, offset_m, offset_s, offset_f = offset
    rh, rm, rs = raw_tc
    raw_total = to_frame_count(rh, rm, rs, raw_frames, fps)
    offset_total = to_frame_count(offset_h, offset_m, offset_s, offset_f, fps)
    return frames_to_timecode(raw_total + offset_total, fps)


def extract_segments(paragraphs, known_speakers=None, merge_bundled=True):
    """
    Walk paragraphs, grouping dialogue text between bare timecode markers.

    Multiple dialogue paragraphs can pile up before a single timecode marker
    is hit (e.g. an untimed filler line with no timecode of its own, sitting
    right before a line that does have one). When that happens and
    merge_bundled is True, everything bundled into that one flush is joined
    into a SINGLE entry — because the source never gave them separate
    timestamps in the first place.

    This is deliberately NOT the same as two lines that each already have
    their own timecode paragraph immediately after them, even if those two
    timecodes happen to carry the identical value (common with 1-second
    source resolution during fast back-and-forth dialogue) — those are left
    as separate entries, since each one really was individually timed in
    the source; collapsing them would erase real distinct lines.

    Returns (entries, leftover):
      entries  -> list of ((h, m, s), frames, speaker, text), in order
      leftover -> list of (speaker, text) segments at the end of the doc
                  that were never followed by a timecode (untimed, skipped)
    """
    buffered = []      # (speaker, text) not yet timed, for the CURRENT flush
    entries = []
    current_speaker = None

    for para in paragraphs:
        text = para.strip()
        if not text:
            continue

        tc_match = TIMECODE_RE.match(text)
        if tc_match:
            h, m, s, f = tc_match.groups()
            raw_tc = (int(h), int(m), int(s))
            raw_frames = int(f) if f is not None else 0
            if buffered:
                if merge_bundled and len(buffered) > 1:
                    speakers = []
                    for spk, _ in buffered:
                        if spk not in speakers:
                            speakers.append(spk)
                    combined_text = " ".join(txt for _, txt in buffered)
                    entries.append((raw_tc, raw_frames, " / ".join(speakers), combined_text))
                else:
                    for spk, txt in buffered:
                        entries.append((raw_tc, raw_frames, spk, txt))
            buffered = []
            continue

        spk_match = SPEAKER_RE.match(text)
        if spk_match and (known_speakers is None or spk_match.group(1) in known_speakers):
            current_speaker = spk_match.group(1)
            dialogue = spk_match.group(2).strip()
        else:
            dialogue = text

        if current_speaker is None:
            current_speaker = "UNKNOWN"

        buffered.append((current_speaker, dialogue))

    return entries, buffered


def speaker_label(speaker, text):
    """Return the bracketed line for a given speaker + text, matching
    house style: named individuals -> [NAME] text ; Crew -> [CREW] text ;
    interviewer/question -> [Q]: text. When speakers were merged together
    (e.g. "Crew / Yasaswi ard"), each part is upper-cased and rejoined."""
    s = speaker.strip()
    if "/" in s:
        parts = [speaker_label(p.strip(), "").split("]")[0].lstrip("[") for p in s.split("/")]
        return f"[{' / '.join(parts)}] {text}"
    if s.lower() == "q":
        return f"[Q]: {text}"
    if s.lower() == "crew":
        return f"[CREW] {text}"
    return f"[{s.upper()}] {text}"


def build_output_doc(entries, media_name, offset, fps,
                      running_header_label="Transcription Media #",
                      body_label="MEDIA #:", show_frames=True):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Pt(612)    # 8.5in
    section.page_height = Pt(792)   # 11in
    section.left_margin = Pt(144)   # 2in
    section.right_margin = Pt(144)  # 2in
    section.top_margin = Pt(72)     # 1in
    section.bottom_margin = Pt(72)  # 1in

    # Running page header: "Transcription Media # <name>" — bold, gray, Arial Bold 12pt
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

    for raw_tc, raw_frames, speaker, text in entries:
        new_tc = offset_timecode(raw_tc, raw_frames, offset, fps)
        if not show_frames:
            new_tc = new_tc.rsplit(":", 1)[0]  # drop the :FF, keep HH:MM:SS
        add_para(f"[{media_name}]")
        add_para(f"[{new_tc}]")
        add_para(speaker_label(speaker, text))
        doc.add_paragraph("")  # spacer between blocks

    return doc
def run_qc_checks(entries):
    errors = []
    prev_time = None

    for i, (raw_tc, raw_frames, speaker, text) in enumerate(entries):
        h, m, s = raw_tc

        if raw_tc is None:
            errors.append(f"Line {i+1}: Missing timestamp")

        if not speaker or speaker.strip() == "":
            errors.append(f"Line {i+1}: Empty speaker")

        if h > 23 or m > 59 or s > 59:
            errors.append(f"Line {i+1}: Invalid SMPTE time {h}:{m}:{s}")

        current_time = h*3600 + m*60 + s
        if prev_time is not None and current_time < prev_time:
            errors.append(f"Line {i+1}: Time overlap detected")

        prev_time = current_time

    return errors
def validate_offset(offset_str):
    import re
    pattern = r"^\d{2}:\d{2}:\d{2}:\d{2}$"
    return re.match(pattern, offset_str)

def run(input_path, output_path, media_name, offset_str, fps=30, speakers=None,
        running_header_label="Transcription Media #", body_label="MEDIA #:",
        merge_bundled=True, show_frames=True):
    fps = int(round(float(fps)))
    if fps <= 0:
        raise ValueError(f"fps must be a positive number, got '{fps}'")

    known = set(s.strip() for s in speakers.split(",")) if speakers else None
    src = Document(input_path)
    paragraphs = [p.text for p in src.paragraphs]

    entries, leftover = extract_segments(paragraphs, known_speakers=known,
                                          merge_bundled=merge_bundled)
            qc_errors = run_qc_checks(entries)

if qc_errors:
    return {"errors": qc_errors}


    if leftover:
        print(f"WARNING: {len(leftover)} trailing dialogue segment(s) had no "
              f"following timecode and were skipped:")
        for spk, txt in leftover:
            print(f"  [{spk}] {txt[:60]}")
if not validate_offset(offset_str):
    return {"errors": ["Invalid offset format. Use HH:MM:SS:FF"]}

offset = parse_offset(offset_str)
    offset = parse_offset(offset_str)
    if offset[3] >= fps:
        print(f"WARNING: offset frame value {offset[3]} is >= fps ({fps}); "
              f"double-check the fps you passed in.")

    out_doc = build_output_doc(entries, media_name, offset, fps,
                                running_header_label=running_header_label,
                                body_label=body_label, show_frames=show_frames)
    out_doc.save(output_path)
    print(f"Wrote {len(entries)} synced entries to {output_path} at {fps}fps")
    return len(entries)


def main():
    ap = argparse.ArgumentParser(
        description="Format a raw transcript into synced [MEDIA]/[TIMECODE]/[SPEAKER] layout.")
    ap.add_argument("--input", required=True, help="Path to raw .docx transcript")
    ap.add_argument("--output", required=True, help="Path to write formatted .docx")
    ap.add_argument("--media-name", required=True, help="Media/file name, e.g. Conversation_proxy")
    ap.add_argument("--offset", required=True, help="Burnt-in start timecode, HH:MM:SS:FF")
    ap.add_argument("--fps", default=30,
                     help="Frame rate of the source media (e.g. 24, 25, 30, 50, 60). Default 30.")
    ap.add_argument("--speakers", default=None,
                     help="Comma-separated whitelist of exact speaker labels (optional)")
    ap.add_argument("--running-header-label", default="Transcription Media #",
                     help="Bold gray running page-header prefix text")
    ap.add_argument("--body-label", default="MEDIA #:",
                     help="Bold label on the first body line")
    ap.add_argument("--no-merge-bundled", action="store_true",
                     help="Disable merging of fragments that land on the same timecode "
                          "(default is to merge them into one line)")
    ap.add_argument("--hide-frames", action="store_true",
                     help="Drop the :FF frame component from displayed timecodes "
                          "(default is to show HH:MM:SS:FF with an unpadded hour)")
    args = ap.parse_args()
    run(args.input, args.output, args.media_name, args.offset, fps=args.fps,
        speakers=args.speakers, running_header_label=args.running_header_label,
        body_label=args.body_label, merge_bundled=not args.no_merge_bundled,
        show_frames=not args.hide_frames)


if __name__ == "__main__":
    main()
