import re
from io import BytesIO

from docx import Document

HEADER_RE = re.compile(r'^\s*(.+?)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*$')

def normalize_time(t: str) -> str:
    """
    Accepts 'm:ss', 'mm:ss', or 'h:mm:ss' (or 'hh:mm:ss').
    Returns '[MM:SS]' or '[HH:MM:SS]' (zero-padded).
    """
    parts = t.strip().split(':')
    parts = [p.zfill(2) for p in parts]
    if len(parts) == 2:
        mm, ss = parts
        return f'[{mm}:{ss}]'
    elif len(parts) == 3:
        hh, mm, ss = parts
        return f'[{hh}:{mm}:{ss}]'
    else:
        return f'[{t}]'

def squash_text(paragraphs):
    """
    Clean and join multiple paragraphs of one block into a single line.
    """
    pieces = []
    for p in paragraphs:
        txt = (p or "").strip()
        if not txt:
            continue
        txt = re.sub(r'\s+', ' ', txt)
        pieces.append(txt)
    return ' '.join(pieces).strip()

def parse_blocks(doc: Document):
    """
    Yields (speaker, time_str, content) for each block.
    """
    blocks = []
    cur_speaker = None
    cur_time = None
    cur_text = []

    def flush():
        nonlocal blocks, cur_speaker, cur_time, cur_text
        if cur_speaker and cur_time:
            content = squash_text(cur_text)
            blocks.append((cur_speaker, cur_time, content))
        cur_speaker, cur_time, cur_text = None, None, []

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        m = HEADER_RE.match(text)
        if m:
            flush()
            cur_speaker = m.group(1).strip()
            cur_time = normalize_time(m.group(2))
            cur_text = []
        else:
            if cur_speaker and cur_time:
                cur_text.append(text)

    flush()
    return blocks


# ------------------ THIS PART IS THE ONLY CHANGE ------------------ #
def build_output_document(blocks):
    doc = Document()

    p = doc.add_paragraph("TRANSCRIPT:")
    p.runs[0].bold = True

    for speaker, time_str, content in blocks:
        label = speaker.upper()

        p = doc.add_paragraph()
        bold_run = p.add_run(f"{time_str} {label}:")
        bold_run.bold = True

        if content:
            p.add_run(f" {content}")

    return doc
# ------------------------------------------------------------------ #


def write_output(blocks, out_path: str):
    doc = build_output_document(blocks)
    doc.save(out_path)

def convert_docx(input_path: str, output_path: str):
    src = Document(input_path)
    blocks = parse_blocks(src)
    write_output(blocks, output_path)

def convert_docx_bytes(data: bytes) -> bytes:
    src = Document(BytesIO(data))
    blocks = parse_blocks(src)
    doc = build_output_document(blocks)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

def run_streamlit_app():
    import streamlit as st

    st.set_page_config(page_title="Otter Transcript Formatter", page_icon="DOC")
    st.title("Otter Transcript Formatter")
    st.write(
        "Upload an Otter-style `.docx` transcript formatted like the EX_INPUT sample. "
        "We'll convert each block into the single-line style used by EX_OUTPUT."
    )

    uploaded = st.file_uploader("Choose a .docx file", type=["docx"])
    if uploaded is not None:
        try:
            result = convert_docx_bytes(uploaded.getvalue())
        except Exception as exc:
            st.error(f"Conversion failed: {exc}")
            return

        output_name = uploaded.name.rsplit(".", 1)[0] + "_formatted.docx"
        st.success("Conversion complete. Download your formatted transcript below.")
        st.download_button(
            label="Download formatted .docx",
            data=result,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Convert Otter-style transcript .docx files to single-line transcript style."
    )
    parser.add_argument("input", nargs="?", help="Path to EX_INPUT-style .docx")
    parser.add_argument("output", nargs="?", help="Path to write the reformatted .docx")
    parser.add_argument(
        "--streamlit",
        action="store_true",
        help="Launch the Streamlit interface instead of running the CLI converter.",
    )
    args, _ = parser.parse_known_args()

    if args.streamlit or (args.input is None and args.output is None):
        run_streamlit_app()
    elif args.input and args.output:
        convert_docx(args.input, args.output)
    else:
        parser.error("Specify both input and output paths, or use --streamlit to launch the UI.")

if __name__ == "__main__":
    main()
