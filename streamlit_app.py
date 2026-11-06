"""Chapter Draft Test

Runs Prompt B through the Anthropic API with the same inputs an incognito
claude.ai run receives, so the two drafts can be compared:

  system   = the published claude.ai system prompt for the model
             + your profile preferences
  message  = the attached documents (Prompt B, the Book File)
             + "Write Chapter N."

Nothing else is added to the request.

Deploy on Streamlit Community Cloud with this file and requirements.txt.
Put ANTHROPIC_API_KEY in the app's Secrets. Optionally add
system_prompt.txt and preferences.txt beside this file so they load on start.
"""

import datetime
import io
import json
import re
from pathlib import Path

import streamlit as st
from anthropic import Anthropic
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

HERE = Path(__file__).parent
AUTHOR = "Walter deMilly"
MODELS = ["claude-opus-5-5", "claude-sonnet-5-5", "claude-fable-5-1", "Other"]
SYSTEM_PROMPT_PAGE = (
    "https://platform.claude.com/docs/en/release-notes/system-prompts/overview"
)


# ---------------------------------------------------------------- reading

def docx_to_text(file) -> str:
    """Body text of a .docx in reading order, tables included."""
    doc = Document(file)
    lines = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            lines.append(Paragraph(child, doc).text)
        elif tag == "tbl":
            for row in Table(child, doc).rows:
                lines.append(" | ".join(c.text.strip() for c in row.cells))
            lines.append("")
    text = "\n".join(lines)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def file_to_text(uploaded) -> str:
    if uploaded.name.lower().endswith(".docx"):
        return docx_to_text(uploaded)
    return uploaded.getvalue().decode("utf-8", errors="replace")


def default_text(name: str) -> str:
    path = HERE / name
    return path.read_text(encoding="utf-8") if path.exists() else ""


def fill_date(system_text: str) -> str:
    now = datetime.datetime.now()
    stamp = now.strftime("%A, %B %d, %Y")
    for key in ("{{currentDateTime}}", "{{currentDate}}", "{{current_date}}"):
        system_text = system_text.replace(key, stamp)
    return system_text


def documents_block(docs) -> str:
    """Anthropic's documented multi-document layout."""
    parts = ["<documents>"]
    for i, (name, text) in enumerate(docs, start=1):
        parts.append(
            f'<document index="{i}">\n<source>{name}</source>\n'
            f"<document_content>\n{text}\n</document_content>\n</document>"
        )
    parts.append("</documents>")
    return "\n".join(parts)


# ---------------------------------------------------------------- output

def word_count(text: str) -> int:
    return len(re.findall(r"[A-Za-z0-9’'\-]+", text))


def chapter_part(reply: str) -> str:
    """The longest stretch between --- rules: the chapter itself."""
    chunks = re.split(r"(?m)^\s*(?:---|\*\*\*)\s*$", reply)
    return max(chunks, key=word_count).strip() if chunks else reply


def smart_quotes(s: str) -> str:
    s = re.sub(r'(^|[\s(\[{—–-])"', r"\1“", s).replace('"', "”")
    s = re.sub(r"(^|[\s(\[{—–-])'(?=\w)", r"\1‘", s).replace("'", "’")
    return s


def add_inline(paragraph, text: str):
    for piece in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not piece:
            continue
        if piece.startswith("**") and piece.endswith("**"):
            paragraph.add_run(piece[2:-2]).bold = True
        elif piece.startswith("*") and piece.endswith("*"):
            paragraph.add_run(piece[1:-1]).italic = True
        else:
            paragraph.add_run(piece)


def to_docx(text: str, title: str) -> bytes:
    doc = Document()
    for block in re.split(r"\n\s*\n", text.strip()):
        line = " ".join(block.split("\n")).strip()
        if not line or line in ("---", "***"):
            continue
        if line.startswith("#"):
            doc.add_heading(smart_quotes(line.lstrip("#").strip()), level=2)
        elif re.fullmatch(r"\*\*[^*]+\*\*", line):
            doc.add_heading(smart_quotes(line.strip("*")), level=2)
        else:
            add_inline(doc.add_paragraph(), smart_quotes(line))
    cp = doc.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.title = title
    cp.comments = cp.keywords = cp.subject = cp.category = ""
    cp.created = cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ---------------------------------------------------------------- the call

def run_once(client, model, max_tokens, system_text, docs_xml, message, extra, show):
    kwargs = dict(
        model=model,
        max_tokens=max_tokens,
        system=[{"type": "text", "text": system_text}],
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": docs_xml, "cache_control": {"type": "ephemeral"}},
                {"type": "text", "text": message},
            ],
        }],
    )
    kwargs.update(extra)
    text = ""
    with client.messages.stream(**kwargs) as stream:
        for piece in stream.text_stream:
            text += piece
            if len(text) % 400 < len(piece):
                show.markdown(text[-3000:])
        final = stream.get_final_message()
    show.markdown(text[-3000:])
    return text, final.usage, final.stop_reason


# ---------------------------------------------------------------- the page

def remember_upload(uploaded, state_key):
    """Load an uploaded file into a text box once per new file."""
    if uploaded is None:
        return
    ident = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get(state_key + "_from") != ident:
        st.session_state[state_key] = file_to_text(uploaded)
        st.session_state[state_key + "_from"] = ident


def main():
    st.set_page_config(page_title="Chapter Draft Test", layout="wide")
    st.title("Chapter Draft Test")
    st.caption(
        "Sends the published claude.ai system prompt, your preferences, "
        "Prompt B, the Book File and “Write Chapter N.” Nothing else."
    )

    for key, fname in (("system_prompt", "system_prompt.txt"), ("preferences", "preferences.txt")):
        if key not in st.session_state:
            st.session_state[key] = default_text(fname)
    st.session_state.setdefault("results", [])

    with st.sidebar:
        st.header("Model")
        choice = st.selectbox("Model", MODELS, index=0)
        model = st.text_input("Model string", "") if choice == "Other" else choice
        max_tokens = st.number_input("Max output tokens", 4000, 128000, 32000, 1000)
        runs = st.number_input("Runs of the same chapter", 1, 3, 1)
        with st.expander("Extra request parameters (JSON)"):
            extra_raw = st.text_area(
                "Merged into the request. Leave empty to send the model's defaults.",
                "", height=80,
            )
        try:
            key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:  # noqa: BLE001  (no secrets file)
            key = ""
        if not key:
            key = st.text_input("Anthropic API key", type="password")

    st.subheader("1. System prompt")
    st.markdown(
        f"Copy the claude.ai system prompt for the model from [Anthropic's page]"
        f"({SYSTEM_PROMPT_PAGE}). Paste it here, upload it, or put it beside the "
        "app as system_prompt.txt."
    )
    remember_upload(st.file_uploader("Upload system prompt", type=["txt", "md", "docx"], key="sp_up"), "system_prompt")
    st.text_area("System prompt", key="system_prompt", height=160)

    st.subheader("2. Your profile preferences")
    st.markdown("Copy them from Settings › Profile. They reach incognito chats, so they belong in the test.")
    remember_upload(st.file_uploader("Upload preferences", type=["txt", "md", "docx"], key="pr_up"), "preferences")
    st.text_area("Preferences", key="preferences", height=120)

    st.subheader("3. The two documents")
    c1, c2 = st.columns(2)
    prompt_file = c1.file_uploader("Prompt B (.docx)", type=["docx"], key="pb")
    book_file = c2.file_uploader("Book File (.docx)", type=["docx"], key="bf")

    st.subheader("4. The message")
    chapter = st.text_input("Chapter number", "")
    message = f"Write Chapter {chapter.strip()}." if chapter.strip().isdigit() else ""
    if message:
        st.code(message)

    ready = all([model, key, st.session_state.system_prompt.strip(), prompt_file, book_file, message])
    if not ready:
        st.info("Fill in the system prompt, attach both documents and give the chapter number.")
    if st.button("Write the chapter", type="primary", disabled=not ready):
        try:
            extra = json.loads(extra_raw) if extra_raw.strip() else {}
        except json.JSONDecodeError as e:
            st.error(f"Extra parameters aren't valid JSON: {e}")
            return

        system_text = fill_date(st.session_state.system_prompt.strip())
        prefs = st.session_state.preferences.strip()
        if prefs:
            system_text += f"\n\n<userPreferences>\n{prefs}\n</userPreferences>"
        docs = [(prompt_file.name, docx_to_text(prompt_file)), (book_file.name, docx_to_text(book_file))]
        docs_xml = documents_block(docs)

        sent = {
            "model": model,
            "max_tokens": max_tokens,
            "system prompt (characters)": len(system_text),
            "preferences included": bool(prefs),
            "documents": [f"{n} ({word_count(t):,} words)" for n, t in docs],
            "message": message,
            "extra parameters": extra,
        }
        client = Anthropic(api_key=key, timeout=1800, max_retries=2)
        for k in range(1, int(runs) + 1):
            st.markdown(f"**Run {k} of {runs}**")
            show = st.empty()
            try:
                reply, usage, stop = run_once(client, model, int(max_tokens), system_text, docs_xml, message, extra, show)
            except Exception as e:  # noqa: BLE001
                st.error(f"The request failed: {e}")
                break
            st.session_state.results.append({
                "chapter": int(chapter), "run": k, "model": model, "reply": reply,
                "stop": stop, "sent": sent,
                "usage": {
                    "input": usage.input_tokens, "output": usage.output_tokens,
                    "cache read": getattr(usage, "cache_read_input_tokens", 0) or 0,
                    "cache write": getattr(usage, "cache_creation_input_tokens", 0) or 0,
                },
                "time": datetime.datetime.now().strftime("%H:%M"),
            })
            show.empty()

    if st.session_state.results:
        st.subheader("Results")
        if st.button("Clear results"):
            st.session_state.results = []
            st.rerun()
    for i, r in enumerate(reversed(st.session_state.results)):
        chap = chapter_part(r["reply"])
        name = f"Chapter_{r['chapter']:02d}_API_{r['model']}_run{r['run']}"
        with st.expander(
            f"Chapter {r['chapter']}, run {r['run']}, {r['model']}, {r['time']}: "
            f"{word_count(chap):,} words", expanded=(i == 0)
        ):
            st.write(
                f"Stop reason: {r['stop']}.  Tokens: {r['usage']['input']:,} in, "
                f"{r['usage']['output']:,} out, {r['usage']['cache read']:,} read from cache."
            )
            if r["stop"] == "max_tokens":
                st.warning("The reply hit the token ceiling. Raise Max output tokens and run again.")
            d1, d2 = st.columns(2)
            d1.download_button("Chapter as Word file", to_docx(chap, name.replace("_", " ")),
                               file_name=name + ".docx", key=f"dx{i}")
            d2.download_button("Whole reply as text", r["reply"].encode("utf-8"),
                               file_name=name + "_full_reply.txt", key=f"tx{i}")
            with st.popover("What was sent"):
                st.json(r["sent"])
            st.markdown(r["reply"])


if __name__ == "__main__":
    main()
