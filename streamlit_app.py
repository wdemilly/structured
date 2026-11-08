"""Chapter Draft Test

Runs Prompt B through the Anthropic API so the draft can be read beside the
one an incognito claude.ai window writes, and then reviews the draft.

  draft    system   = the published claude.ai system prompt for the model
                      + the API's own fiction instructions
           message  = the attached documents (Prompt B, the Book File)
                      + "Write Chapter N."

  review   the same system prompt and documents, then the finished chapter,
           a list of places the app found to look at, the review
           instructions and "Review Chapter N."  The reviewer answers with
           fixes and questions.  The app applies each fix by finding its
           passage word for word, and returns the chapter as a Word file
           with every fix as a tracked change, a clean copy, and notes.

A chapter written elsewhere (an incognito window, say) can be reviewed on its
own in section 6.

Deploy on Streamlit Community Cloud with this file and requirements.txt.
Put ANTHROPIC_API_KEY in the app's Secrets. The app fetches the newest
claude.ai system prompt for the chosen model from Anthropic's published page;
a system_prompt.txt beside this file takes its place. The fiction
instructions and the review instructions are built in; api_instructions.txt
or review_instructions.txt beside the file replaces them.

Section 4 takes one chapter or a run of chapters, e.g. "2 to 4". In a run, each
chapter is drafted once and reviewed, and the reviewed chapter goes under
CHAPTERS SO FAR before the next is written. The page then offers the Book File
with the new chapters in it, with and without the catalogue.
"""

import datetime
import io
import json
import re
import threading
import time
import urllib.request
import zipfile
from pathlib import Path

import streamlit as st
from anthropic import Anthropic
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

HERE = Path(__file__).parent
AUTHOR = "Walter deMilly"
MODELS = ["claude-opus-5-5", "claude-sonnet-5-5", "claude-fable-5-1", "Other"]
SYSTEM_PROMPT_PAGE = "https://platform.claude.com/docs/en/release-notes/system-prompts/{model}"

# The text of API_Fiction_Instructions_v2.docx in WritingApp.
FICTION_INSTRUCTIONS = (
    "Fiction instructions\n\n"
    "These instructions apply when you write fiction from an attached prompt. "
    "Follow that prompt for the prose. In the narration, keep three habits as well. "
    "Dialogue belongs to the people speaking.\n\n"
    "1. Set off an aside with commas or parentheses, or give it a sentence of its own. "
    "Use no em dash.\n\n"
    "2. Give a deliberate act to someone who can act, never to an abstraction such as "
    "an argument, a chapter, a mechanism or a thesis.\n\n"
    "3. Where two clauses could each stand as a sentence, end the first and begin the "
    "second. Join them with a comma and “and” only in a run you mean as a pile-up."
)

# Review instructions, v1. They name no book, so they serve any novel the pipeline writes.
REVIEW_INSTRUCTIONS = """Review instructions

The chapter above, in the chapter tags, was drafted by another writer from the attached prompt and Book File. Read the prompt, the Book File and the chapter as a reader of the book. Then check the chapter against them. Mend what breaks a written rule. Bring the author the questions that are his to decide.

Below the chapter, in the places tags, is a list the app made of places to look at: sentences of narration with a comma and “and,” sentences of narration with a dash or a figure, and stretches of wording that appear in both the chapter and the Book File. The list shows where to look. You decide what each place needs, if anything.

What to check

1. The chapter entry. Every event in the entry happens. The chapter stops where the entry says. Where the entry has a person named or a thing said at a given moment, the chapter has it there.

2. Later chapters. The chapter shows nothing that the chapter list puts in a later chapter.

3. The physical world. Every action can happen where it happens, under the Book File’s laws: its place, its air or vacuum, its weight or lack of it, its light and its tide.

4. Facts. The chapter may state as fact something the Book File doesn’t give: a law, the name of a law or an office, an event, a date, a relationship, a piece of history. Each of these is a question for the author. Leave it in the chapter as written.

5. Borrowed wording. The narrator tells the Book File’s facts in his own words, from where the person in the scene is standing. He repeats no sentence of an earlier chapter. Where a stretch of the chapter repeats a sentence of the Book File, put the fact into the narrator’s own words. Names and terms of art stay as they are.

6. The sentence. The prompt’s rules for sentences, its list of things to leave out, and the habits in your instructions.

How to mend a join

A comma and “and” before a second verb with the same subject joins no clauses and stays.

Where a sentence of narration joins two clauses with a comma and “and,” mend it in one of the ways the prompt gives. Try them in this order. Give both actions to one subject. Put one clause inside the other with when, as, before, after, until, while or because. Carry the lesser fact in a phrase attached to the other: a participle, an appositive, or a phrase with “with.” End the first clause and begin a new sentence only where the second thing deserves a sentence of its own.

Before you settle a mend, read it in place with the sentences on either side. The paragraph should still rise and fall, with long sentences among the short ones. Where a mend would leave a run of short plain sentences, each opening on its subject and verb, choose another form. You can also take the neighbouring sentence into the passage and recast the two together. Mend the joins in one paragraph together, so that the paragraph varies. Where every form reads worse than the join, leave the join as it is. The prompt allows it for a pile-up the narrator means.

How to write a fix

Change as little as you can. Keep the narrator’s voice and the chapter’s own words wherever they serve. Bring in no new fact, image, comparison or event. Keep the passage about as long as it was.

The passage is the chapter’s exact text, copied character for character, from one point in a paragraph to another point in the same paragraph. It can run to a few sentences when a mend needs the sentences around it. Give each passage once. Let no two passages overlap. To cut a sentence, give it as the passage and leave the replacement empty.

With each fix, quote the line of the prompt, your instructions or the Book File that it rests on.

What to leave alone

Dialogue belongs to the people speaking. Change a line of speech only where it states a fact the Book File contradicts, or where the chapter entry calls for something said and the chapter leaves it out.

Whatever breaks no written rule stays as the writer wrote it, however you might have written it yourself. Where nothing breaks a rule, give no fixes.

How to answer

Answer with the tagged blocks below and nothing else, the fixes first and then the questions.

<fix>
<passage>the chapter’s exact text</passage>
<replacement>the new text</replacement>
<rule>the line it rests on, quoted</rule>
</fix>

<question>
<passage>the chapter’s exact text the question is about</passage>
<ask>the question for the author, in a sentence or two</ask>
</question>"""


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


def chapter_from_upload(uploaded) -> str:
    """A chapter file as text, one blank line between paragraphs."""
    if uploaded.name.lower().endswith(".docx"):
        doc = Document(uploaded)
        return "\n\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    text = uploaded.getvalue().decode("utf-8", errors="replace").replace("\r\n", "\n")
    if not re.search(r"\n\s*\n", text):
        text = "\n\n".join(line for line in text.split("\n") if line.strip())
    return text


def default_text(name: str, fallback: str = "") -> str:
    path = HERE / name
    return path.read_text(encoding="utf-8-sig") if path.exists() else fallback


def fill_date(system_text: str) -> str:
    now = datetime.datetime.now()
    stamp = now.strftime("%A, %B %d, %Y")
    for key in ("{{currentDateTime}}", "{{currentDate}}", "{{current_date}}"):
        system_text = system_text.replace(key, stamp)
    return system_text


def newest_prompt(page: str):
    """The newest dated version on Anthropic's page: every fenced block under the first dated heading."""
    for section in re.split(r"(?m)^## ", page)[1:]:
        heading, _, body = section.partition("\n")
        blocks = re.findall(r"(?ms)^```[^\n]*\n(.*?)^```[ \t]*$", body)
        if blocks:
            return heading.strip(), "".join(blocks).strip()
    return "", ""


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_system_prompt(model: str):
    """The newest claude.ai system prompt Anthropic has published for the model."""
    url = SYSTEM_PROMPT_PAGE.format(model=model) + ".md"
    req = urllib.request.Request(url, headers={"User-Agent": "Chapter-Draft-Test"})
    with urllib.request.urlopen(req, timeout=30) as r:
        page = r.read().decode("utf-8")
    date, text = newest_prompt(page)
    if len(text) < 5000 or "Claude" not in text:
        raise ValueError("the page didn't hold a system prompt in the expected form")
    return date, text


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


def split_reply(reply: str):
    """The chapter (the longest stretch between --- rules) and whatever follows it."""
    chunks = [c.strip() for c in re.split(r"(?m)^\s*(?:---|\*\*\*)\s*$", reply)]
    if not chunks:
        return reply.strip(), ""
    k = max(range(len(chunks)), key=lambda i: word_count(chunks[i]))
    notes = "\n\n".join(c for c in chunks[k + 1:] if c)
    return chunks[k], notes


def smart_quotes(s: str) -> str:
    s = re.sub(r'(^|[\s(\[{—–-])"', r"\1“", s).replace('"', "”")
    s = re.sub(r"(^|[\s(\[{—–“-])'(?=\w)", r"\1‘", s).replace("'", "’")
    s = re.sub(r"“’(?=\w)", "“‘", s)
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


def set_properties(doc, title: str):
    cp = doc.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.title = title
    cp.comments = cp.keywords = cp.subject = cp.category = ""
    cp.created = cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)


def save_bytes(doc) -> bytes:
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def is_heading(line: str) -> bool:
    return line.startswith("#") or bool(re.fullmatch(r"\*\*[^*]+\*\*", line))


def add_block(doc, line: str):
    if line.startswith("#"):
        doc.add_heading(line.lstrip("#").strip(), level=2)
    elif re.fullmatch(r"\*\*[^*]+\*\*", line):
        doc.add_heading(line.strip("*"), level=2)
    else:
        add_inline(doc.add_paragraph(), line)


def to_docx(text: str, title: str) -> bytes:
    doc = Document()
    for block in re.split(r"\n\s*\n", text.strip()):
        line = " ".join(block.split("\n")).strip()
        if not line or line in ("---", "***"):
            continue
        add_block(doc, smart_quotes(line))
    set_properties(doc, title)
    return save_bytes(doc)


# ---------------------------------------------------------------- the review: places to look

def paragraphs_of(text: str):
    """The chapter as a list of paragraphs, each on one line."""
    out = []
    for block in re.split(r"\n\s*\n", text.strip()):
        line = " ".join(block.split())
        if line and line not in ("---", "***"):
            out.append(line)
    return out


SPEECH = re.compile(r"“[^”]*”")
SENTENCE_BREAK = re.compile(r"(?<=[.!?…])\s+(?=[A-Z“‘(])")


def narration_sentences(paras):
    """Each sentence of narration, with speech taken out, as it stands in the chapter."""
    for p in paras:
        if is_heading(p):
            continue
        pieces, pos = [], 0
        for m in SPEECH.finditer(p):
            pieces.append(p[pos:m.start()])
            pos = m.end()
        tail = p[pos:]
        if "“" in tail:  # speech that runs on past the paragraph
            tail = tail[:tail.index("“")]
        pieces.append(tail)
        for piece in pieces:
            for s in SENTENCE_BREAK.split(piece):
                s = s.strip(" ,")
                if len(s.split()) >= 4:
                    yield s


WORD = re.compile(r"[a-z0-9']+")


def _tokens(text: str):
    flat = text.lower().replace("’", "'").replace("‘", "'")
    return [(m.group(), m.start(), m.end()) for m in WORD.finditer(flat)]


def borrowed_wording(chapter: str, book_text: str, n: int = 5):
    """Stretches of the chapter whose wording also appears in the Book File."""
    lines = [l.strip() for l in book_text.split("\n") if l.strip()]
    grams = {}
    for li, line in enumerate(lines):
        words = [w for w, _, _ in _tokens(line)]
        for i in range(len(words) - n + 1):
            grams.setdefault(tuple(words[i:i + n]), li)
    toks = _tokens(chapter)
    words = [w for w, _, _ in toks]
    hit = [False] * len(toks)
    where = {}
    for i in range(len(words) - n + 1):
        g = tuple(words[i:i + n])
        if g in grams:
            for k in range(i, i + n):
                hit[k] = True
            where[i] = grams[g]
    spans, i = [], 0
    while i < len(toks):
        if not hit[i]:
            i += 1
            continue
        j = i
        while j + 1 < len(toks) and hit[j + 1]:
            j += 1
        li = next((where[k] for k in range(i, j + 1) if k in where), None)
        text = " ".join(chapter[toks[i][1]:toks[j][2]].split())
        spans.append((text, lines[li] if li is not None else ""))
        i = j + 1
    share = sum(hit) / max(1, len(toks))
    return spans, share


def places_to_look(paras, book_text: str):
    joins, dashes, figures = [], [], []
    for s in narration_sentences(paras):
        if re.search(r",\s+and\s", s):
            joins.append(s)
        if re.search(r"[—–]|\s-\s", s):
            dashes.append(s)
        if re.search(r"\d", s):
            figures.append(s)
    spans, share = borrowed_wording("\n\n".join(paras), book_text)
    note = ""
    if share > 0.5:
        note = ("Most of this chapter's wording also appears in the Book File, so the chapter is "
                "probably in it already. The app has left out the list of shared wording.")
        spans = []

    def section(title, items, fmt=lambda x: f"- {x}"):
        body = "\n".join(fmt(x) for x in items) if items else "None."
        return f"{title}\n{body}"

    text = "\n\n".join([
        section("Sentences of narration with a comma and “and”:", joins),
        section("Sentences of narration with a dash:", dashes),
        section("Sentences of narration with a figure:", figures),
        section("Wording that appears in both the chapter and the Book File, with the Book File line:",
                spans, lambda x: f"- “{x[0]}”\n  Book File line: “{x[1][:240]}”"),
    ] + ([note] if note else []))
    found = dict(joins=joins, dashes=dashes, figures=figures, borrowed=spans, note=note)
    return text, found


# ---------------------------------------------------------------- the review: fixes

def _tag(block: str, name: str) -> str:
    m = re.search(rf"<{name}>(.*?)</{name}>", block, re.S)
    return m.group(1).strip() if m else ""


def parse_review(reply: str):
    fixes = [dict(passage=_tag(b, "passage"), replacement=_tag(b, "replacement"), rule=_tag(b, "rule"))
             for b in re.findall(r"<fix>(.*?)</fix>", reply, re.S)]
    questions = [dict(passage=_tag(b, "passage"), ask=_tag(b, "ask"))
                 for b in re.findall(r"<question>(.*?)</question>", reply, re.S)]
    return fixes, questions


def _plain_quotes(s: str) -> str:
    return s.translate(str.maketrans("‘’“”", "''\"\""))


def locate(chapter: str, passage: str):
    """Where the passage stands in the chapter, word for word, once."""
    words = _plain_quotes(passage).split()
    if not words:
        return None, "the passage was empty"
    pattern = r"\s+".join(re.escape(w) for w in words)
    found = list(re.finditer(pattern, _plain_quotes(chapter)))
    if not found:
        return None, "the passage isn't in the chapter word for word"
    if len(found) > 1:
        return None, "the passage appears in the chapter more than once"
    a, b = found[0].span()
    if "\n" in chapter[a:b]:
        return None, "the passage runs across two paragraphs"
    return (a, b), ""


def apply_fixes(chapter: str, fixes):
    applied, skipped, taken = [], [], []
    for f in fixes:
        if not f["rule"]:
            skipped.append({**f, "why": "the fix quoted no rule"})
            continue
        span, why = locate(chapter, f["passage"])
        if span is None:
            skipped.append({**f, "why": why})
            continue
        a, b = span
        new = " ".join(f["replacement"].split())
        if not new:  # a cut sentence takes one space with it
            if b < len(chapter) and chapter[b] == " ":
                b += 1
            elif a > 0 and chapter[a - 1] == " ":
                a -= 1
        if any(a < y and x < b for x, y in taken):
            skipped.append({**f, "why": "the passage overlaps another fix"})
            continue
        before = chapter[a - 1] if a > 0 else "\n"
        new = smart_quotes(before + new)[1:] if new else ""
        old = chapter[a:b]
        if new == old:
            skipped.append({**f, "why": "the fix changes nothing"})
            continue
        taken.append((a, b))
        applied.append({**f, "start": a, "end": b, "old": old, "new": new})
    applied.sort(key=lambda f: f["start"])
    return applied, skipped


def apply_to_text(chapter: str, applied) -> str:
    out, pos = [], 0
    for f in applied:
        out.append(chapter[pos:f["start"]])
        out.append(f["new"])
        pos = f["end"]
    out.append(chapter[pos:])
    return "\n\n".join(" ".join(p.split()) for p in "".join(out).split("\n\n"))


def _segments(par: str, par_start: int, fixes):
    """The paragraph as unchanged, deleted and inserted stretches, character by character."""
    segs, pos = [], 0
    for f in fixes:
        a, b = f["start"] - par_start, f["end"] - par_start
        old, new = par[a:b], f["new"]
        h = 0
        while h < min(len(old), len(new)) and old[h] == new[h]:
            h += 1
        t = 0
        while t < min(len(old), len(new)) - h and old[len(old) - 1 - t] == new[len(new) - 1 - t]:
            t += 1
        segs += [("eq", par[pos:a + h]), ("del", old[h:len(old) - t]), ("ins", new[h:len(new) - t])]
        pos = b - t
    segs.append(("eq", par[pos:]))
    return segs


def to_tracked_docx(paras, applied, title: str) -> bytes:
    """The reviewed chapter with every fix as a tracked change under the author's name."""
    doc = Document()
    stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev, start = 1, 0
    for par in paras:
        end = start + len(par)
        mine = [f for f in applied if start <= f["start"] and f["end"] <= end]
        if not mine:
            add_block(doc, par)
        else:
            p = doc.add_paragraph()
            for kind, text in _segments(par, start, mine):
                text = text.replace("*", "")
                if not text:
                    continue
                run = OxmlElement("w:r")
                t = OxmlElement("w:delText" if kind == "del" else "w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = text
                run.append(t)
                if kind == "eq":
                    p._p.append(run)
                    continue
                mark = OxmlElement("w:del" if kind == "del" else "w:ins")
                mark.set(qn("w:id"), str(rev))
                mark.set(qn("w:author"), AUTHOR)
                mark.set(qn("w:date"), stamp)
                rev += 1
                mark.append(run)
                p._p.append(mark)
        start = end + 2
    set_properties(doc, title)
    return save_bytes(doc)


def notes_docx(rv, title: str) -> bytes:
    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_heading("Questions for you", level=2)
    if not rv["questions"]:
        doc.add_paragraph("None.")
    for q in rv["questions"]:
        doc.add_paragraph(f"“{q['passage']}”", style="List Bullet")
        doc.add_paragraph(q["ask"])
    doc.add_heading("Fixes made", level=2)
    if not rv["applied"]:
        doc.add_paragraph("None.")
    for f in rv["applied"]:
        doc.add_paragraph(f"Was: “{f['old']}”", style="List Bullet")
        doc.add_paragraph(f"Now: “{f['new']}”" if f["new"] else "Now: cut.")
        doc.add_paragraph(f"Rests on: {f['rule']}")
    doc.add_heading("Fixes the app couldn’t place", level=2)
    if not rv["skipped"]:
        doc.add_paragraph("None.")
    for f in rv["skipped"]:
        doc.add_paragraph(f"“{f['passage']}”", style="List Bullet")
        doc.add_paragraph(f"Proposed: “{f['replacement']}”")
        doc.add_paragraph(f"Why it wasn’t made: {f['why']}.")
    set_properties(doc, title)
    return save_bytes(doc)


# ---------------------------------------------------------------- the Book File and its chapters

CHAPTERS_MARK = "CHAPTERS SO FAR"
CATALOGUE_MARK = "THE SENTENCE CATALOGUE"


def chapter_title(book_doc, n: int) -> str:
    """The chapter's line in the Book File's chapter list, such as '1. The Good Intent'."""
    inside = False
    for p in book_doc.paragraphs:
        t = p.text.strip()
        if t == "The chapter list":
            inside = True
            continue
        if inside and t == "Decision register":
            break
        if inside and re.match(rf"^{n}\.\s+\S", t):
            return t
    return f"{n}. Chapter {n}"


def insert_chapters_text(book_text: str, added) -> str:
    """The Book File's text with this run's chapters placed under CHAPTERS SO FAR."""
    if not added:
        return book_text
    block = []
    for title, text in added:
        block.append(title)
        block.extend(p.replace("*", "") for p in paragraphs_of(text))
    lines = book_text.split("\n")
    marks = [i for i, line in enumerate(lines) if line.strip() == CHAPTERS_MARK]
    if not marks:
        return "\n".join(lines + ["", CHAPTERS_MARK] + block)
    start = marks[0]
    end = next((i for i in range(start + 1, len(lines)) if lines[i].strip() == CATALOGUE_MARK), len(lines))
    return "\n".join(lines[:end] + block + lines[end:])


def _style(doc, *names):
    """The first of the named paragraph styles the document has, as a style object."""
    have = {s.name: s for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH}
    return next((have[n] for n in names if n in have), None)


def book_file_with_chapters(book_bytes: bytes, added, keep_catalogue: bool = True) -> bytes:
    """The attached Book File as a Word file, with this run's chapters under CHAPTERS SO FAR."""
    doc = Document(io.BytesIO(book_bytes))
    body = doc.element.body
    paras = doc.paragraphs
    catalogue = next((p for p in paras if p.text.strip() == CATALOGUE_MARK), None)
    if not any(p.text.strip() == CHAPTERS_MARK for p in paras):
        doc.add_paragraph(CHAPTERS_MARK, style=_style(doc, "Heading 2"))
        catalogue = None
    heading = _style(doc, "Heading 3")
    first_style = _style(doc, "First Paragraph", "Body Text", "Normal")
    body_style = _style(doc, "Body Text", "Normal")
    new = []
    for title, text in added:
        new.append(doc.add_paragraph(title, style=heading))
        for k, line in enumerate(paragraphs_of(text)):
            p = doc.add_paragraph(style=first_style if k == 0 else body_style)
            add_inline(p, line)
            new.append(p)
    if catalogue is not None:
        for p in new:
            catalogue._p.addprevious(p._p)
        if not keep_catalogue:
            el = catalogue._p
            while el is not None and el.tag != qn("w:sectPr"):
                nxt = el.getnext()
                body.remove(el)
                el = nxt
    cp = doc.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    return save_bytes(doc)


def book_file_names(book_name: str, first: int, last: int):
    stem = Path(book_name).stem
    span = f"Ch{first:02d}" if first == last else f"Ch{first:02d}_to_Ch{last:02d}"
    return f"{stem}_with_{span}.docx", f"{stem}_NO_CATALOGUE_with_{span}.docx"


# ---------------------------------------------------------------- the calls

EFFORTS = ["model default", "low", "medium", "high", "xhigh", "max"]


def stream_call(client, kwargs, status, preview=None, label="The draft"):
    """Runs the request on its own thread and keeps a progress line on the page while it runs."""
    state = {"text": "", "phase": "sending", "done": False, "error": None, "final": None}

    def work():
        try:
            with client.messages.stream(**kwargs) as stream:
                for event in stream:
                    kind = getattr(event, "type", "")
                    if kind == "content_block_start":
                        block = getattr(getattr(event, "content_block", None), "type", "")
                        state["phase"] = "writing" if block == "text" else "thinking"
                    elif kind == "content_block_delta":
                        delta = getattr(event, "delta", None)
                        if getattr(delta, "type", "") == "text_delta":
                            state["text"] += delta.text
                            state["phase"] = "writing"
                state["final"] = stream.get_final_message()
        except Exception as e:  # noqa: BLE001
            state["error"] = e
        finally:
            state["done"] = True

    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    started, shown = time.time(), 0
    while not state["done"]:
        mins, secs = divmod(int(time.time() - started), 60)
        clock = f"{mins}:{secs:02d}"
        if state["phase"] == "writing":
            status.info(f"{label}: writing, {word_count(state['text']):,} words so far, {clock} elapsed.")
            if preview is not None and len(state["text"]) - shown >= 400:
                shown = len(state["text"])
                preview.markdown(smart_quotes(state["text"][-3000:]))
        elif state["phase"] == "thinking":
            status.info(f"{label}: thinking, {clock} elapsed. Nothing appears until the writing starts. "
                        f"Leave the page alone.")
        else:
            status.info(f"{label}: sending the documents, {clock} elapsed.")
        time.sleep(2)
    worker.join()
    status.empty()
    if state["error"] is not None:
        raise state["error"]
    final = state["final"]
    text = "".join(getattr(b, "text", "") for b in final.content if getattr(b, "type", "") == "text")
    return text, final.usage, final.stop_reason


def request(model, max_tokens, system_text, docs_xml, message, extra, effort):
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
    config = dict(kwargs.get("output_config") or {})
    config.pop("effort", None)  # the effort comes from the sidebar, never from the JSON box
    if effort and effort != "model default":
        config["effort"] = effort
    if config:
        kwargs["output_config"] = config
    else:
        kwargs.pop("output_config", None)
    return kwargs


def review_chapter(client, model, max_tokens, effort, system_text, docs, docs_xml, chapter_no,
                   chapter_raw, instructions, extra, status):
    paras = paragraphs_of(smart_quotes(chapter_raw))
    chapter = "\n\n".join(paras)
    places_text, found = places_to_look(paras, docs[1][1])
    message = (f"<chapter>\n{chapter}\n</chapter>\n\n<places>\n{places_text}\n</places>\n\n"
               f"{instructions}\n\nReview Chapter {chapter_no}.")
    reply, usage, stop = stream_call(
        client, request(model, max_tokens, system_text, docs_xml, message, extra, effort), status,
        label=f"The review of Chapter {chapter_no}")
    fixes, questions = parse_review(reply)
    applied, skipped = apply_fixes(chapter, fixes)
    return dict(reply=reply, stop=stop, paras=paras, chapter=chapter, places=places_text,
                found=found, applied=applied, skipped=skipped, questions=questions,
                usage=usage_dict(usage))


def usage_dict(usage):
    return {
        "input": usage.input_tokens, "output": usage.output_tokens,
        "cache read": getattr(usage, "cache_read_input_tokens", 0) or 0,
        "cache write": getattr(usage, "cache_creation_input_tokens", 0) or 0,
    }


def tokens_line(u) -> str:
    return (f"Tokens: {u['input']:,} in, {u['output']:,} out, {u['cache read']:,} read from cache, "
            f"{u.get('cache write', 0):,} written to cache.")


def final_chapter_text(result) -> str:
    """The chapter as it goes into the Book File: the reviewed text when there is a review."""
    rv = result.get("review")
    if rv:
        return apply_to_text(rv["chapter"], rv["applied"])
    return "\n\n".join(paragraphs_of(smart_quotes(split_reply(result["reply"])[0])))


def write_chapters(client, model, settings, system_text, prompt_doc, book, chapters, runs, extra,
                   review_instructions, sent):
    """Drafts each chapter in turn. In a run of several chapters each one is drafted once, reviewed
    if review is on, and placed in the Book File before the next is written."""
    series = len(chapters) > 1
    runs = 1 if series else int(runs)
    added, done = [], []
    for n in chapters:
        book_text = insert_chapters_text(book["text"], added)
        docs = [prompt_doc, (book["name"], book_text)]
        docs_xml = documents_block(docs)
        message = f"Write Chapter {n}."
        title = chapter_title(Document(io.BytesIO(book["bytes"])), n)
        stopped = False
        for k in range(1, runs + 1):
            st.markdown(f"**Chapter {n}" + (f", run {k} of {runs}**" if runs > 1 else "**"))
            status, preview = st.empty(), st.empty()
            try:
                reply, usage, stop = stream_call(
                    client, request(model, settings["max_tokens"], system_text, docs_xml, message, extra,
                                    settings["effort"]),
                    status, preview, label=f"Chapter {n}")
            except Exception as e:  # noqa: BLE001
                st.error(f"The request for Chapter {n} failed: {e}")
                stopped = True
                break
            now = datetime.datetime.now()
            result = {
                "chapter": n, "run": k, "model": model, "reply": reply, "stop": stop,
                "sent": {**sent, "message": message,
                         "documents": [f"{nm} ({word_count(t):,} words)" for nm, t in docs]},
                "usage": usage_dict(usage), "time": now.strftime("%H:%M"), "stamp": now.strftime("%H%M"),
                "title": title,
            }
            if stop == "max_tokens":
                st.warning(f"Chapter {n} hit the token ceiling before it finished.")
            elif settings["review_on"]:
                try:
                    result["review"] = review_chapter(
                        client, model, settings["review_tokens"], settings["review_effort"], system_text,
                        docs, docs_xml, n, split_reply(reply)[0], review_instructions, extra, status)
                except Exception as e:  # noqa: BLE001
                    st.error(f"Chapter {n} is safe, but its review failed: {e}")
            preview.empty()
            st.session_state.results.append(result)
            done.append(result)
            if stop == "max_tokens":
                stopped = True
        if stopped and series:
            st.error(f"The run stopped at Chapter {n}. The chapters before it are under Results.")
            break
        if runs == 1 and not stopped:
            added.append((title, final_chapter_text(done[-1])))
    if added:
        st.session_state.book_out = {"name": book["name"], "bytes": book["bytes"], "added": added,
                                     "first": chapters[0], "last": chapters[0] + len(added) - 1}
    return done


# ---------------------------------------------------------------- the page

def remember_upload(uploaded, state_key):
    """Load an uploaded file into a text box once per new file."""
    if uploaded is None:
        return
    ident = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get(state_key + "_from") != ident:
        st.session_state[state_key] = file_to_text(uploaded)
        st.session_state[state_key + "_from"] = ident


def parse_chapters(raw: str):
    """'2' gives [2]; '2 to 4' or '2-4' gives [2, 3, 4]."""
    m = re.fullmatch(r"\s*(\d+)\s*(?:(?:-|–|to)\s*(\d+))?\s*", raw or "")
    if not m:
        return []
    first = int(m.group(1))
    last = int(m.group(2) or first)
    return list(range(first, last + 1)) if 1 <= first <= last else []


def result_name(r) -> str:
    if r.get("source"):
        return r["source"]
    stamp = f"_{r['stamp']}" if r.get("stamp") else ""
    return f"Chapter_{r['chapter']:02d}_API_{r['model']}{stamp}_run{r['run']}"


def result_files(r):
    """Every file the page offers for one result, as (file name, bytes)."""
    name = result_name(r)
    title = name.replace("_", " ")
    files = []
    if not r.get("source"):
        chap = split_reply(r["reply"])[0]
        files.append((name + ".docx", to_docx(chap, title)))
        files.append((name + "_full_reply.txt", smart_quotes(r["reply"]).encode("utf-8")))
    rv = r.get("review")
    if rv:
        files.append((name + "_reviewed_tracked.docx", to_tracked_docx(rv["paras"], rv["applied"], title + " reviewed")))
        files.append((name + "_reviewed_clean.docx",
                      to_docx(apply_to_text(rv["chapter"], rv["applied"]), title + " reviewed")))
        files.append((name + "_review_notes.docx", notes_docx(rv, title + " review notes")))
    return files


def book_out_files(bo):
    full, bare = book_file_names(bo["name"], bo["first"], bo["last"])
    return [(full, book_file_with_chapters(bo["bytes"], bo["added"], True)),
            (bare, book_file_with_chapters(bo["bytes"], bo["added"], False))]


def everything_zip(results, bo) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for r in results:
            for fname, data in result_files(r):
                z.writestr(fname, data)
        if bo:
            for fname, data in book_out_files(bo):
                z.writestr(fname, data)
    return out.getvalue()


def show_review(rv, files, i):
    st.write(
        f"Review: {len(rv['applied'])} fixes made, {len(rv['questions'])} questions for you, "
        f"{len(rv['skipped'])} fixes the app couldn’t place.  " + tokens_line(rv["usage"])
    )
    if rv["stop"] == "max_tokens":
        st.warning("The review hit its token ceiling, so it may be incomplete. Raise the review's "
                   "Max output tokens and review the chapter again in section 6.")
    labels = [("_reviewed_tracked.docx", "Reviewed chapter, tracked changes"),
              ("_reviewed_clean.docx", "Reviewed chapter, clean"),
              ("_review_notes.docx", "Review notes")]
    cols = st.columns(3)
    for col, (end, label) in zip(cols, labels):
        hit = next(((f, d) for f, d in files if f.endswith(end)), None)
        if hit:
            col.download_button(label, hit[1], file_name=hit[0], key=f"{end}{i}", on_click="ignore")
    if rv["questions"]:
        st.markdown("**Questions for you**")
        for q in rv["questions"]:
            st.markdown(f"- “{q['passage']}”  \n{q['ask']}")


def main():
    st.set_page_config(page_title="Chapter Draft Test", layout="wide")
    st.title("Chapter Draft Test")
    st.caption(
        "Sends the published claude.ai system prompt, the API's own fiction instructions, "
        "Prompt B, the Book File and “Write Chapter N.” Nothing else. "
        "Then, if review is on, it reviews the draft against the same documents."
    )

    if "system_prompt" not in st.session_state:
        st.session_state.system_prompt = default_text("system_prompt.txt")
    if "instructions" not in st.session_state:
        st.session_state.instructions = default_text("api_instructions.txt", FICTION_INSTRUCTIONS)
    if "review_instructions" not in st.session_state:
        st.session_state.review_instructions = default_text("review_instructions.txt", REVIEW_INSTRUCTIONS)
    st.session_state.setdefault("results", [])
    st.session_state.setdefault("book_out", None)

    with st.sidebar:
        st.header("Model")
        choice = st.selectbox("Model", MODELS, index=0)
        model = st.text_input("Model string", "") if choice == "Other" else choice
        effort = st.selectbox("Effort for the draft", EFFORTS, index=EFFORTS.index("xhigh"),
                              help="Your web chats run at Extra, which is xhigh.")
        max_tokens = st.number_input("Max output tokens", 4000, 128000, 128000, 1000,
                                     help="Thinking counts toward this. At xhigh a draft can use 80,000.")
        runs = st.number_input("Runs of the same chapter", 1, 3, 1,
                               help="For a single chapter. In a run of several chapters each is drafted once.")
        st.header("Review")
        review_on = st.checkbox("Review each draft after it’s written", value=True)
        review_effort = st.selectbox("Effort for the review", EFFORTS, index=0,
                                     help="The model's default is medium, which is enough to check a chapter.")
        review_tokens = st.number_input("Max output tokens for a review", 4000, 64000, 32000, 1000)
        with st.expander("Extra request parameters (JSON)"):
            extra_raw = st.text_area(
                "Merged into every request. Leave empty to send the model's defaults. "
                "The effort comes from the two boxes above, so an effort in here is ignored.",
                "", height=80,
            )
        try:
            key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:  # noqa: BLE001  (no secrets file)
            key = ""
        if not key:
            key = st.text_input("Anthropic API key", type="password")

    st.subheader("1. System prompt")
    ss = st.session_state
    if (HERE / "system_prompt.txt").exists():
        st.success("Loaded from system_prompt.txt beside the app.")
    else:
        untouched = not ss.system_prompt.strip() or ss.system_prompt == ss.get("sp_fetched_text")
        if model and untouched and ss.get("sp_fetched_model") != model:
            ss.sp_fetched_model = model
            try:
                ss.sp_fetched_date, ss.sp_fetched_text = fetch_system_prompt(model)
                ss.system_prompt = ss.sp_fetched_text
                ss.sp_error = ""
            except Exception as e:  # noqa: BLE001
                ss.sp_fetched_text, ss.sp_error = "", str(e)
                ss.system_prompt = ""
        if ss.get("sp_fetched_text") and ss.system_prompt == ss.sp_fetched_text:
            st.success(f"Fetched from Anthropic's page: the claude.ai system prompt for "
                       f"{ss.sp_fetched_model}, the version of {ss.sp_fetched_date}. Nothing to paste.")
        elif ss.get("sp_error") and not ss.system_prompt.strip():
            st.warning(f"The app couldn't fetch the system prompt from Anthropic's page "
                       f"({ss.sp_error}). Copy it from [the page]({SYSTEM_PROMPT_PAGE.format(model=model)}) "
                       f"and paste it here or upload it.")
            if st.button("Try fetching again"):
                ss.sp_fetched_model = None
                st.rerun()
    remember_upload(st.file_uploader("Upload system prompt", type=["txt", "md", "docx"], key="sp_up"), "system_prompt")
    st.text_area("System prompt", key="system_prompt", height=160)

    st.subheader("2. Instructions for the API")
    st.markdown(
        "These are the API's own fiction instructions, from API_Fiction_Instructions_v2.docx "
        "in WritingApp. They're already filled in. Your instructions in Settings stay with "
        "your editorial work and don't go into this test."
    )
    remember_upload(st.file_uploader("Upload other instructions", type=["txt", "md", "docx"], key="in_up"), "instructions")
    st.text_area("Instructions", key="instructions", height=150)

    st.subheader("3. The two documents")
    c1, c2 = st.columns(2)
    prompt_file = c1.file_uploader("Prompt B (.docx)", type=["docx"], key="pb")
    book_file = c2.file_uploader("Book File (.docx)", type=["docx"], key="bf")

    st.subheader("4. The chapters")
    raw = st.text_input("Chapter, or a run of chapters", "", placeholder="e.g. 2, or 2 to 4")
    chapters = parse_chapters(raw)
    if raw.strip() and not chapters:
        st.error("Give one chapter number, or a first and last, e.g. 2 to 4.")
    if chapters:
        st.code("\n".join(f"Write Chapter {n}." for n in chapters))
    if len(chapters) > 1:
        st.markdown(
            "Each chapter is drafted once and reviewed if review is on. The reviewed chapter goes under "
            "CHAPTERS SO FAR before the next one is written. At xhigh, allow about twenty minutes a chapter, "
            "and keep this tab open and the computer awake until the run finishes."
        )

    st.subheader("5. Review instructions")
    with st.expander("The instructions the reviewer reads. They're already filled in."):
        remember_upload(st.file_uploader("Upload other review instructions", type=["txt", "md", "docx"],
                                         key="rv_up"), "review_instructions")
        st.text_area("Review instructions", key="review_instructions", height=300)

    settings = dict(max_tokens=int(max_tokens), effort=effort, review_on=review_on,
                    review_tokens=int(review_tokens), review_effort=review_effort)

    def setup():
        try:
            extra = json.loads(extra_raw) if extra_raw.strip() else {}
        except json.JSONDecodeError as e:
            st.error(f"Extra parameters aren't valid JSON: {e}")
            return None
        system_text = fill_date(st.session_state.system_prompt.strip())
        instructions = st.session_state.instructions.strip()
        if instructions:
            system_text += f"\n\n<userPreferences>\n{instructions}\n</userPreferences>"
        prompt_doc = (prompt_file.name, docx_to_text(prompt_file))
        book_bytes = book_file.getvalue()
        book = {"name": book_file.name, "bytes": book_bytes, "text": docx_to_text(io.BytesIO(book_bytes))}
        sent = {
            "model": model,
            "max_tokens": int(max_tokens),
            "effort": effort,
            "review effort": review_effort if review_on else "no review",
            "system prompt (characters)": len(system_text),
            "instructions": instructions.split("\n", 1)[0] if instructions else "none",
            "extra parameters": extra,
        }
        client = Anthropic(api_key=key, timeout=1800, max_retries=2)
        return extra, system_text, prompt_doc, book, sent, client

    ready = all([model, key, st.session_state.system_prompt.strip(), prompt_file, book_file, chapters])
    if not ready:
        st.info("Fill in the system prompt, attach both documents and give the chapter number.")
    label = "Write the chapters" if len(chapters) > 1 else "Write the chapter"
    if st.button(label, type="primary", disabled=not ready):
        got = setup()
        if got:
            extra, system_text, prompt_doc, book, sent, client = got
            st.session_state.book_out = None
            write_chapters(client, model, settings, system_text, prompt_doc, book, chapters, runs, extra,
                           st.session_state.review_instructions.strip(), sent)

    st.subheader("6. Review a chapter you already have")
    st.markdown(
        "Attach a chapter written elsewhere, e.g. in an incognito window, as a Word or text file. "
        "It's reviewed against the documents in section 3, as the first chapter named in section 4."
    )
    own = st.file_uploader("Chapter to review (.docx or .txt)", type=["docx", "txt"], key="own")
    if st.button("Review this chapter", disabled=not (ready and own)):
        got = setup()
        if got:
            extra, system_text, prompt_doc, book, sent, client = got
            n = chapters[0]
            docs = [prompt_doc, (book["name"], book["text"])]
            text = split_reply(chapter_from_upload(own))[0]
            status = st.empty()
            try:
                rv = review_chapter(client, model, settings["review_tokens"], review_effort, system_text,
                                    docs, documents_block(docs), n, text,
                                    st.session_state.review_instructions.strip(), extra, status)
                now = datetime.datetime.now()
                st.session_state.results.append({
                    "chapter": n, "run": 0, "model": model, "reply": text, "stop": "",
                    "sent": {**sent, "message": f"Review Chapter {n}."},
                    "usage": rv["usage"], "time": now.strftime("%H:%M"), "stamp": now.strftime("%H%M"),
                    "source": Path(own.name).stem, "review": rv,
                })
            except Exception as e:  # noqa: BLE001
                st.error(f"The review failed: {e}")
            status.empty()

    results = st.session_state.results
    bo = st.session_state.book_out
    if results:
        st.subheader("Results")
        st.caption("Downloading never restarts the app, so a download is safe while nothing is running.")
        a1, a2 = st.columns(2)
        a1.download_button("Download everything (.zip)", everything_zip(results, bo),
                           file_name=f"Chapter_Draft_Test_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                           key="zip_all", on_click="ignore", type="primary")
        if a2.button("Clear results"):
            st.session_state.results = []
            st.session_state.book_out = None
            st.rerun()
    if bo:
        names = [t for t, _ in bo["added"]]
        span = f"Chapter {bo['first']}" if bo["first"] == bo["last"] else f"Chapters {bo['first']} to {bo['last']}"
        with st.expander(f"The Book File with {span} added", expanded=True):
            st.write("Added under CHAPTERS SO FAR, as reviewed: " + "; ".join(names) + ".")
            b1, b2 = st.columns(2)
            for col, (fname, data), label in zip((b1, b2), book_out_files(bo),
                                                 ("Book File with the chapters", "The same, NO_CATALOGUE")):
                col.download_button(label, data, file_name=fname, key=f"bo_{label}", on_click="ignore")
    for i, r in enumerate(reversed(results)):
        chap, notes = split_reply(r["reply"])
        files = result_files(r)
        if r.get("source"):
            label = f"Chapter {r['chapter']}, review of {r['source']}, {r['model']}, {r['time']}"
        else:
            label = f"Chapter {r['chapter']}, run {r['run']}, {r['model']}, {r['time']}"
        with st.expander(f"{label}: {word_count(chap):,} words", expanded=(i == 0)):
            if not r.get("source"):
                st.write(f"Stop reason: {r['stop']}.  " + tokens_line(r["usage"]))
                if r["stop"] == "max_tokens":
                    st.warning("The reply hit the token ceiling. Raise Max output tokens and run again.")
                d1, d2 = st.columns(2)
                d1.download_button("Chapter as Word file", files[0][1], file_name=files[0][0],
                                   key=f"dx{i}", on_click="ignore")
                d2.download_button("Whole reply as text", files[1][1], file_name=files[1][0],
                                   key=f"tx{i}", on_click="ignore")
            rv = r.get("review")
            if rv:
                show_review(rv, files, i)
            with st.popover("What was sent"):
                st.json(r["sent"])
            tabs = st.tabs(["The draft", "After review", "Review details"] if rv else ["The draft"])
            with tabs[0]:
                st.markdown(smart_quotes(chap))
                if notes:
                    st.markdown("**The drafter's notes**")
                    st.markdown(smart_quotes(notes))
            if rv:
                with tabs[1]:
                    st.markdown(apply_to_text(rv["chapter"], rv["applied"]))
                with tabs[2]:
                    st.markdown("**Fixes made**")
                    for f in rv["applied"]:
                        st.markdown(f"- Was: “{f['old']}”  \nNow: "
                                    + (f"“{f['new']}”" if f["new"] else "cut.")
                                    + f"  \nRests on: {f['rule']}")
                    if rv["skipped"]:
                        st.markdown("**Fixes the app couldn’t place**")
                        for f in rv["skipped"]:
                            st.markdown(f"- “{f['passage']}”  \n{f['why']}.")
                    st.markdown("**The places the app listed for the reviewer**")
                    st.text(rv["places"])


if __name__ == "__main__":
    main()
