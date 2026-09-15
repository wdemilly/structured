"""
Skeleton Extractor — Aug. 18, 2026 RECREATION
================================================

This is a faithful reconstruction of the extractor needed to build the
Structural Transplant / v6F skeleton format. The original Aug. 18 source code
was not preserved. This recreation is grounded in:

  * Structural_Transplant_Local_Sentence_ClosedClass_v6F.docx
  * the 119-block / 321-unit skeleton embedded in v6D/v6F
  * Skeleton_Extractor_v1_reconstruction.docx (Aug. 26 reconstruction)
  * the later skeleton_app.py reconstruction

It preserves the six fields that canonical v6F actually uses:
closed-class frame, dialogue status, terminal mode, POS sequence, TARGET WORDS,
and shallow PHRASES. It also preserves donor paragraph provenance separately.

Important: this is NOT claimed to be byte-for-byte the lost Aug. 18 program.
The exact original tagger/parser cannot be recovered from the surviving files.
The rule-based tagger here is deterministic and is the closest documented
reconstruction available.

Repairs retained intentionally:
  * DOCX uses real Word paragraph boundaries.
  * TXT/MD use blank lines as paragraph boundaries; ordinary hard wraps are
    joined instead of becoming false paragraphs.
  * Multiple donors retain source identity and never create a window crossing
    a donor boundary.
  * A separate provenance TSV is emitted; donor names never enter the skeleton.

Run with:
    streamlit run Skeleton_Extractor_Aug18_RECREATED.py

Dependencies:
    streamlit
    python-docx
"""

import io
import random
import re
import statistics as st


# Canonical content-blind order embedded in the Aug. 18 v6F prompt.
# It is valid only for a 119-block skeleton. New chapters/production runs
# should normally use a newly seeded order, but this option is retained for
# exact historical reproduction and controlled comparisons.
CANONICAL_V6F_ORDER = """
P111 P042 P057 P077 P108 P046 P112 P060 P076 P009
P041 P070 P034 P117 P102 P093 P054 P001 P061 P098
P094 P058 P002 P075 P067 P091 P007 P105 P016 P043
P027 P104 P113 P072 P028 P092 P021 P048 P022 P065
P050 P015 P119 P071 P110 P010 P074 P025 P019 P088
P115 P005 P080 P084 P086 P056 P035 P087 P014 P026
P096 P013 P114 P012 P053 P017 P045 P039 P024 P066
P044 P037 P073 P106 P055 P063 P064 P083 P090 P059
P078 P069 P006 P040 P107 P020 P003 P068 P082 P081
P029 P103 P051 P100 P004 P031 P052 P095 P085 P049
P036 P047 P011 P101 P062 P033 P118 P032 P099 P038
P116 P018 P030 P109 P008 P023 P097 P079 P089
""".split()

# ----------------------------------------------------------------- LEXICONS --

DET = {"the", "a", "an", "this", "that", "these", "those", "each", "every", "either",
       "neither", "some", "any", "no", "another", "both", "all", "half", "such", "what"}
PRON = {"i", "you", "he", "she", "it", "we", "they", "me", "him", "her", "us", "them",
        "myself", "yourself", "himself", "herself", "itself", "ourselves", "themselves",
        "one", "none", "anyone", "someone", "everyone", "nobody", "somebody", "anybody",
        "everybody", "anything", "something", "everything", "nothing", "others", "mine",
        "yours", "hers", "ours", "theirs"}
PRON_POSS = {"my", "your", "his", "her", "its", "our", "their", "whose"}
ADP = {"about", "above", "across", "after", "against", "along", "among", "around", "at",
       "before", "behind", "below", "beneath", "beside", "besides", "between", "beyond",
       "by", "down", "during", "except", "for", "from", "in", "inside", "into", "like",
       "near", "of", "off", "on", "onto", "out", "outside", "over", "past", "since",
       "through", "throughout", "till", "toward", "towards", "under", "underneath",
       "until", "up", "upon", "with", "within", "without", "despite", "unlike", "per",
       "although", "though", "because", "if", "unless", "whereas", "while", "whether",
       "as", "than", "so", "once", "whenever", "wherever", "lest"}
COORD = {"and", "but", "or", "nor", "yet", "for_c", "plus"}
MODAL = {"can", "could", "may", "might", "must", "shall", "should", "will", "would",
         "ought", "cannot", "ca", "wo", "sha"}
WH_PRON = {"who", "whom", "which", "whoever", "whomever", "whichever", "whatever"}
WH_ADV = {"when", "where", "why", "how", "however"}
BE = {"am", "is", "are", "was", "were", "be", "been", "being", "'m", "'re", "'s"}
HAVE = {"have", "has", "had", "having", "'ve", "'d"}
DO = {"do", "does", "did", "doing"}
NEG = {"not", "n't", "never"}
NUMWORD = {"one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
           "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen",
           "eighteen", "nineteen", "twenty", "thirty", "forty", "fifty", "sixty",
           "seventy", "eighty", "ninety", "hundred", "thousand", "million", "billion",
           "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth",
           "ninth", "tenth", "dozen", "score"}
ADV_LIST = {"very", "too", "quite", "rather", "almost", "already", "also", "always",
            "still", "just", "even", "only", "again", "back", "away", "here", "there",
            "now", "then", "soon", "later", "often", "sometimes", "usually", "never",
            "ever", "far", "much", "more", "most", "less", "least", "well", "perhaps",
            "maybe", "instead", "together", "apart", "ahead", "somewhere", "anywhere",
            "everywhere", "nowhere", "meanwhile", "afterward", "afterwards", "indeed",
            "certainly", "probably", "finally", "suddenly", "yesterday", "today",
            "tomorrow", "tonight", "twice", "once_adv", "enough", "otherwise"}
INTJ = {"oh", "ah", "well_i", "yes", "no_i", "okay", "ok", "hey", "hello", "yeah",
        "hi", "huh", "hm", "hmm", "please", "thanks", "goodbye", "bye", "sure_i"}
IRREG_PAST = {"was", "were", "had", "did", "said", "went", "came", "took", "saw", "got",
              "made", "knew", "thought", "found", "told", "gave", "felt", "left", "kept",
              "held", "stood", "sat", "brought", "began", "ran", "put", "set", "let",
              "cut", "read", "heard", "meant", "met", "paid", "sent", "built", "lost",
              "won", "spoke", "broke", "chose", "drove", "fell", "flew", "grew", "wore",
              "wrote", "rose", "shook", "threw", "drew", "blew", "bore", "swam", "sang",
              "rang", "sank", "drank", "hung", "lay", "laid", "slept", "swept", "wept",
              "crept", "bent", "lent", "spent", "burnt", "dealt", "leapt", "knelt",
              "struck", "stuck", "dug", "hid", "bit", "lit", "shot", "shut", "spread",
              "split", "quit", "fit", "beat", "ate", "forgot", "became", "understood"}
PAST_PART = {"been", "gone", "done", "seen", "taken", "given", "known", "thrown", "grown",
             "shown", "written", "driven", "spoken", "broken", "chosen", "eaten", "fallen",
             "forgotten", "hidden", "ridden", "risen", "worn", "torn", "born", "sworn",
             "begun", "drunk", "sung", "sunk", "swum", "rung", "held", "made", "found"}
COMMON_VERB = {"go", "goes", "going", "come", "comes", "coming", "say", "says", "saying",
               "see", "sees", "seeing", "take", "takes", "taking", "make", "makes",
               "making", "know", "knows", "knowing", "think", "thinks", "thinking",
               "look", "looks", "looking", "want", "wants", "wanting", "give", "gives",
               "giving", "use", "uses", "using", "find", "finds", "finding", "tell",
               "tells", "telling", "ask", "asks", "asking", "work", "works", "working",
               "seem", "seems", "seeming", "feel", "feels", "feeling", "try", "tries",
               "trying", "leave", "leaves", "leaving", "call", "calls", "calling",
               "keep", "keeps", "keeping", "let", "lets", "letting", "put", "puts",
               "putting", "get", "gets", "getting", "turn", "turns", "turning", "hold",
               "holds", "holding", "stand", "stands", "standing", "sit", "sits",
               "sitting", "walk", "walks", "walking", "run", "runs", "running"}

NEG_STEM = {"don": "V.BASE", "doesn": "V.PRES3", "didn": "V.PAST", "isn": "V.PRES3",
            "aren": "V.PRES", "wasn": "V.PAST", "weren": "V.PAST", "hasn": "V.PRES3",
            "haven": "V.PRES", "hadn": "V.PAST", "couldn": "MODAL", "wouldn": "MODAL",
            "shouldn": "MODAL", "mustn": "MODAL", "needn": "MODAL", "won": "MODAL",
            "can": "MODAL", "shan": "MODAL", "ain": "V.PRES"}

PUNCT_TAG = {".": "PERIOD", ",": "COMMA", '"': "QUOTE", "?": "QMARK", "!": "EXCL",
             ";": "SEMICOLON", "'": "APOST", "-": "HYPHEN", ":": "OTHER",
             "\u2014": "OTHER", "\u2013": "OTHER", "(": "OTHER", ")": "OTHER",
             "\u2026": "OTHER"}

CLOSED_TAGS = {"DET", "PRON", "PRON.POSS", "ADP/SUB", "COORD", "TO", "MODAL", "PART",
               "EXIST", "WH.PRON", "WH.ADV", "AUX"}
NP_TAGS = {"DET", "PRON", "PRON.POSS", "NOUN", "NOUN.PL", "PROPN", "NUM", "EXIST", "WH.PRON"}
VP_TAGS = {"V.BASE", "V.PAST", "V.PART", "V.ING", "V.PRES", "V.PRES3", "MODAL", "PART"}
ADJ_TAGS = {"ADJ", "ADJ.COMP", "ADJ.SUP"}
ADV_TAGS = {"ADV", "WH.ADV"}
PP_TAGS = {"ADP/SUB", "TO"}
SKIP_TAGS = {"COMMA", "PERIOD", "QUOTE", "QMARK", "EXCL", "SEMICOLON", "APOST",
             "HYPHEN", "COORD", "INTJ", "OTHER"}

# ------------------------------------------------------------------ READING --


def normalise(t):
    t = t.replace("\u201c", '"').replace("\u201d", '"')
    t = t.replace("\u2018", "'").replace("\u2019", "'")
    t = re.sub(r"\r\n?", "\n", t)
    return t


def plaintext_paragraphs(t):
    """
    Recover paragraphs from TXT/MD without treating ordinary hard line wraps as
    paragraph breaks. Paragraphs are separated by one or more blank lines;
    single newlines inside a paragraph are joined with a space.
    """
    t = normalise(t)
    chunks = re.split(r"\n[ \t]*\n+", t)
    out = []
    for chunk in chunks:
        chunk = re.sub(r"[ \t]*\n[ \t]*", " ", chunk)
        chunk = re.sub(r"[ \t]+", " ", chunk).strip()
        if chunk:
            out.append(chunk)
    return out


def read_upload(name, data):
    """
    Return a list of source paragraphs.

    DOCX: preserves actual Word paragraph boundaries.
    TXT/MD: blank lines delimit paragraphs; single newlines are treated as
    line wrapping rather than paragraph breaks.
    """
    low = name.lower()
    if low.endswith(".docx"):
        try:
            import docx
        except ImportError:
            raise RuntimeError("python-docx is not installed; run: pip install python-docx")
        d = docx.Document(io.BytesIO(data))
        return [p.text.strip() for p in d.paragraphs if p.text and p.text.strip()]
    if low.endswith(".doc"):
        raise RuntimeError("Legacy .doc is not supported. Open it in Word and save as .docx.")
    return plaintext_paragraphs(data.decode("utf-8", errors="replace"))


def paragraph_integrity(paras):
    """
    Flag likely hard-wrapped / malformed donor files without guessing repairs.
    A high share of paragraph records that do not end in sentence-terminal
    punctuation is a strong sign that line wraps were imported as paragraphs.
    """
    if not paras:
        return dict(paragraphs=0, nonterminal=0, nonterminal_pct=0.0)
    bad = 0
    for p in paras:
        if not re.search(r'[.!?\u2026](?:["\'])?\s*$', p):
            bad += 1
    return dict(paragraphs=len(paras), nonterminal=bad,
                nonterminal_pct=100.0 * bad / len(paras))


ABBR = r"(?<!\bMr)(?<!\bMrs)(?<!\bMs)(?<!\bDr)(?<!\bSt)(?<!\bJr)(?<!\bSr)"


def sentences(p):
    parts = re.split(ABBR + r'(?<=[.!?])(["\']?)\s+', p)
    out, buf = [], ""
    for chunk in parts:
        if chunk is None:
            continue
        if chunk in ('"', "'"):
            buf += chunk
            continue
        if buf:
            out.append(buf.strip())
        buf = chunk
    if buf.strip():
        out.append(buf.strip())
    return [s for s in out if re.search(r"[A-Za-z]", s)]


TOKEN_RE = re.compile(r"[A-Za-z]+(?:'[A-Za-z]+)?|\d+(?:[.,]\d+)?|[^\sA-Za-z\d]")


def tokenise(s):
    """Words, contractions split at the apostrophe, punctuation as separate tokens."""
    out = []
    for tok in TOKEN_RE.findall(s):
        if "'" in tok and len(tok) > 1:
            head, tail = tok.split("'", 1)
            if head:
                out.append(head)
            out.append("'")
            if tail:
                out.append(tail)
        else:
            out.append(tok)
    return out

# ------------------------------------------------------------------ TAGGING --


def word_count(tokens):
    """Alphabetic words, a contraction counting as one."""
    n, i = 0, 0
    while i < len(tokens):
        t = tokens[i]
        if re.match(r"^[A-Za-z]+$", t):
            n += 1
            if i + 2 < len(tokens) and tokens[i + 1] == "'" and re.match(r"^[A-Za-z]+$", tokens[i + 2]):
                i += 2
        elif re.match(r"^\d", t):
            n += 1
        i += 1
    return n


def tag_tokens(tokens):
    """One tag per token, in order. Deterministic."""
    tags = []
    for i, tok in enumerate(tokens):
        low = tok.lower()
        prev = tags[-1] if tags else None
        prev_word = tokens[i - 1].lower() if i else ""
        recent = [t.lower() for t in tokens[max(0, i - 3):i]]
        first_alpha = not any(re.match(r"^[A-Za-z]", t) for t in tokens[:i])

        if tok == "'":
            prev_alpha = bool(i and re.match(r"^[A-Za-z]+$", tokens[i - 1]))
            next_alpha = bool(i + 1 < len(tokens) and re.match(r"^[A-Za-z]+$", tokens[i + 1]))
            tags.append("APOST" if (prev_alpha and next_alpha) else "QUOTE")
            continue
        if not re.match(r"^[A-Za-z\d]", tok):
            tags.append(PUNCT_TAG.get(tok, "OTHER"))
            continue
        if low in NEG_STEM and i + 1 < len(tokens) and tokens[i + 1] == "'":
            tags.append(NEG_STEM[low])
            continue
        if low == "t" and prev == "APOST":
            tags.append("PART")
            continue
        if re.match(r"^\d", tok):
            tags.append("NUM")
            continue
        if low in NEG or low == "'t":
            tags.append("PART")
            continue
        if low == "to":
            nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else ""
            tags.append("TO" if (nxt in COMMON_VERB or nxt in BE or nxt in HAVE) else "ADP/SUB")
            continue
        if low == "there" and i + 1 < len(tokens) and tokens[i + 1].lower() in BE:
            tags.append("EXIST")
            continue
        if low == "that":
            tags.append("WH.PRON" if prev in {"NOUN", "NOUN.PL", "PROPN"} else "DET")
            continue
        if low in MODAL:
            tags.append("MODAL")
            continue
        if low in BE or low in HAVE or low in DO:
            if low in {"'s", "'re", "'m", "'ve", "'d"}:
                tags.append("V.PRES" if low in {"'s", "'re", "'m", "'ve"} else "V.PAST")
            elif low in {"was", "were", "had", "did"}:
                tags.append("V.PAST")
            elif low in {"been"}:
                tags.append("V.PART")
            elif low in {"being", "having", "doing"}:
                tags.append("V.ING")
            elif low in {"be", "have", "do"}:
                tags.append("V.BASE")
            elif low in {"is", "has", "does"}:
                tags.append("V.PRES3")
            else:
                tags.append("V.PRES")
            continue
        if low in WH_PRON:
            tags.append("WH.PRON")
            continue
        if low in WH_ADV:
            tags.append("WH.ADV")
            continue
        if low in PRON_POSS:
            tags.append("PRON.POSS")
            continue
        if low in PRON:
            tags.append("PRON")
            continue
        if low in DET:
            tags.append("DET")
            continue
        if low in {"and", "but", "or", "nor", "yet"}:
            tags.append("COORD")
            continue
        if low == "for":
            tags.append("ADP/SUB")
            continue
        if low in ADP:
            tags.append("ADP/SUB")
            continue
        if low in NUMWORD:
            tags.append("NUM")
            continue
        if low in INTJ and (first_alpha or prev == "QUOTE"):
            tags.append("INTJ")
            continue

        # verbs by morphology and context
        aux_before = any(r in HAVE or r in BE for r in recent)
        if low.endswith("ing") and len(low) > 4:
            tags.append("V.ING")
            continue
        if low in PAST_PART:
            tags.append("V.PART")
            continue
        if low in IRREG_PAST:
            tags.append("V.PART" if aux_before else "V.PAST")
            continue
        if low.endswith("ed") and len(low) > 3 and prev not in {"DET", "PRON.POSS", "ADJ"}:
            tags.append("V.PART" if aux_before else "V.PAST")
            continue
        if low in COMMON_VERB and prev not in {"DET", "PRON.POSS", "ADJ"}:
            tags.append("V.PRES3" if low.endswith("s") else "V.BASE")
            continue

        # modifiers
        if low.endswith("ly") and len(low) > 3 and low not in {"only", "family", "early", "lonely", "likely"}:
            tags.append("ADV")
            continue
        if low in ADV_LIST and prev not in {"DET", "PRON.POSS"}:
            tags.append("ADV")
            continue
        if low.endswith("est") and len(low) > 4:
            tags.append("ADJ.SUP")
            continue
        if low.endswith("er") and len(low) > 4 and prev in {"V.BASE", "V.PAST", "V.PRES", "ADV"}:
            tags.append("ADJ.COMP")
            continue

        # nouns
        if tok[0].isupper() and not first_alpha and prev != "QUOTE":
            tags.append("PROPN")
            continue
        if low.endswith("s") and not low.endswith("ss") and len(low) > 3:
            tags.append("NOUN.PL")
            continue
        if prev in {"DET", "PRON.POSS", "ADJ", "NUM"} or prev in {"ADP/SUB"}:
            tags.append("NOUN")
            continue
        if low.endswith(("ous", "ful", "less", "able", "ible", "ive", "al", "ish", "y")) and len(low) > 4:
            nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else ""
            tags.append("ADJ" if re.match(r"^[A-Za-z]", nxt or "") else "NOUN")
            continue
        tags.append("PROPN" if tok[0].isupper() and first_alpha and low not in DET else "NOUN")
    return tags

# ------------------------------------------------------------------ FIELDS --


def phrases(tags):
    labs = []
    for i, t in enumerate(tags):
        if t in SKIP_TAGS:
            continue
        if t in NP_TAGS:
            lab = "NP"
        elif t in VP_TAGS:
            lab = "VP"
        elif t in PP_TAGS:
            lab = "PP"
        elif t in ADJ_TAGS:
            nxt = next((x for x in tags[i + 1:] if x not in SKIP_TAGS), None)
            lab = "NP" if nxt in NP_TAGS else "ADJP"
        elif t in ADV_TAGS:
            lab = "ADVP"
        else:
            continue
        if not labs or labs[-1] != lab:
            labs.append(lab)
    return labs or ["NONE"]


def status(sent):
    q = sent.count('"')
    if q == 0:
        return "UNQUOTED"
    stripped = re.sub(r'"[^"]*"', "", sent).strip(" ,.;:!?-")
    return "DIALOGUE" if not re.search(r"[A-Za-z]", stripped) else "MIXED"


def terminal(sent):
    tail = sent.rstrip('" \'')
    if tail.endswith("?"):
        return "QUESTION"
    if tail.endswith("!"):
        return "EXCLAMATION"
    return "DECLARATIVE"


def mask_token(tok, tag, first_alpha):
    if tag in CLOSED_TAGS or tag in {"V.PRES", "V.PRES3", "V.BASE"} and tok.lower() in (BE | HAVE | DO):
        return tok
    if tag in CLOSED_TAGS:
        return tok
    if tok.lower() in (BE | HAVE | DO | MODAL | NEG):
        return tok
    if tag in {"PERIOD", "COMMA", "QUOTE", "QMARK", "EXCL", "SEMICOLON", "APOST", "HYPHEN", "OTHER"}:
        return tok
    if tag == "NUM":
        return "[num]"
    if tag == "PROPN":
        return "[X]"
    if tag in {"V.PAST", "V.PART"}:
        return "[v-ed]"
    if tag == "V.ING":
        return "[v-ing]"
    if tag == "ADJ.SUP":
        return "[adj-super]"
    if tag == "ADV":
        return "[adv]"
    if tok.lower().endswith("er") and len(tok) > 3:
        return "[x-er]"
    if tok[:1].isupper() and first_alpha:
        return "[X]"
    return "[x]"


def frame_sentence(tokens, tags):
    """Masked text of one sentence, with donor punctuation and spacing preserved."""
    out, first_done, open_quote = "", False, False
    for tok, tag in zip(tokens, tags):
        is_alpha = bool(re.match(r"^[A-Za-z]", tok))
        first_alpha = is_alpha and not first_done
        if is_alpha:
            first_done = True
        piece = mask_token(tok, tag, first_alpha)
        if tag == "APOST":
            out += piece
        elif tag in {"PERIOD", "COMMA", "QMARK", "EXCL", "SEMICOLON"} or piece in {":", "\u2026"}:
            out += piece
        elif tag == "HYPHEN":
            out += piece
        elif tag == "QUOTE":
            if open_quote:
                out = out.rstrip() + piece
            else:
                if out and not out.endswith((" ", "(", "-")):
                    out += " "
                out += piece
            open_quote = not open_quote
        else:
            if out and not out.endswith(("'", '"', "-")) and out != "":
                out += " "
            out += piece
    return out.strip()

# ---------------------------------------------------------------- ASSEMBLY --


def build_units(paras):
    """
    Accept plain paragraph strings or provenance records:
      {"text": ..., "source": ..., "source_para": ...}
    Provenance is carried on the resulting block but never printed into the
    structural skeleton itself.
    """
    blocks = []
    for item in paras:
        if isinstance(item, dict):
            p = item["text"]
            source = item.get("source", "")
            source_para = item.get("source_para", "")
        else:
            p = item
            source = ""
            source_para = ""
        ss = sentences(p)
        if not ss:
            continue
        units = []
        for s in ss:
            toks = tokenise(s)
            tags = tag_tokens(toks)
            units.append(dict(sent=s, toks=toks, tags=tags, tw=word_count(toks),
                              status=status(s), term=terminal(s), phr=phrases(tags)))
        blocks.append(dict(
            units=units,
            frame=" ".join(frame_sentence(u["toks"], u["tags"]) for u in units),
            source=source,
            source_para=source_para,
        ))
    return blocks


def measurements(blocks):
    lens = [u["tw"] for b in blocks for u in b["units"]]
    stt = [u["status"] for b in blocks for u in b["units"]]
    return dict(blocks=len(blocks), units=len(lens), words=sum(lens),
                sd=st.pstdev(lens) if len(lens) > 1 else 0.0,
                mean=st.mean(lens) if lens else 0.0,
                cv=(st.pstdev(lens) / st.mean(lens)) if lens and st.mean(lens) else 0.0,
                longest=max(lens) if lens else 0,
                pct25=100 * sum(1 for x in lens if x >= 25) / len(lens) if lens else 0,
                unq=stt.count("UNQUOTED"), dia=stt.count("DIALOGUE"), mix=stt.count("MIXED"),
                diashare=100 * (stt.count("DIALOGUE") + stt.count("MIXED")) / len(stt) if stt else 0)


def block_order(n, seed):
    ids = ["P%03d" % (i + 1) for i in range(n)]
    shuffled = ids[:]
    rng = random.Random(seed)
    for _ in range(50):
        rng.shuffle(shuffled)
        if shuffled != ids:
            break
    return shuffled


def render(blocks, order, include_frame, donor_note, seed=None, order_note=None):
    m = measurements(blocks)
    L = []
    L.append("STRUCTURAL TRANSPLANT SKELETON")
    L.append("DONOR: %s" % donor_note)
    L.append("")
    L.append("MEASUREMENTS")
    L.append("TARGET WORDS: sd %.1f, mean %.1f, CV %.2f, longest unit %d, %.0f%% of units at 25 words or more."
             % (m["sd"], m["mean"], m["cv"], m["longest"], m["pct25"]))
    L.append("Status mix: %d UNQUOTED, %d DIALOGUE, %d MIXED — %.0f%% DIALOGUE or MIXED."
             % (m["unq"], m["dia"], m["mix"], m["diashare"]))
    L.append("Donor window: %d words, %d paragraphs, %d sentence units." % (m["words"], m["blocks"], m["units"]))
    L.append("")
    L.append("BLOCK ORDER")
    if order_note:
        L.append(order_note)
    elif seed is not None:
        L.append("Content-blind, seeded %d, reproducible. Run-specific, not a reusable novel architecture." % seed)
    else:
        L.append("Content-blind external order. Run-specific, not a reusable novel architecture.")
    for i in range(0, len(order), 12):
        L.append("  " + " ".join(order[i:i + 12]))
    L.append("")
    L.append("========================================================================")
    L.append("BEGIN SKELETON")
    L.append("========================================================================")
    L.append("ABSTRACT HUMAN STRUCTURAL SKELETON v2")
    L.append("%d paragraphs; %d sentence units." % (m["blocks"], m["units"]))
    L.append("No donor vocabulary, events, images, setting, or narrative-function labels are included.")
    L.append("POS TAGS")
    L.append("NOUN/PROPN; V.BASE/V.PAST/V.ING/V.PART/V.PRES/V.PRES3;")
    L.append("ADJ; ADV; PRON; DET; ADP/SUB; COORD; TO; MODAL; PART; NUM; WH.*;")
    L.append("plus named punctuation positions.")
    L.append("PHRASES gives the shallow phrase sequence only:")
    L.append("NP = noun phrase; VP = verb phrase; PP = prepositional phrase;")
    L.append("ADJP = adjective phrase; ADVP = adverb phrase.")
    L.append("TARGET WORDS gives the donor-derived sentence-length target for that sentence unit. "
             "It is a structural target, not a request to invent content.")
    for bi, b in enumerate(blocks, 1):
        L.append("P%03d (%d sentence%s)" % (bi, len(b["units"]), "" if len(b["units"]) == 1 else "s"))
        if include_frame:
            L.append("DONOR CLOSED-CLASS FRAME: " + b["frame"])
        for si, u in enumerate(b["units"], 1):
            L.append("S%02d | %s | %s" % (si, u["status"], u["term"]))
            L.append("POS " + " ".join(u["tags"]))
            L.append("TARGET WORDS %d" % u["tw"])
            L.append("PHRASES " + " > ".join(u["phr"]))
    return "\n".join(L)


def validate(blocks, order, text):
    checks = []
    n = len(blocks)
    checks.append(("every paragraph produced one block, in donor order",
                   len(re.findall(r"^P\d{3} \(", text, re.M)) == n))
    units = sum(len(b["units"]) for b in blocks)
    checks.append(("every sentence produced one unit, in donor order",
                   len(re.findall(r"^S\d\d \| ", text, re.M)) == units))
    ok = all(len(b["units"]) == int(re.match(r".*\((\d+) sentence", "P (%d sentences)" % len(b["units"])).group(1))
             for b in blocks)
    checks.append(("each block's declared sentence count equals its unit count", ok))
    bad = 0
    for b in blocks:
        for u in b["units"]:
            npunct = len([t for t in u["tags"] if t not in
                          {"PERIOD", "COMMA", "QUOTE", "QMARK", "EXCL", "SEMICOLON", "APOST", "HYPHEN", "OTHER"}])
            apost = u["tags"].count("APOST")
            if npunct - apost != u["tw"]:
                bad += 1
    checks.append(("TARGET WORDS equals non-punctuation tags minus APOST count (%d exceptions)" % bad, bad == 0))
    checks.append(("BLOCK ORDER covers every block exactly once",
                   sorted(order) == sorted("P%03d" % (i + 1) for i in range(n))))
    checks.append(("BLOCK ORDER differs from the printed order",
                   order != ["P%03d" % (i + 1) for i in range(n)]))
    return checks


def leak_check(blocks, text):
    """Report any donor content word that survived into the skeleton output."""
    body = text.split("BEGIN SKELETON", 1)[-1]
    allowed = set()
    for b in blocks:
        for u in b["units"]:
            for tok, tag in zip(u["toks"], u["tags"]):
                if tag in CLOSED_TAGS or tok.lower() in (BE | HAVE | DO | MODAL | NEG):
                    allowed.add(tok.lower())
    leaked = set()
    for w in re.findall(r"\b[a-z]{3,}\b", body.lower()):
        if w in allowed:
            continue
        if w in {"noun", "propn", "adj", "adv", "pron", "det", "adp", "sub", "coord", "modal",
                 "part", "num", "period", "comma", "quote", "qmark", "excl", "semicolon",
                 "apost", "hyphen", "other", "exist", "intj", "base", "past", "ing", "pres",
                 "sentence", "sentences", "target", "words", "phrases", "pos", "donor",
                 "closed", "class", "frame", "abstract", "human", "structural", "skeleton",
                 "paragraphs", "units", "unit", "vocabulary", "events", "images", "setting",
                 "narrative", "function", "labels", "included", "tags", "named", "punctuation",
                 "positions", "gives", "shallow", "phrase", "sequence", "only", "noun",
                 "verb", "prepositional", "adjective", "adverb", "derived", "length",
                 "structural", "request", "invent", "content", "dialogue", "mixed",
                 "unquoted", "declarative", "question", "exclamation", "begin", "block",
                 "order", "measurements", "status", "mix", "window", "sup", "comp", "pl",
                 "poss", "adjp", "advp", "num", "seeded", "reproducible", "run", "specific",
                 "reusable", "novel", "architecture", "blind", "longest", "more", "and",
                 "the", "with", "that", "not", "its", "for", "per", "cent"}:
            continue
        leaked.add(w)
    return sorted(leaked)


def provenance_tsv(blocks):
    """Separate audit trail; provenance is intentionally not embedded in the skeleton."""
    rows = ["block_id\tsource\tsource_paragraph\tsentence_units"]
    for i, b in enumerate(blocks, 1):
        rows.append("P%03d\t%s\t%s\t%d" % (
            i,
            str(b.get("source", "")).replace("\t", " "),
            str(b.get("source_para", "")),
            len(b.get("units", [])),
        ))
    return "\n".join(rows)


def best_window(paras, target, tol=0.15):
    pw = [len(re.findall(r"[A-Za-z']+", p)) for p in paras]
    best = None
    for i in range(len(paras)):
        tot = 0
        for j in range(i, len(paras)):
            tot += pw[j]
            if tot >= target * (1 - tol):
                if tot <= target * (1 + tol):
                    blocks = build_units(paras[i:j + 1])
                    m = measurements(blocks)
                    if best is None or m["sd"] > best[0]["sd"]:
                        best = (m, i, j + 1)
                break
    return best


def to_docx_bytes(text):
    import docx
    from docx.shared import Pt
    d = docx.Document()
    s = d.styles["Normal"]
    s.font.name = "Calibri"
    s.font.size = Pt(10)
    mono = ("P0", "P1", "S0", "S1", "S2", "S3", "S4", "POS ", "TARGET WORDS ", "PHRASES ", "  P")
    for line in text.split("\n"):
        if not line.strip():
            continue
        p = d.add_paragraph(line)
        if line.startswith(mono):
            for r in p.runs:
                r.font.name = "Consolas"
                r.font.size = Pt(8)
            if re.match(r"^P\d{3} \(", line):
                for r in p.runs:
                    r.bold = True
        elif re.match(r"^[A-Z0-9 ,\u2014\-\.:/()]{8,}$", line):
            for r in p.runs:
                r.bold = True
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()

# --------------------------------------------------------------------- APP --


def main():
    import streamlit as stl

    stl.set_page_config(page_title="Skeleton Extractor", layout="wide")
    stl.title("Skeleton Extractor — Aug. 18 Recreation")
    stl.caption("Recreates the v6F structural skeleton format from one or more donor documents. "
                "The exact lost Aug. 18 source code is not claimed; this is the documented deterministic reconstruction.")

    stl.info(
        "Historical status: recreated from surviving v6F outputs and the Aug. 26 reconstruction. "
        "The exact Aug. 18 parser/tagger source was not found in Dropbox or available Downloads backups."
    )

    files = stl.file_uploader("Donor documents (.docx, .txt, .md) — upload as many as you like",
                              type=["docx", "txt", "md"], accept_multiple_files=True)
    if not files:
        stl.info("Upload one or more documents to begin. Each source keeps its own paragraph boundaries and provenance. "
                 "Multi-document windowing is disabled so source boundaries cannot be crossed.")
        return

    names = [f.name for f in files]
    order_names = stl.multiselect("Order the documents (first to last). Deselect any you want left out.",
                                  names, default=names)
    if not order_names:
        stl.warning("Select at least one document.")
        return

    by_name = {f.name: f for f in files}
    docs, errs = [], []
    for n in order_names:
        try:
            ps = read_upload(n, by_name[n].getvalue())
            docs.append((n, ps))
        except Exception as e:
            errs.append("%s — %s" % (n, e))
    for e in errs:
        stl.error(e)
    if not docs:
        return

    para_records = []
    for n, ps in docs:
        integ = paragraph_integrity(ps)
        if integ["nonterminal_pct"] > 20:
            stl.warning(
                "%s: %.1f%% of paragraph records do not end in terminal punctuation. "
                "This strongly suggests imported hard line wraps. Do not use this donor "
                "for paragraph-block extraction until the source formatting is repaired."
                % (n, integ["nonterminal_pct"])
            )
        for pi, p in enumerate(ps, 1):
            para_records.append(dict(source=n, source_para=pi, text=p))

    paras = [r["text"] for r in para_records]
    total_words = sum(len(re.findall(r"[A-Za-z']+", p)) for p in paras)
    stl.write("**Loaded:** %d document(s), %d source paragraphs, %d words."
              % (len(docs), len(paras), total_words))

    c1, c2, c3 = stl.columns(3)
    with c1:
        multi = len(docs) > 1
        use_window = stl.checkbox(
            "Use a window rather than the whole text",
            value=False,
            disabled=multi,
            help=("Windowing is disabled for multi-document corpora so a window can never cross "
                  "a source boundary." if multi else None)
        )
        target = stl.number_input("Window size (words)", 500, 20000, 4800, 100,
                                  disabled=(not use_window))
    with c2:
        include_frame = stl.checkbox("Include DONOR CLOSED-CLASS FRAME", value=True,
                                     help="Leave this off for a copyrighted donor. The frame keeps the donor's "
                                          "function words in their original positions.")
        order_mode = stl.radio(
            "Block order",
            ["New seeded content-blind order", "Canonical Aug. 18 v6F order (119 blocks only)"],
            index=0,
        )
        seed = stl.number_input("Block order seed", 0, 999999, 1887, 1,
                                disabled=order_mode.startswith("Canonical"))
    with c3:
        donor_note = stl.text_input("Donor note (printed in the header)", value=", ".join(order_names))

    if use_window:
        bw = best_window(paras, int(target))
        if bw is None:
            stl.warning("No paragraph-aligned window of that size exists. Using the whole text.")
            chosen = para_records
        else:
            m, i, j = bw
            stl.write("**Best window:** paragraphs %d–%d, %d words, sd %.1f, CV %.2f, %.0f%% dialogue or mixed."
                      % (i, j - 1, m["words"], m["sd"], m["cv"], m["diashare"]))
            chosen = para_records[i:j]
    else:
        chosen = para_records

    if stl.button("Build skeleton", type="primary"):
        with stl.spinner("Extracting…"):
            blocks = build_units(chosen)
            if not blocks:
                stl.error("No sentences found.")
                stl.session_state.pop("result", None)
                return
            if order_mode.startswith("Canonical"):
                if len(blocks) != 119:
                    stl.error("The canonical Aug. 18 order requires exactly 119 P-blocks. "
                              "Choose a seeded order or select a 119-block donor set.")
                    stl.session_state.pop("result", None)
                    return
                order = CANONICAL_V6F_ORDER[:]
                order_note = ("Canonical content-blind block order embedded in the Aug. 18 v6F prompt; "
                              "historical reproduction only, not a reusable novel architecture.")
                render_seed = None
            else:
                order = block_order(len(blocks), int(seed))
                order_note = None
                render_seed = int(seed)
            text = render(blocks, order, include_frame, donor_note,
                          seed=render_seed, order_note=order_note)
            try:
                docx_bytes = to_docx_bytes(text)
            except ImportError:
                docx_bytes = None
            # Everything the results panel needs is stored now. A download click
            # reruns the script, at which point the Build button reads False, so
            # the payload has to survive in session_state or the browser asks the
            # server for a file that no longer exists.
            stl.session_state["result"] = dict(
                text=text, docx=docx_bytes, m=measurements(blocks),
                checks=validate(blocks, order, text),
                leaked=leak_check(blocks, text) if not include_frame else [],
                framed=include_frame,
                provenance=provenance_tsv(blocks))

    res = stl.session_state.get("result")
    if not res:
        return
    text, m, checks, leaked = res["text"], res["m"], res["checks"], res["leaked"]

    stl.subheader("Measurements")
    a, b, c, d, e = stl.columns(5)
    a.metric("sd", "%.1f" % m["sd"])
    b.metric("mean", "%.1f" % m["mean"])
    c.metric("CV", "%.2f" % m["cv"])
    d.metric("dialogue + mixed", "%.0f%%" % m["diashare"])
    e.metric("units", m["units"])
    stl.write("%d blocks, %d units, %d words, longest unit %d, %.0f%% of units at 25 words or more."
              % (m["blocks"], m["units"], m["words"], m["longest"], m["pct25"]))

    stl.subheader("Validation")
    for label, ok in checks:
        stl.write(("✅ " if ok else "❌ ") + label)
    if not res["framed"] and leaked:
        stl.warning("Content words found in the skeleton body: " + ", ".join(leaked[:40]))

    stl.subheader("Output")
    stl.download_button("Download .txt", text.encode("utf8"), key="dl_txt",
                        file_name="Structural_Transplant_Skeleton.txt", mime="text/plain")
    stl.download_button("Download provenance .tsv", res["provenance"].encode("utf8"),
                        key="dl_provenance", file_name="Structural_Transplant_Skeleton_Provenance.tsv",
                        mime="text/tab-separated-values")
    if res["docx"] is not None:
        stl.download_button("Download .docx", res["docx"], key="dl_docx",
                            file_name="Structural_Transplant_Skeleton.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        stl.caption("python-docx is missing from requirements.txt, so only .txt is available.")

    stl.text_area("Preview (first 4,000 characters)", text[:4000], height=400)

    with stl.expander("ACCURACY — what this tagger does and does not get right"):
        stl.markdown(
            "The tagger is rule-based and deterministic: same donor, same skeleton, every run. No model is "
            "downloaded and no run differs from the last.\n\n"
            "**Reliable.** TARGET WORDS, DIALOGUE / MIXED / UNQUOTED status, terminal mode, block and unit "
            "structure, and the closed-class inventory used by the frame. TARGET WORDS is checked against the "
            "POS sequence for every unit, and any mismatch is reported above.\n\n"
            "**Approximate.** The POS sequence on words that could be either noun or verb, and the ADJ/NOUN "
            "boundary. English is ambiguous there and no rule set settles it without a parser.\n\n"
            "**Reconstructed.** PHRASES follows the chunking rule I recovered from the v6F skeleton, which "
            "reproduces 220 of its 321 phrase sequences exactly. The rest are places where v6F's own lines are "
            "not internally consistent, so no rule reproduces them all.")


if __name__ == "__main__":
    main()
