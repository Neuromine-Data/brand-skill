import re

EM_DASH = "\u2014"
EN_DASH = "\u2013"
BANNED = [
    r"delv(?:e|es|ed|ing)",
    r"robust(?:ly|ness)?",
    r"leverag(?:e|es|ed|ing)",
    r"seamless(?:ly)?",
    r"utili[sz](?:e|es|ed|ing|ation)",
    r"cutting-edge",
    r"comprehensive(?:ly)?",
    r"empower(?:s|ed|ing|ment)?",
    r"crucial(?:ly)?",
    r"landscapes?",
    r"journeys?",
    r"unlock(?:s|ed|ing)?",
    r"harness(?:es|ed|ing)?",
    r"tapestr(?:y|ies)",
    r"navigat(?:e|es|ed|ing)",
    r"in today's fast-paced world",
    r"it is important to note",
    r"in conclusion",
    r"overall,",
    r"not only",
    r"may potentially",
    r"i hope this helps",
    r"great question",
]
BANNED_RE = re.compile(r"(?<![\w-])(?:" + "|".join(BANNED) + r")(?![\w-])", re.IGNORECASE)
DASH_RE = re.compile("[" + EM_DASH + EN_DASH + "]")
URL_RE = re.compile(r"(?:https?|ftp|mailto):[^\s<>()\"']+|www\.[^\s<>()\"']+", re.IGNORECASE)


def snippet(text, start, end, width=25):
    a = max(0, start - width)
    b = min(len(text), end + width)
    out = text[a:b].replace(EM_DASH, "<em dash>").replace(EN_DASH, "<en dash>")
    out = " ".join(out.split()).encode("ascii", "replace").decode("ascii")
    return ("..." if a else "") + out + ("..." if b < len(text) else "")


def find_banned(text):
    text = URL_RE.sub(" ", text or "")
    hits = []
    for m in DASH_RE.finditer(text):
        hits.append(("em dash" if m.group() == EM_DASH else "en dash", snippet(text, m.start(), m.end())))
    for m in BANNED_RE.finditer(text):
        hits.append((m.group().lower(), snippet(text, m.start(), m.end())))
    return hits
