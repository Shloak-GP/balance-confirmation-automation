"""Fixed rules that turn a reply's text into a suggestion. No LLM, no network.

Everything here is a suggestion: the owner confirms every item on screen.
"""
import math
import re
from difflib import SequenceMatcher

# Phrases that mean "our balance agrees with yours".
NO_CHANGE = [
    r"no changes?", r"no change in (the )?balance", r"we (hereby )?confirm", r"confirmed",
    r"agreed?", r"correct as per our (books|records)", r"in agreement", r"nil difference",
    r"balance is correct", r"tallies", r"matches", r"\bi (hereby )?confirm\b", r"confirm(ed)? the same", r"same as (per )?(your|our)",
]
# Words that attach a condition to a reply. Any of these blocks a pre-filled answer:
# in "confirmed except bill 4471 for Rs 12,400" the 12,400 is a bill, not their balance.
CAVEATS = [r"\bexcept\b", r"\bbut\b", r"subject to", r"\bpending\b", r"\bhowever\b", r"\bexcluding\b"]
# Words that say the customer disagrees. A differing amount is then expected; no automatic CLEAR.
DISAGREE = [r"not agree", r"disagree", r"don't agree", r"\bdiffer", r"not in agreement", r"not confirm",
            r"cannot confirm", r"can't confirm", r"\bnot correct\b", r"\bincorrect\b"]

AMOUNT_RE = re.compile(
    r"(?:(?:rs\.?|inr|₹)\s*)?"
    # Indian (1,67,75,671) or Western (3,452,112) grouping, plain 4+ digits, or 12.50. '/-' after is fine.
    r"(?<![\d/,])(?<!\d\.)(\d{1,3}(?:,\d{2})*,\d{3}(?:\.\d{1,2})?|\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d{4,}(?:\.\d{1,2})?|\d{1,3}\.\d{2})(?!\d|/\d|,\d)"
    r"\s*(?:/-)?\s*(dr|cr|credited|debited)?",
    re.I,
)
PERIOD_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*(?:to|-|till|upto|up to)\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", re.I)
FROM_RE = re.compile(r"\n\s*From:\s", re.I)


def latest_reply(text):
    """Mails arrive forwarded by a regional manager with the whole thread under the customer's reply.
    Only the customer's newest message counts: the part after the first 'From:' header up to the next one
    (older messages below it are ours or earlier replies, with their own amounts)."""
    parts = FROM_RE.split(text or "")
    if len(parts) >= 2 and not find_amounts(parts[0]):
        return parts[1]
    return parts[0]
REF_RE = re.compile(r"(?:\b(?:no|nos|number|ref|invoice|bill)\b\.?|#)\s*[:.-]?\s*\d[\w/-]*", re.I)
CODE_RE = re.compile(r"\b(?!rs\.?\d)(?!inr\d)[a-z]+[-/]?\d[\w/-]*", re.I)
DATE_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b")


def norm(s):
    return re.sub(r"\s+", " ", (s or "").lower()).strip()


def name_key(s):
    s = norm(s)
    s = re.sub(r"\b(m/s\.?|messrs\.?|pvt|private|ltd|limited|llp|co|company|and|&)\b\.?", " ", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def parse_amount(txt):
    """'4,12,350.00' -> 412350.0. Returns None if it is not a number."""
    try:
        return round(float(str(txt).replace(",", "").replace("₹", "").replace("rs", "").strip()), 2)
    except ValueError:
        return None


ATT_MARK = "\n\x00ATTACHMENT\x00\n"   # ingest puts this between the mail body and text read from attachments
SIGN_OFF = re.compile(r"(?im)^[ \t>]*(thanks\s*(?:&|and)?\s*regards|best regards|kind regards|with regards|warm regards|regards|"
                      r"thanking you|thanks|thank you|yours (?:faithfully|sincerely))\b")
AMT_WORDS = re.compile(r"(?i)(balance|amount|outstanding|due|payable|receivable|total|net|rs\b|inr|₹|as per|debit|credit)")
PHONE_RES = [re.compile(p) for p in (
    r"<tel:[^>]*>", r"(?i)\b(?:t|f|m|tel|fax|mob|mobile|ph|phone|cell|contact(?: no\.?)?|whatsapp)\b[\s.:()a-z-]*\+?[\d\s()-]{8,}",
    r"\+\s*91[\s-]?[\d\s-]{9,}", r"\b0\d{2,4}[-\s]\d{5,8}\b", r"\b\d{10,13}\b", r"\b\d{3,5}\s\d{5,6}\b")]


OUR_LETTER = re.compile(r"(?i)confirmation of balance as on|audit of our acc|balance furnished below")


def our_letter_attached(text):
    return any(OUR_LETTER.search(a) for a in (text or "").split(ATT_MARK)[1:])


QUOTE_START = re.compile(r"(?im)^[ \t>]*(on .{3,160}wrote:|-{2,}\s*original message|original message:|_{8,})")


def prepare(text, parts=False):
    """What may be searched for the customer's balance: the customer's newest message up to the sign-off
    (signatures hold phone numbers and PIN codes), plus any text read from attachments."""
    body, *atts = (text or "").split(ATT_MARK)
    atts = [a for a in atts if not OUR_LETTER.search(a)]   # our own printed statement: every amount in it is OURS
    body = latest_reply(body)
    q = QUOTE_START.search(body)           # "On Mon ... wrote:" - what follows is the older mail being answered
    if q:
        body = body[:q.start()]
    m = SIGN_OFF.search(body)
    if m:
        body = body[:m.start()]
    if parts:
        return body, "\n".join(atts)
    return "\n".join([body] + atts)


def find_amounts(text):
    """All money-looking numbers in the text, ignoring dates and years."""
    clean = re.sub(r"<?https?://\S+>?|\S+@\S+", " ", (text or "").replace("\x00ATTACHMENT\x00", " "))  # links, e-mail addresses
    clean = re.sub(r"(?i)\bp\.?\s?o\.?\s?box\s*:?\s*\d+", " ", clean)   # postal box numbers
    clean = re.sub(r"(?i)\b[1-9]\d{5}\b(?!\s*(?:dr|cr|credited|debited)\b)(?=\s+[a-z]{3,})", " ", clean)  # PIN code before a city
    for pr in PHONE_RES:                                      # phone / fax numbers
        clean = pr.sub(" ", clean)
    clean = DATE_RE.sub(" ", clean)
    # Not money: reference numbers ("bill no. 4471", "Invoice No: 12345", "#889")
    clean = REF_RE.sub(" ", clean)
    # Not money: codes with letters in them ("CN10002", "GSTIN 27AAACG...") - but keep "Rs4000"
    clean = CODE_RE.sub(" ", clean)
    out = []
    for m in AMOUNT_RE.finditer(clean):
        v = parse_amount(m.group(1))
        if v is None:
            continue
        if 1900 <= v <= 2100 and "," not in m.group(1) and "." not in m.group(1):
            continue  # a year
        g = m.group(1)
        strong = bool(re.match(r"(?i)(rs|inr|₹)", m.group(0).strip())) or "," in g or "." in g or "/-" in m.group(0)
        if not strong and not AMT_WORDS.search(clean[max(0, m.start() - 60):m.start()]):
            continue  # bare digits (PIN codes, pages, ...) only count right after words like balance / amount / Rs
        out.append({"value": v, "side": (m.group(2) or "").upper(), "pos": m.start(), "strong": strong})
    return out


def has_any(text, patterns):
    t = norm(text)
    return [p for p in patterns if re.search(p, t)]


def _newest_text(text):
    """Forwarder's note + the customer's newest message. Older quoted mails (which may contain tables of other
    customers' numbers) are left out. Plain PDFs / direct mails have no 'From:' headers and are used whole."""
    parts = FROM_RE.split(text or "")
    return "\n".join(parts[:2]) if len(parts) > 1 else (text or "")


def find_customer(text, filename, customers, mail_from="", own_domains=("gpglobal.com",)):
    """customers: dict No. -> {'name': ...}. Returns (no, how, name_ok, candidates)."""
    newest = _newest_text(text)
    hay = [("file name", filename or ""), ("mail/PDF text", newest)]
    codes = sorted(customers, key=len, reverse=True)
    for where, h in hay:
        up = h.upper()
        for code in codes:
            if re.search(r"(?<![A-Z0-9])" + re.escape(code.upper()) + r"(?![A-Z0-9])", up):
                nk = name_key(customers[code]["name"])
                name_ok = bool(nk) and (nk in name_key(filename) or nk in name_key(text)
                                        or _best_ratio(nk, name_key(filename) + " " + name_key(text)) >= 0.85)
                return code, f"customer No. found in {where}", name_ok, []
    # No customer No. anywhere: sender e-mail first, then the name.
    return match_by_email_or_name(text, filename, customers, mail_from, own_domains)


FREE_MAIL = {"gmail.com", "yahoo.com", "yahoo.co.in", "yahoo.in", "rediffmail.com", "hotmail.com", "outlook.com",
             "live.com", "ymail.com", "aol.com", "protonmail.com", "icloud.com", "msn.com"}
EMAIL_RE = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
GENERIC = {"pvt", "private", "ltd", "limited", "llp", "co", "company", "and", "the", "of", "m", "s", "ms", "messrs"}


def customer_emails(text, mail_from, own_domains):
    """E-mail addresses on the customer's side of the newest message (not ours)."""
    own = {d.lower() for d in own_domains}
    parts = FROM_RE.split(text or "")
    head = []
    if mail_from:
        head += EMAIL_RE.findall(mail_from)
    if len(parts) >= 2:                      # forwarded: first header after the forwarder's note = the customer
        head += EMAIL_RE.findall("\n".join(parts[1].splitlines()[:6]))
    elif parts:
        head += EMAIL_RE.findall("\n".join(parts[0].splitlines()[:8]))
    out = []
    for a in head:
        a = a.lower().strip(".")
        if a.split("@")[-1] not in own and a not in out:
            out.append(a)
    return out


def _tokens(s):
    t = norm(s)
    t = re.sub(r"\b([a-z])\.(?=[a-z]\b)", r"\1", t)   # initials: "k.p. enterprises" -> "kp. enterprises"
    t = re.sub(r"\b([a-z])\.(?=[a-z]\b)", r"\1", t)
    return [x for x in re.sub(r"[^a-z0-9 ]+", " ", t).split() if x not in GENERIC and len(x) > 1]


def match_by_email_or_name(text, filename, customers, mail_from="", own_domains=("gpglobal.com",)):
    parts = FROM_RE.split(text or "")
    URLS = re.compile(r"<?https?://\S+>?|\S+@\S+")
    ev_text = URLS.sub(" ", " ".join([filename or "", (text or "")[:250], parts[0][:400],
                                      parts[1][:1200] if len(parts) > 1 else ""]))
    ev_list = _tokens(ev_text)
    ev = set(ev_list)
    # document frequency of every name token over the customer list (rare words identify a customer)
    df = {}
    toks = {}
    for no, c in customers.items():
        t = set(_tokens(c["name"]))
        toks[no] = t
        for w in t:
            df[w] = df.get(w, 0) + 1
    n = max(len(customers), 1)
    w = lambda tok: math.log((n + 1) / (df.get(tok, 0) + 1)) + 0.1

    def phrase_in_evidence(no):
        nt = _tokens(re.sub(r"\(.*?\)", " ", customers[no]["name"]))
        if len(nt) < 2:
            return False
        return any(ev_list[i:i + len(nt)] == nt for i in range(len(ev_list) - len(nt) + 1))

    def name_score(no):
        t = toks[no]
        tot = sum(w(x) for x in t)
        base = sum(w(x) for x in t if x in ev) / tot if tot else 0.0
        return base + (0.5 if phrase_in_evidence(no) else 0.0)

    # 1) sender address
    addrs = customer_emails(text, mail_from, own_domains)
    by_addr, by_domain = {}, {}
    for no, c in customers.items():
        for a in EMAIL_RE.findall(str(c.get("email") or "").lower()):
            by_addr.setdefault(a, set()).add(no)
            d = a.split("@")[-1]
            if d not in FREE_MAIL:
                by_domain.setdefault(d, set()).add(no)
    exact = set()
    for a in addrs:
        exact |= by_addr.get(a, set())
    if exact:
        ranked = sorted(exact, key=lambda x: -name_score(x))
        if len(ranked) == 1 or name_score(ranked[0]) - name_score(ranked[1]) >= 0.15:
            return ranked[0], "matched by sender e-mail " + addrs[0], True, ranked[:5]
        return None, "sender e-mail belongs to several customers", False, ranked[:5]
    dom = set()
    for a in addrs:
        dom |= by_domain.get(a.split("@")[-1], set())
    pool = dom or set(customers)
    # 2) name
    scored = sorted(((name_score(no), no) for no in pool), reverse=True)
    scored = [x for x in scored if x[0] >= 0.6]
    cands = [no for _, no in scored[:5]]
    base = lambda no: " ".join(_tokens(re.sub(r"\(.*?\)", " ", customers[no]["name"])))
    siblings = [no for sc, no in scored[1:] if base(no) == base(scored[0][1])] if scored else []
    if siblings:  # same company with several accounts (e.g. "(IPOL)" and "(RPO)"): the name cannot tell them apart
        return None, "this customer has several accounts", False, cands
    if scored and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.15):
        how = "matched by name" + (" and e-mail domain" if dom else "")
        return scored[0][1], how, scored[0][0] >= 0.8 or bool(dom), cands
    return None, "not found" if not cands else "several customers look alike", False, cands


def pick_by_amount(text, cands, customers):
    """Among look-alike accounts, the one whose ERP balance equals an amount the customer states.
    Only if exactly one account matches; otherwise None (he picks)."""
    amts = {abs(a["value"]) for a in find_amounts(prepare(text))}
    hit = [no for no in cands if customers[no].get("erp") is not None and abs(customers[no]["erp"]) in amts]
    return hit[0] if len(hit) == 1 else None


def _best_ratio(needle, hay):
    """Best similarity of needle against any same-length word window of hay."""
    n = len(needle.split())
    words = hay.split()
    best = 0.0
    for i in range(max(1, len(words) - n + 1)):
        best = max(best, SequenceMatcher(None, needle, " ".join(words[i:i + n])).ratio())
    return best


def suggest(text, erp_balance, name=""):
    """Suggest the customer's balance and status.

    Returns dict(theirs, status, reason, needs_typing, flags).
    status is CLEAR / DISPUTE / None (owner decides).
    """
    flags = []
    letter = our_letter_attached(text)
    if letter:
        flags.append("The attachment is our own balance-confirmation letter. Its printed amounts are ours: look at it for the customer's signature or a handwritten balance.")
    body, atts_text = prepare(text, parts=True)
    text = body + "\n" + atts_text
    if PERIOD_RE.search(body or ""):
        flags.append("The reply gives a period (from-to dates). It may be a period total, not the closing balance.")
    # Words like "except" / "differences" count only in what the customer WROTE; attachments (our letter, their
    # ledger) carry boilerplate such as "reconcile the differences if any".
    phrases = body if norm(body) else text
    caveats = has_any(phrases, CAVEATS)
    disagree = has_any(phrases, DISAGREE)
    nochange = [] if disagree else has_any(phrases, NO_CHANGE)
    amounts = find_amounts(text)
    if name and len({a["value"] for a in amounts}) > 1:
        # One reply may list several customers ("K.P. Enterprises Rs 44,26,391 / Sunrise Enterprises Rs 1,98,36,248"):
        # take the amounts on the lines that carry THIS customer's name.
        nt = [t for t in _tokens(re.sub(r"\(.*?\)", " ", name)) if len(t) > 1]
        hit = []
        for need in (len(nt), max(1, (len(nt) + 1) // 2)):   # every name word on the line, else at least half
            for ln in text.splitlines():
                lt = set(_tokens(ln))
                if nt and sum(t in lt for t in nt) >= need:
                    hit += find_amounts(ln)
            if hit:
                break
        others = {a["value"] for a in amounts if erp_balance is None or abs(a["value"]) != abs(erp_balance)}
        if hit:
            amounts = hit
        elif len(others) > 1:
            flags.append("The reply has several different amounts and none is next to this customer's name.")
            return dict(theirs=None, status=None, needs_typing=True, flags=flags,
                        reason="Several amounts in the reply. Read it and type the customer's balance.")
    if any(a["strong"] for a in amounts):
        amounts = [a for a in amounts if a["strong"]]  # bare digits are only used when nothing looks like money
    if caveats:
        flags.append("Reply contains: " + ", ".join(sorted({c.strip('\\b') for c in caveats})))

    if not norm(text):
        return dict(theirs=None, status=None, needs_typing=True, flags=flags,
                    reason="No readable text (scanned or handwritten). Please type the customer's balance.")

    erp = None if erp_balance is None else round(float(erp_balance), 2)
    if caveats or any("period" in f for f in flags):
        # e.g. "confirmed except bill 4471 for Rs 12,400": the 12,400 is a bill, not their balance.
        return dict(theirs=None, status=None, needs_typing=True, flags=flags,
                    reason="The reply has conditions. Read it and type the customer's balance.")
    others = [a for a in amounts if erp is None or abs(a["value"]) != abs(erp)]
    same = [a for a in amounts if erp is not None and abs(a["value"]) == abs(erp)]

    if others:
        # The customer's figure is usually the last amount that differs from ours.
        pick = others[-1]["value"]
        return dict(theirs=pick, status=status_for(erp, pick), needs_typing=False, flags=flags,
                    reason="Found an amount different from our balance. Check it against the form.")
    if nochange and not caveats:
        return dict(theirs=erp, status="CLEAR", needs_typing=False, flags=flags,
                    reason="Customer says the balance is confirmed / no changes.")
    if same and not caveats and not disagree:
        return dict(theirs=erp, status="CLEAR", needs_typing=False, flags=flags,
                    reason="The only amount found equals our balance.")
    return dict(theirs=None, status=None, needs_typing=True, flags=flags,
                reason="Could not tell the customer's balance from the text. Please read the form and type it.")


def status_for(erp, theirs):
    """The owner's rule: any difference at all is a dispute."""
    if erp is None or theirs is None:
        return None
    return "CLEAR" if round(float(erp), 2) == round(float(theirs), 2) else "DISPUTE"


def difference(erp, theirs):
    if erp is None or theirs is None:
        return None
    return round(float(erp) - float(theirs), 2)
