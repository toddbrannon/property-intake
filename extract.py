"""
Property PDF -> structured record.   Part of the property-intake pipeline.

    python extract.py path/to/file.pdf [more.pdf ...]

For each PDF this script:
  1. loads the field list from fields.csv (edit that file to change what is extracted),
  2. sends the PDF to the Claude API with the field list and strict rules,
  3. requires an explicit null (never a guess) for anything not in the document,
  4. requires a page citation, a confidence level and the source text for every value,
  5. validates types, checks price/income/yield agree,
  6. writes <name>.json next to the output folder, appends one row to master.xlsx,
     and appends any flagged fields to review.csv.

Settings come from the .env file in this folder (see .env.example).
"""
import argparse, base64, csv, json, os, re, sys, datetime, time
from pathlib import Path

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------- settings
def load_env():
    env = HERE / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
load_env()

MODEL   = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")
# List price per million tokens for the model above (Sonnet 5.5), used only to display cost per document.
PRICE_IN, PRICE_OUT = float(os.environ.get("PRICE_IN", "2")), float(os.environ.get("PRICE_OUT", "10"))
OUT_DIR = Path(os.environ.get("OUTPUT_DIR", HERE / "output"))
MASTER  = OUT_DIR / "master.xlsx"
REVIEW  = OUT_DIR / "review.csv"
LOG     = OUT_DIR / "log.txt"

def log(msg):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"{stamp}  {msg}\n")
    print(msg)

# ---------------------------------------------------------------- fields.csv
def load_fields():
    fields = {}
    with (HERE / "fields.csv").open(encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            key = row["key"].strip()
            if not key:
                continue
            t = row["type"].strip().lower()
            if t not in TYPE_CHECK:
                raise SystemExit(f"fields.csv: field '{key}' has unknown type '{t}'. Allowed: {', '.join(TYPE_CHECK)}")
            fields[key] = {
                "ja": row["label_ja"].strip(), "en": row["label_en"].strip(), "type": t,
                "required": row["required"].strip().lower() in ("yes", "y", "true", "1"),
                "options": [o.strip() for o in row.get("options", "").split("|") if o.strip()],
            }
    if not fields:
        raise SystemExit("fields.csv has no fields.")
    return fields

TYPE_CHECK = {
    "string":  lambda v: isinstance(v, str),
    "integer": lambda v: isinstance(v, int) and not isinstance(v, bool),
    "number":  lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
    "date":    lambda v: isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", v) is not None,
    "enum":    lambda v: isinstance(v, str),
    "list":    lambda v: isinstance(v, list),
}

# ---------------------------------------------------------------- the prompt
PROMPT = """You are extracting data from a Japanese real-estate property document (物件概要書 or similar).
Return ONLY a JSON object with exactly these keys: {keys}
plus one extra key "_summary_en".

For EVERY field key return an object: {{"value": <value or null>, "page": <int or null>, "confidence": "high"|"medium"|"low", "source": "<the exact Japanese text you read, at most 60 characters>", "note": "<see rules>"}}

"_summary_en": 2-4 plain English sentences summarizing the property for a colleague who does not read Japanese: what it is, where, price, size, income or yield if stated, and the main items a buyer should check. Use only facts stated in the document.

Field guide (key: Japanese label, type):
{guide}

Rules. Follow them exactly:
- If the document does not state a value, set value to null. NEVER infer, estimate or compute a missing value.
- Prices: an integer number of yen (6億2,000万円 → 620000000; 3,348万円 → 33480000). "応相談" or any non-numeric price → null, and put the wording in price_note if that field exists.
- Areas: square metres as a number. If only 坪 is given, convert (1坪 = 3.30579㎡) and set confidence "medium".
- Dates: ISO "YYYY-MM" or "YYYY-MM-DD". Convert Japanese eras (平成30年 → 2018, 令和8年 → 2026).
- Percentages: numbers without the % sign.
- enum fields: the value must be one of the listed options.
- list fields: a list of short English phrases, or [].
- Use "low" confidence whenever the text is ambiguous, partly obscured, or the value needed interpretation.
- Write "source" in the document's original Japanese. Do not translate it.
- "note": when confidence is not "high", or value is null, one short English sentence saying why (what was converted, what was ambiguous, or that the document does not state it). Otherwise "".
"""

def build_prompt(fields):
    guide = "\n".join(
        f"- {k}: {v['ja']}, {v['type']}" + (f", one of {'/'.join(v['options'])}" if v["options"] else "")
        for k, v in fields.items())
    return PROMPT.format(keys=json.dumps(list(fields), ensure_ascii=False), guide=guide)

# ---------------------------------------------------------------- where the answer comes from
# Three ways to get Claude's reading of a PDF, tried in this order:
#   1. demo_extractions/<name>.json   recorded answers for the three sample PDFs (no key needed)
#   2. <name>.pdf.claude.json beside the PDF   you pasted the PDF and the prompt into the Claude app
#                                              yourself and saved the reply there (no key needed)
#   3. the Claude API                 automatic; needs ANTHROPIC_API_KEY in .env
DEMO_DIR = HERE / "demo_extractions"

def get_raw(pdf_path: Path, fields) -> tuple:
    stored = DEMO_DIR / f"{pdf_path.stem}.json"
    if stored.exists():
        return json.loads(stored.read_text(encoding="utf-8")), "demo"
    manual = pdf_path.with_name(pdf_path.name + ".claude.json")
    if manual.exists():
        text = manual.read_text(encoding="utf-8").strip()
        text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.M).strip()
        return json.loads(text), "manual"
    if os.environ.get("ANTHROPIC_API_KEY"):
        return call_claude(pdf_path, fields), "api"
    raise RuntimeError(
        f"Cannot read '{pdf_path.name}': no ANTHROPIC_API_KEY in .env, not a sample document, and no "
        f"'{manual.name}' beside it. Either add an API key, or run  python extract.py --prompt , paste the "
        "prompt and the PDF into the Claude app, and save the reply as that file (see SOP Step 3b).")

def call_claude(pdf_path: Path, fields) -> dict:
    import anthropic
    # Keys created outside a workspace must name one; set ANTHROPIC_WORKSPACE_ID (wrkspc_...) in that case.
    ws = os.environ.get("ANTHROPIC_WORKSPACE_ID", "").strip()
    client = anthropic.Anthropic(default_headers={"anthropic-workspace-id": ws} if ws else None)
    pdf_b64 = base64.standard_b64encode(pdf_path.read_bytes()).decode()
    t0 = time.time()
    resp = client.messages.create(
        model=MODEL, max_tokens=4000,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
            {"type": "text", "text": build_prompt(fields)},
        ]}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    raw = json.loads(text)
    u = getattr(resp, "usage", None)
    tin, tout = getattr(u, "input_tokens", 0) or 0, getattr(u, "output_tokens", 0) or 0
    raw["_meta"] = {"model": getattr(resp, "model", MODEL), "input_tokens": tin, "output_tokens": tout,
                    "seconds": round(time.time() - t0, 1),
                    "cost_usd": round(tin / 1e6 * PRICE_IN + tout / 1e6 * PRICE_OUT, 4)}
    return raw

# ---------------------------------------------------------------- validate
# Plain-English explanations shown on the review page: who raised the item, and why.
WHY = {
    "required": ("Rule in the field list",
                 "This field is marked required in fields.csv. Claude returned it empty because the document does not state it, "
                 "so a person decides what to enter. The system never fills a required field by guessing."),
    "low":      ("Claude's own judgment",
                 "Claude marked its reading as low confidence, meaning the text was ambiguous or needed interpretation. "
                 "Anything Claude is unsure of is sent to a person instead of being accepted."),
    "type":     ("Validation check",
                 "Claude's answer was not the expected kind of value (for example text where a number belongs), so it was "
                 "discarded rather than written to the spreadsheet."),
    "enum":     ("Validation check",
                 "The value is not one of the allowed options in fields.csv, so a person picks the right category."),
    "yield":    ("Arithmetic cross-check",
                 "The script recomputed yield as annual income divided by price. It disagrees with the yield printed in the "
                 "document by more than 0.15 points, so either a figure or the broker's yield is wrong."),
}

def validate(raw: dict, pdf_name: str, fields) -> dict:
    record, review = {}, []
    def flag(key, kind, issue, cell):
        by, why = WHY[kind]
        review.append({"field": key, "label": fields[key]["ja"], "issue": issue, "source": cell.get("source", ""),
                       "page": cell.get("page"), "raised_by": by, "why": why, "claude_note": cell.get("note", "")})
    for key, spec in fields.items():
        cell = raw.get(key) or {}
        val, page, conf, src = cell.get("value"), cell.get("page"), cell.get("confidence", "low"), cell.get("source", "")
        kind = issue = None
        if val is None:
            if spec["required"]:
                kind, issue = "required", "required field missing from document"
        elif not TYPE_CHECK[spec["type"]](val):
            kind, issue = "type", f"type mismatch: expected {spec['type']}, got {type(val).__name__}"
            val = None
        elif spec["type"] == "enum" and val not in spec["options"]:
            kind, issue = "enum", f"'{val}' is not one of {spec['options']}"
        elif conf == "low":
            kind, issue = "low", "low confidence - confirm against source"
        record[key] = {"value": val, "page": page, "confidence": conf, "source": src, "note": cell.get("note", "")}
        if kind:
            flag(key, kind, issue, cell)
    # Cross-check: stated yield vs income / price (only if all three present)
    g = lambda k: record.get(k, {}).get("value")
    p, inc, y = g("price_jpy"), g("annual_income_jpy"), g("gross_yield_pct")
    if p and inc and y and "gross_yield_pct" in fields:
        implied = round(inc / p * 100, 2)
        if abs(implied - y) > 0.15:
            flag("gross_yield_pct", "yield", f"stated yield {y}% but income ÷ price = {implied}%",
                 {**record["gross_yield_pct"], "note": f"¥{inc:,} ÷ ¥{p:,} = {implied}%, document says {y}%."})
    confs = [f["confidence"] for f in record.values()]
    stats = {"high": confs.count("high"), "medium": confs.count("medium"), "low": confs.count("low"),
             "empty": sum(1 for f in record.values() if f["value"] is None)}
    return {"file": pdf_name, "processed": datetime.datetime.now().isoformat(timespec="seconds"),
            "fields": record, "review": review, "status": "needs_review" if review else "auto_accepted",
            "summary_en": raw.get("_summary_en", ""), "meta": raw.get("_meta", {}), "stats": stats}

# ---------------------------------------------------------------- outputs
def append_master(result: dict, fields):
    from openpyxl import Workbook, load_workbook
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if MASTER.exists():
        wb = load_workbook(MASTER); ws = wb.active
    else:
        wb = Workbook(); ws = wb.active; ws.title = "Properties"
        ws.append(["file", "processed", "status"] + [fields[k]["ja"] for k in fields])
        from openpyxl.styles import Font
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = "D2"
    row = [result["file"], result["processed"], result["status"]]
    for k in fields:
        v = result["fields"][k]["value"]
        row.append("; ".join(v) if isinstance(v, list) else v)
    ws.append(row)
    wb.save(MASTER)

REVIEW_COLS = ["file", "field", "label_ja", "issue", "source_text", "page", "raised_by", "why", "claude_note",
               "decision", "corrected_value"]

def append_review(result: dict):
    rows = []
    if REVIEW.exists():
        with REVIEW.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
    for it in result["review"]:
        rows.append({"file": result["file"], "field": it["field"], "label_ja": it["label"], "issue": it["issue"],
                     "source_text": it["source"], "page": it["page"] or "", "raised_by": it.get("raised_by", ""),
                     "why": it.get("why", ""), "claude_note": it.get("claude_note", ""), "decision": "", "corrected_value": ""})
    with REVIEW.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=REVIEW_COLS, extrasaction="ignore"); w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in REVIEW_COLS})

def process(pdf: Path, fields) -> dict:
    raw, how = get_raw(pdf, fields)
    result = validate(raw, pdf.name, fields)
    result["read_by"] = how
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / f"{pdf.stem}.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    append_master(result, fields)
    append_review(result)
    m = result.get("meta") or {}
    extra = f", {m.get('input_tokens',0)}+{m.get('output_tokens',0)} tokens, ${m.get('cost_usd',0)}, {m.get('seconds',0)}s" if how == "api" and m else ""
    log(f"{pdf.name}: {result['status']} ({len(result['review'])} items for review) [read by: {how}{extra}]")
    return result

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdfs", nargs="*")
    ap.add_argument("--prompt", action="store_true", help="print the extraction prompt (to paste into the Claude app) and exit")
    args = ap.parse_args()
    fields = load_fields()
    if args.prompt:
        print(build_prompt(fields)); return
    if not args.pdfs:
        ap.error("give at least one PDF, or --prompt")
    for pdf in map(Path, args.pdfs):
        if not pdf.exists():
            log(f"{pdf}: file not found"); continue
        process(pdf, fields)

if __name__ == "__main__":
    main()
