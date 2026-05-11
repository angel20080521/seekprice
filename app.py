"""
seekprice – 报价单 vs 合同 比对系统
端口: 5003
"""
import logging
import os
import re
import uuid

from flask import Flask, jsonify, render_template, request

import docx
import pdfplumber

# ─── Flask setup ─────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_FOLDER = "/tmp/seekprice_uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXT = {"doc", "docx", "pdf"}

# Hard-coded extension map: maps known user-provided extensions to literal strings.
# Values come entirely from this constant, never from user input.
_EXT_MAP: dict[str, str] = {"doc": "doc", "docx": "docx", "pdf": "pdf"}

# Tolerance for floating-point financial comparisons (covers rounding differences)
NUMERIC_COMPARISON_TOLERANCE = 0.02

logger = logging.getLogger(__name__)


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in ALLOWED_EXT


def _get_safe_ext(filename: str) -> str | None:
    """
    Return a whitelisted extension string (from _EXT_MAP values, not user input),
    or None if the extension is not recognised.
    """
    if "." not in filename:
        return None
    raw = filename.rsplit(".", 1)[-1].lower()
    return _EXT_MAP.get(raw)          # always returns a literal from _EXT_MAP


def _safe_upload_path(filename_ext: str) -> str:
    """
    Build a safe temporary file path inside UPLOAD_FOLDER.
    filename_ext must be a value from _EXT_MAP (caller's responsibility).
    Returns an absolute path; raises ValueError if it escapes UPLOAD_FOLDER.
    """
    name = f"{uuid.uuid4().hex}.{filename_ext}"
    path = os.path.realpath(os.path.join(UPLOAD_FOLDER, name))
    base = os.path.realpath(UPLOAD_FOLDER)
    if not path.startswith(base + os.sep):
        raise ValueError("生成的文件路径异常")
    return path


# ─── Number / string helpers ─────────────────────────────────────────────────

def clean_amount(raw) -> str | None:
    """Strip currency symbols, commas, spaces; return numeric string or None."""
    if not raw:
        return None
    s = re.sub(r"[¥￥,\s]", "", str(raw))
    m = re.search(r"\d[\d.]*", s)
    return m.group() if m else None


def clean_rate(raw) -> str | None:
    """Normalise a tax-rate value to e.g. '13%'."""
    if not raw:
        return None
    s = re.sub(r"\s+", "", str(raw))
    m = re.search(r"[\d.]+%?", s)
    if m:
        v = m.group()
        return v if "%" in v else v + "%"
    return None


def norm_text(s) -> str:
    """Collapse all whitespace for fingerprinting / matching."""
    return re.sub(r"\s+", "", str(s or "")).upper()


# ─── Column identification ────────────────────────────────────────────────────

def identify_columns(headers: list) -> dict:
    """
    Given a header row, return a mapping: field_name → column_index.
    Recognised fields:
      id, name (产品名称/项目内容), model (产品型号),
      qty, tax_rate, tax_amount, amount_ex, amount_inc
    """
    fields: dict[str, int] = {}
    for i, h in enumerate(headers):
        if h is None:
            continue
        n = re.sub(r"\s+", "", str(h))          # collapse whitespace

        if re.search(r"编号|序号", n) and "id" not in fields:
            fields["id"] = i
        # Distinguish brand/project name from product model
        if re.search(r"产品名称|品名|项目内容", n) and "name" not in fields:
            fields["name"] = i
        if re.search(r"产品型号|型号", n) and "model" not in fields:
            fields["model"] = i
        if re.search(r"数量", n) and "qty" not in fields:
            fields["qty"] = i
        if re.search(r"税率", n) and "tax_rate" not in fields:
            fields["tax_rate"] = i
        if re.search(r"税额", n) and "tax_amount" not in fields:
            fields["tax_amount"] = i
        # Total inclusive-tax amount: 含税金额 or 金额(含税)
        # Exclude unit-price columns (单价)
        if re.search(r"含税金额|金额.{0,5}含税", n) and "amount_inc" not in fields:
            fields["amount_inc"] = i
        # Total exclusive-tax amount: 金额(未税) / 金额(不含税)
        if re.search(r"金额.{0,5}(未税|不含税)", n) and "amount_ex" not in fields:
            fields["amount_ex"] = i

    return fields


# ─── Skip-row detection ───────────────────────────────────────────────────────

# Skip rows that are totals/subtotals
_SKIP_RE = re.compile(r"总计|合计|小计")


def is_skip_row(row: list) -> bool:
    first = str(row[0]).strip() if row and row[0] is not None else ""
    return not first or bool(_SKIP_RE.search(first))


# ─── Table → item-list parser ─────────────────────────────────────────────────

def parse_table(
    table: list[list],
    tbl_idx: int,
    page_num: int | None = None,
) -> list[dict]:
    """
    Convert a 2-D table (list of lists of str|None) into a list of item dicts.
    Only returns rows that contain at least one numeric field.
    """
    if not table or len(table) < 2:
        return []

    # Locate the header row (first row containing 数量 or 金额 or 税率)
    header_row_idx = 0
    for ri, row in enumerate(table):
        combined = "".join(str(c) for c in row if c)
        if re.search(r"数量|金额|税率", combined):
            header_row_idx = ri
            break

    headers = table[header_row_idx]
    cols = identify_columns(headers)

    # Must have at least one useful numeric column
    useful = set(cols) & {"qty", "amount_inc", "amount_ex", "tax_rate", "tax_amount"}
    if not useful:
        return []

    items = []
    for ri, row in enumerate(table[header_row_idx + 1 :], header_row_idx + 1):
        if is_skip_row(row):
            continue

        item: dict = {"_tbl": tbl_idx, "_row": ri}
        if page_num is not None:
            item["_page"] = page_num
            item["position"] = f"第{page_num}页-表格{tbl_idx + 1}-第{ri}行"
        else:
            item["position"] = f"表格{tbl_idx + 1}-第{ri}行"

        for field, ci in cols.items():
            if ci >= len(row):
                continue
            raw = str(row[ci]).strip() if row[ci] is not None else ""
            if field in ("qty", "amount_inc", "amount_ex", "tax_amount"):
                item[field] = clean_amount(raw)
            elif field == "tax_rate":
                item[field] = clean_rate(raw)
            else:
                item[field] = norm_text(raw)   # id / name / model: normalised text

        # Only keep rows that carry at least one numeric value
        if any(item.get(f) for f in ("qty", "amount_inc", "amount_ex", "tax_amount")):
            items.append(item)

    return items


# ─── File extractors ──────────────────────────────────────────────────────────

def extract_docx(path: str) -> list[dict]:
    doc = docx.Document(path)
    rows: list[dict] = []
    for ti, tbl in enumerate(doc.tables):
        data = [[cell.text for cell in row.cells] for row in tbl.rows]
        rows.extend(parse_table(data, ti))
    return rows


def extract_pdf(path: str) -> list[dict]:
    rows: list[dict] = []
    tbl_global = 0
    with pdfplumber.open(path) as pdf:
        for pnum, page in enumerate(pdf.pages, 1):
            for tbl in page.extract_tables() or []:
                parsed = parse_table(tbl, tbl_global, pnum)
                rows.extend(parsed)
                tbl_global += 1
    return rows


def extract(path: str) -> list[dict]:
    ext = path.rsplit(".", 1)[-1].lower()
    if ext in ("doc", "docx"):
        return extract_docx(path)
    if ext == "pdf":
        return extract_pdf(path)
    return []


# ─── Row matching ─────────────────────────────────────────────────────────────

def _best_key(row: dict) -> str:
    """Return the most-specific text key for matching: model > name."""
    return norm_text(row.get("model") or row.get("name") or "")


def _fp_full(row: dict) -> str:
    """Fingerprint: id + model/name (both normalised, joined with |)."""
    id_ = norm_text(row.get("id", ""))
    key = _best_key(row)
    return f"{id_}|{key}" if id_ and key else ""


def _fp_name(row: dict) -> str:
    """Model/name-only fingerprint (min 3 chars to avoid trivial matches)."""
    key = _best_key(row)
    return key if len(key) >= 3 else ""


def match_rows(
    contract_rows: list[dict],
    quote_rows: list[dict],
) -> list[tuple[dict, dict]]:
    """
    Match each contract row to a quote row.
    Priority:
      1. Exact (id + name) fingerprint match
      2. Name-only fingerprint match
      3. Sequential fallback
    Returns list of (contract_row, quote_row) pairs.
    """
    used_q: set[int] = set()

    # Pre-build lookup indices for quote rows
    q_by_full: dict[str, list[int]] = {}
    q_by_name: dict[str, list[int]] = {}
    for qi, qr in enumerate(quote_rows):
        fp = _fp_full(qr)
        if fp:
            q_by_full.setdefault(fp, []).append(qi)
        fn = _fp_name(qr)
        if fn:
            q_by_name.setdefault(fn, []).append(qi)

    pairs: list[tuple[dict, dict | None]] = []

    for cr in contract_rows:
        matched_qi: int | None = None

        # 1) Full fingerprint
        fp = _fp_full(cr)
        if fp and fp in q_by_full:
            for qi in q_by_full[fp]:
                if qi not in used_q:
                    matched_qi = qi
                    break

        # 2) Name-only fingerprint
        if matched_qi is None:
            fn = _fp_name(cr)
            if fn and fn in q_by_name:
                for qi in q_by_name[fn]:
                    if qi not in used_q:
                        matched_qi = qi
                        break

        if matched_qi is not None:
            used_q.add(matched_qi)
            pairs.append((cr, quote_rows[matched_qi]))
        else:
            pairs.append((cr, None))   # will be resolved in sequential pass

    # 3) Sequential fallback for unresolved contract rows
    leftover_q = [qi for qi in range(len(quote_rows)) if qi not in used_q]
    none_indices = [i for i, (_, qr) in enumerate(pairs) if qr is None]
    for pair_i, qi in zip(none_indices, leftover_q):
        cr, _ = pairs[pair_i]
        pairs[pair_i] = (cr, quote_rows[qi])

    return [(cr, qr) for cr, qr in pairs if qr is not None]


# ─── Comparison ───────────────────────────────────────────────────────────────

COMPARE_FIELDS = [
    ("qty",        "数量"),
    ("amount_ex",  "金额（不含税）"),
    ("tax_rate",   "税率"),
    ("tax_amount", "税额"),
    ("amount_inc", "金额（含税）"),
]


def compare_documents(
    contract_rows: list[dict],
    quote_rows: list[dict],
) -> tuple[list[dict], list[dict]]:
    """
    Returns (mismatches, matched_pairs_info).
    mismatches: list of dicts with position / field / values.
    matched_pairs_info: human-readable summary of which rows were paired.
    """
    pairs = match_rows(contract_rows, quote_rows)
    mismatches: list[dict] = []
    matched_info: list[dict] = []

    for cr, qr in pairs:
        matched_info.append({
            "contractPos": cr.get("position", ""),
            "quotePos": qr.get("position", ""),
            "name": (_best_key(cr) or _best_key(qr)),
        })

        for fkey, fname in COMPARE_FIELDS:
            cv = cr.get(fkey)
            qv = qr.get(fkey)

            # Skip when either side lacks the field
            if cv is None or qv is None:
                continue

            # Numeric comparison with configurable tolerance for rounding differences
            try:
                cf = float(str(cv).replace(",", "").replace("%", ""))
                qf = float(str(qv).replace(",", "").replace("%", ""))
                same = abs(cf - qf) <= NUMERIC_COMPARISON_TOLERANCE
            except ValueError:
                same = str(cv).strip() == str(qv).strip()

            if not same:
                mismatches.append({
                    "contractPosition": cr.get("position", ""),
                    "quotePosition":    qr.get("position", ""),
                    "field":            fname,
                    "contractValue":    str(cv),
                    "quoteValue":       str(qv),
                })

    return mismatches, matched_info


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/compare", methods=["POST"])
def do_compare():
    contract_f = request.files.get("contract")
    quote_f    = request.files.get("quote")

    if not contract_f or not quote_f:
        return jsonify(error="请上传合同文件和报价单文件"), 400
    if not allowed(contract_f.filename) or not allowed(quote_f.filename):
        return jsonify(error="仅支持 .doc / .docx / .pdf 格式"), 400

    c_ext = _get_safe_ext(contract_f.filename)
    q_ext = _get_safe_ext(quote_f.filename)
    if not c_ext or not q_ext:
        return jsonify(error="仅支持 .doc / .docx / .pdf 格式"), 400

    try:
        c_path = _safe_upload_path(c_ext)
        q_path = _safe_upload_path(q_ext)
    except ValueError:
        return jsonify(error="文件路径生成失败，请重试"), 500

    contract_f.save(c_path)
    quote_f.save(q_path)

    try:
        c_rows = extract(c_path)
        q_rows = extract(q_path)

        if not c_rows:
            return jsonify(error="无法从合同文件中提取数据，请确认文件包含数值表格"), 400
        if not q_rows:
            return jsonify(error="无法从报价单文件中提取数据，请确认文件包含数值表格"), 400

        mismatches, matched_info = compare_documents(c_rows, q_rows)

        return jsonify(
            contractItems=len(c_rows),
            quoteItems=len(q_rows),
            matchedPairs=len(matched_info),
            mismatchCount=len(mismatches),
            mismatches=mismatches,
            matched=matched_info,
        )

    except Exception:
        logger.exception("Error processing uploaded files")
        return jsonify(error="处理文件时出错，请检查文件格式是否正确"), 500

    finally:
        for p in (c_path, q_path):
            base = os.path.realpath(UPLOAD_FOLDER)
            real_p = os.path.realpath(p)
            if real_p.startswith(base + os.sep):
                try:
                    os.remove(real_p)
                except OSError:
                    pass


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5003, debug=False)
