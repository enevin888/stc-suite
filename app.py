import difflib
import io
import re
from dataclasses import dataclass

import pandas as pd
import streamlit as st

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import docx
except ImportError:
    docx = None


# ============================================================
# SmartTech Engineering Review Suite
# Expanded scope:
# - SOO revision review
# - Points list review from PDF table / Excel / CSV
# - Wiring diagram review from extracted PDF/CAD text
# - Package review dashboard
# ============================================================

st.set_page_config(
    page_title="SmartTech Engineering Review Suite",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

NAVY = "#020C3D"
ORANGE = "#EC7626"
LIGHT_GRAY = "#E3E6EB"
WHITE = "#FFFFFF"
DARK_TEXT = "#111827"
MUTED_TEXT = "#5B6472"


# ============================================================
# Basic file readers
# ============================================================
def read_text_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""

    uploaded_file.seek(0)
    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    if name.endswith(".docx"):
        if docx is None:
            st.error("python-docx is required for Word files.")
            return ""

        document = docx.Document(io.BytesIO(data))
        parts = [p.text for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    parts.append(row_text)
        return "\n".join(parts)

    if name.endswith(".pdf"):
        return extract_pdf_text(io.BytesIO(data))

    return data.decode("utf-8", errors="ignore")


def extract_pdf_text(file_like) -> str:
    if pdfplumber is None:
        return ""

    text_parts = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                text_parts.append(text)
    return "\n\n".join(text_parts)


def extract_pdf_tables(file_like):
    if pdfplumber is None:
        return []

    tables = []
    with pdfplumber.open(file_like) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_tables = page.extract_tables() or []
            for table in page_tables:
                if table and len(table) > 1:
                    tables.append((page_num, table))
    return tables


def read_points_table(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()

    uploaded_file.seek(0)
    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(data))
        return normalize_points_df(df)

    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(data))
        return normalize_points_df(df)

    if name.endswith(".pdf"):
        tables = extract_pdf_tables(io.BytesIO(data))
        dfs = []

        for page_num, table in tables:
            header = table[0]
            rows = table[1:]
            df = pd.DataFrame(rows, columns=header)
            df["Source Page"] = page_num
            dfs.append(df)

        if dfs:
            return normalize_points_df(pd.concat(dfs, ignore_index=True))

        # fallback text parse
        text = extract_pdf_text(io.BytesIO(data))
        return parse_points_from_text(text)

    return pd.DataFrame()


# ============================================================
# Diff helpers
# ============================================================
def get_diff_counts(original: str, updated: str) -> dict:
    diff = list(difflib.ndiff(original.splitlines(), updated.splitlines()))
    added = sum(1 for line in diff if line.startswith("+ "))
    removed = sum(1 for line in diff if line.startswith("- "))
    unchanged = sum(1 for line in diff if line.startswith("  "))
    return {"added": added, "removed": removed, "changed": added + removed, "unchanged": unchanged}


def make_html_side_by_side_diff(original: str, updated: str, context_only: bool, left_title="Original", right_title="Updated") -> str:
    return difflib.HtmlDiff(tabsize=4, wrapcolumn=115).make_table(
        fromlines=original.splitlines(),
        tolines=updated.splitlines(),
        fromdesc=left_title,
        todesc=right_title,
        context=context_only,
        numlines=5,
    )


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower()).strip()


def find_added_removed_lines(original: str, updated: str):
    diff = list(difflib.ndiff(original.splitlines(), updated.splitlines()))
    added = [line[2:].strip() for line in diff if line.startswith("+ ") and line[2:].strip()]
    removed = [line[2:].strip() for line in diff if line.startswith("- ") and line[2:].strip()]
    return added, removed


# ============================================================
# SOO engineering comments
# ============================================================
def find_alarm_names(text: str) -> set:
    alarms = set()
    patterns = [
        r"\b[A-Z]{1,8}_[A-Z0-9]{1,10}\b",
        r"\bAlarm\d+\.\(?\d+\)?\b",
        r"\b[A-Z]{2,8}\s*-\s*[^.\n]+",
    ]

    for pattern in patterns:
        for match in re.findall(pattern, text):
            cleaned = match.strip(" .:-")
            if any(key in cleaned.upper() for key in ["ALM", "HI", "LO", "FAIL", "FAULT", "FREEZE", "ZP", "MAT", "SAT", "DAT", "CSAT"]):
                alarms.add(cleaned)

    return alarms


def find_setpoints(text: str) -> set:
    setpoints = set()
    patterns = [
        r"[^.\n]*(?:setpoint|adj\.|adj|°f|%rh|h2o|in\. wc|btu)[^.\n]*",
        r"[^.\n]*\d+(?:\.\d+)?\s*(?:°f|%rh|%|\" h2o|in\. wc|btu / lbm)[^.\n]*",
    ]

    for pattern in patterns:
        for match in re.findall(pattern, text, flags=re.IGNORECASE):
            cleaned = " ".join(match.strip().split())
            if len(cleaned) > 8:
                setpoints.add(cleaned)

    return setpoints


def detect_numbering_changes(added_lines, removed_lines):
    numbering_pattern = r"^\s*((?:#{1,6}\s*)?[A-Z]\.|\d+\.|[a-z]\.|\(\d+\)|\([a-z]\))"
    added_nums = [line for line in added_lines if re.search(numbering_pattern, line)]
    removed_nums = [line for line in removed_lines if re.search(numbering_pattern, line)]
    return added_nums, removed_nums


def make_engineering_comments(original: str, updated: str):
    comments = []
    added_lines, removed_lines = find_added_removed_lines(original, updated)

    original_alarms = find_alarm_names(original)
    updated_alarms = find_alarm_names(updated)
    added_alarms = sorted(updated_alarms - original_alarms)
    removed_alarms = sorted(original_alarms - updated_alarms)

    original_setpoints = find_setpoints(original)
    updated_setpoints = find_setpoints(updated)
    added_setpoints = sorted(updated_setpoints - original_setpoints)
    removed_setpoints = sorted(original_setpoints - updated_setpoints)

    added_numbering, removed_numbering = detect_numbering_changes(added_lines, removed_lines)

    def add(priority, category, comment, items=None):
        comments.append({"priority": priority, "category": category, "comment": comment, "items": items or []})

    if added_alarms:
        add("Review", "Added alarms", "New alarm references were added. Confirm trigger, delay, annunciation, and reset/clear logic.", added_alarms[:12])

    if removed_alarms:
        add("High", "Removed alarms", "Alarm references appear to have been removed. Confirm these were intentionally deleted.", removed_alarms[:12])

    if added_setpoints or removed_setpoints:
        add("Review", "Setpoint / threshold changes", "Setpoint or adjustable-value wording changed. Confirm values match design intent.", (added_setpoints[:6] + removed_setpoints[:6]))

    if added_numbering or removed_numbering:
        add("Review", "Numbering / formatting", "Numbered or lettered sequence lines changed. Verify numbering is clean and sequential.", (added_numbering[:8] + removed_numbering[:8]))

    updated_lower = normalize_text(updated)

    if "shall modulate" in updated_lower and "set position" in updated_lower:
        add("High", "Possible control contradiction", "Text includes both modulation and fixed-position language. Confirm the device is not described inconsistently.")

    if "shared" in updated_lower and "output" in updated_lower:
        add("Review", "Shared control output", "Shared output logic detected. Confirm single-unit, multi-unit, and disabled-unit behavior.")

    if "damper" in updated_lower and "pressure" in updated_lower and "fan" in updated_lower:
        add("Review", "Fan / damper coordination", "Fan, damper, and pressure control are referenced. Confirm only the intended device controls pressure.")

    if "freeze" in updated_lower or "low temperature" in updated_lower:
        add("Info", "Freeze protection", "Freeze / low-temperature protection exists. Confirm it overrides normal cooling/economizer control.")

    if not comments:
        add("Info", "No major flags detected", "No obvious alarm, setpoint, numbering, or contradiction flags were detected.")

    return comments


# ============================================================
# Points list helpers
# ============================================================
def normalize_col_name(col):
    if col is None:
        return ""
    return re.sub(r"\s+", " ", str(col).strip())


def normalize_points_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [normalize_col_name(c) for c in df.columns]
    df = df.dropna(how="all")

    # remove duplicate blank columns
    df = df.loc[:, [c for c in df.columns if c and not c.lower().startswith("unnamed")]]

    tag_col = find_column(df, ["tag", "point", "point name", "name", "object", "device tag"])
    type_col = find_column(df, ["type", "point type", "io type", "i/o type", "object type", "signal"])
    desc_col = find_column(df, ["description", "desc", "point description", "label"])
    panel_col = find_column(df, ["panel", "controller", "location", "dcp", "panel name"])
    alarm_col = find_column(df, ["alarm", "alarm class", "priority", "category"])

    normalized = pd.DataFrame()
    normalized["Tag"] = df[tag_col].astype(str).str.strip() if tag_col else ""
    normalized["Type"] = df[type_col].astype(str).str.strip() if type_col else ""
    normalized["Description"] = df[desc_col].astype(str).str.strip() if desc_col else ""
    normalized["Panel"] = df[panel_col].astype(str).str.strip() if panel_col else ""
    normalized["Alarm"] = df[alarm_col].astype(str).str.strip() if alarm_col else ""

    if "Source Page" in df.columns:
        normalized["Source Page"] = df["Source Page"]

    normalized = normalized[normalized["Tag"].astype(str).str.strip().ne("")]
    normalized = normalized[~normalized["Tag"].astype(str).str.lower().isin(["nan", "none", "tag", "point", "point name"])]
    normalized["Tag Key"] = normalized["Tag"].astype(str).str.upper().str.replace(r"\s+", "", regex=True)

    return normalized.reset_index(drop=True)


def find_column(df, options):
    lowered = {str(c).lower().strip(): c for c in df.columns}
    for option in options:
        if option in lowered:
            return lowered[option]

    for c in df.columns:
        cl = str(c).lower()
        for option in options:
            if option in cl:
                return c

    return None


def parse_points_from_text(text: str) -> pd.DataFrame:
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # Simple tag/type parser fallback
        m = re.match(r"([A-Za-z0-9_\-\.]+)\s+((AI|AO|BI|BO|AV|BV|MSV|MI|MO)\b.*)", line, flags=re.IGNORECASE)
        if m:
            tag = m.group(1)
            rest = m.group(2)
            point_type = m.group(3).upper()
            rows.append({"Tag": tag, "Type": point_type, "Description": rest, "Panel": "", "Alarm": ""})

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["Tag Key"] = df["Tag"].astype(str).str.upper().str.replace(r"\s+", "", regex=True)
    return df


def compare_points(original_df: pd.DataFrame, updated_df: pd.DataFrame):
    results = {}

    if original_df.empty or updated_df.empty:
        return {
            "summary": pd.DataFrame(),
            "added": pd.DataFrame(),
            "removed": pd.DataFrame(),
            "changed": pd.DataFrame(),
            "comments": [{"priority": "High", "category": "Extraction issue", "comment": "Could not extract a usable points table from one or both files.", "items": []}],
        }

    original_keys = set(original_df["Tag Key"])
    updated_keys = set(updated_df["Tag Key"])

    added_keys = sorted(updated_keys - original_keys)
    removed_keys = sorted(original_keys - updated_keys)
    common_keys = sorted(original_keys & updated_keys)

    added = updated_df[updated_df["Tag Key"].isin(added_keys)].copy()
    removed = original_df[original_df["Tag Key"].isin(removed_keys)].copy()

    changed_rows = []
    for key in common_keys:
        o = original_df[original_df["Tag Key"] == key].iloc[0]
        u = updated_df[updated_df["Tag Key"] == key].iloc[0]

        for col in ["Type", "Description", "Panel", "Alarm"]:
            ov = str(o.get(col, "")).strip()
            uv = str(u.get(col, "")).strip()
            if ov != uv:
                changed_rows.append({
                    "Tag": u.get("Tag", o.get("Tag", key)),
                    "Field": col,
                    "Original": ov,
                    "Updated": uv,
                })

    changed = pd.DataFrame(changed_rows)

    comments = []
    if not added.empty:
        comments.append({"priority": "Review", "category": "Added points", "comment": "Points were added to the updated list. Confirm wiring, SOO references, and controller capacity.", "items": added["Tag"].head(12).tolist()})
    if not removed.empty:
        comments.append({"priority": "High", "category": "Removed points", "comment": "Points were removed from the updated list. Confirm they are not still required by SOO or wiring diagrams.", "items": removed["Tag"].head(12).tolist()})
    if not changed.empty:
        comments.append({"priority": "Review", "category": "Changed point fields", "comment": "Point type, description, panel, or alarm fields changed.", "items": changed["Tag"].head(12).tolist()})

    if not comments:
        comments.append({"priority": "Info", "category": "No major point changes", "comment": "No added, removed, or changed point rows were detected.", "items": []})

    summary = pd.DataFrame([
        {"Metric": "Original points", "Count": len(original_df)},
        {"Metric": "Updated points", "Count": len(updated_df)},
        {"Metric": "Added points", "Count": len(added)},
        {"Metric": "Removed points", "Count": len(removed)},
        {"Metric": "Changed fields", "Count": len(changed)},
    ])

    return {"summary": summary, "added": added, "removed": removed, "changed": changed, "comments": comments}


# ============================================================
# Wiring diagram helpers
# ============================================================
def extract_wiring_entities(text: str):
    entities = {
        "terminals": set(re.findall(r"\bTB[-\s]?\d+(?:[-\.]\d+)?\b", text, flags=re.IGNORECASE)),
        "relays": set(re.findall(r"\b(?:CR|R|K)[-\s]?\d+\b", text, flags=re.IGNORECASE)),
        "controllers": set(re.findall(r"\b(?:DCP|PLC|CTRL|CONTROLLER)[-\s]?[A-Z0-9]+\b", text, flags=re.IGNORECASE)),
        "points": set(re.findall(r"\b(?:AI|AO|BI|BO|DI|DO)[-\s]?\d+\b", text, flags=re.IGNORECASE)),
        "wires": set(re.findall(r"\b(?:WIRE|W)[-\s]?\d+[A-Z]?\b", text, flags=re.IGNORECASE)),
    }

    return {k: sorted({v.upper().replace(" ", "") for v in vals}) for k, vals in entities.items()}


def compare_wiring(original_text: str, updated_text: str):
    o = extract_wiring_entities(original_text)
    u = extract_wiring_entities(updated_text)

    rows = []
    comments = []

    for category in ["terminals", "relays", "controllers", "points", "wires"]:
        original_set = set(o[category])
        updated_set = set(u[category])

        added = sorted(updated_set - original_set)
        removed = sorted(original_set - updated_set)

        rows.append({
            "Category": category.title(),
            "Original Count": len(original_set),
            "Updated Count": len(updated_set),
            "Added": len(added),
            "Removed": len(removed),
        })

        if added:
            comments.append({"priority": "Review", "category": f"Added {category}", "comment": f"{len(added)} {category} references appear to have been added.", "items": added[:12]})
        if removed:
            comments.append({"priority": "High", "category": f"Removed {category}", "comment": f"{len(removed)} {category} references appear to have been removed.", "items": removed[:12]})

    if not comments:
        comments.append({"priority": "Info", "category": "No major wiring entity changes", "comment": "No major terminal, relay, controller, I/O, or wire-reference changes were detected from extracted text.", "items": []})

    return pd.DataFrame(rows), comments


# ============================================================
# UI helpers
# ============================================================
def render_metric_row(metrics):
    html = '<div class="summary-row">'
    for icon, value, label, cls in metrics:
        html += f"""
        <div class="metric-card">
            <div class="metric-icon {cls}">{icon}</div>
            <div>
                <div class="metric-number">{value}</div>
                <div class="metric-label">{label}</div>
            </div>
        </div>
        """
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def render_comment_cards(comments):
    priority_colors = {"High": "#EF4444", "Review": ORANGE, "Info": "#64748B"}
    cards = ""
    for comment in comments:
        color = priority_colors.get(comment["priority"], "#64748B")
        items_html = ""
        if comment["items"]:
            safe_items = "".join(f"<li>{str(item)}</li>" for item in comment["items"])
            items_html = f"<ul>{safe_items}</ul>"

        cards += f"""
        <div class="comment-card">
            <div class="comment-top">
                <span class="priority-pill" style="background:{color};">{comment["priority"]}</span>
                <span class="comment-category">{comment["category"]}</span>
            </div>
            <div class="comment-text">{comment["comment"]}</div>
            {items_html}
        </div>
        """
    return cards


def render_comments_panel(title, comments):
    st.markdown(
        f"""
        <div class="review-panel">
            <div class="panel-title">{title}</div>
            <div class="comment-grid">
                {render_comment_cards(comments)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_diff_panel(original_text, updated_text, left_title="Original", right_title="Updated"):
    context_only = st.toggle("Show only changed areas", value=True, help="Turn off to show the full document.")
    html_diff = make_html_side_by_side_diff(original_text, updated_text, context_only=context_only, left_title=left_title, right_title=right_title)

    st.markdown(
        f"""
        <div class="diff-panel">
            <div class="panel-title">Line Changes</div>
            <div class="legend">
                <span class="legend-added">Added</span>
                <span class="legend-removed">Removed</span>
                <span class="legend-changed">Changed Within Line</span>
            </div>
            <div class="diff-wrapper">
                {html_diff}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def file_uploader_card(title, help_text, key, types):
    st.markdown(
        f"""
        <div class="upload-card">
            <div class="upload-title">{title}</div>
            <div class="upload-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    return st.file_uploader(title, type=types, key=key, label_visibility="collapsed")


# ============================================================
# CSS
# ============================================================
st.markdown(
    f"""
    <style>
        html, body, [class*="css"] {{ font-family: "Segoe UI", Arial, sans-serif; }}
        .stApp {{ background: {LIGHT_GRAY}; }}
        section[data-testid="stSidebar"] {{ display: none !important; }}
        .main .block-container {{ padding-top: 1.25rem; padding-left: 2rem; padding-right: 2rem; max-width: 98%; }}

        .brand-header {{
            background: linear-gradient(90deg, {WHITE} 0%, #F8FAFC 72%, rgba(236,118,38,0.24) 100%);
            border: 1px solid #D7DCE5; border-radius: 22px; padding: 20px 26px; margin-bottom: 18px;
            box-shadow: 0 5px 18px rgba(2,12,61,0.10);
        }}
        .brand-row {{ display: flex; align-items: center; gap: 24px; }}
        .fallback-logo {{
            font-size: 30px; font-weight: 950; color: {NAVY}; padding-right: 24px; border-right: 2px solid #D7DCE5;
            line-height: 1.02; min-width: 190px;
        }}
        .fallback-logo span {{ color: {ORANGE}; }}
        .brand-title {{ font-size: 38px; font-weight: 950; color: {NAVY}; letter-spacing: -0.5px; margin-bottom: 5px; line-height: 1.05; }}
        .brand-subtitle {{ font-size: 16px; color: {MUTED_TEXT}; font-weight: 650; }}

        .mode-panel, .upload-card, .review-panel, .diff-panel {{
            background: {WHITE}; border: 1px solid #D7DCE5; border-radius: 20px; padding: 18px;
            box-shadow: 0 5px 18px rgba(2,12,61,0.10); margin-top: 12px;
        }}
        .upload-card {{ border-top: 6px solid {ORANGE}; }}
        .upload-title {{ font-size: 22px; font-weight: 950; color: {NAVY}; margin-bottom: 4px; }}
        .upload-help {{ font-size: 14px; color: {MUTED_TEXT}; font-weight: 650; }}
        .panel-title {{ font-size: 25px; font-weight: 950; color: {NAVY}; letter-spacing: -0.2px; margin-bottom: 10px; }}

        [data-testid="stFileUploader"] label {{ display: none; }}
        [data-testid="stFileUploaderDropzone"] {{
            background-color: #FAFBFC; border: 2px dashed #AAB4C3; border-radius: 14px; min-height: 92px; padding: 16px !important;
        }}
        [data-testid="stFileUploaderDropzone"]:hover {{ border-color: {ORANGE}; background-color: #FFF7F1; }}
        [data-testid="stFileUploaderDropzone"] * {{ color: {NAVY} !important; font-weight: 850; }}
        [data-testid="stFileUploaderFile"] {{
            background: #F8FAFC !important; border: 1px solid #CBD5E1 !important; border-radius: 12px !important;
            color: {NAVY} !important; padding: 8px !important; min-height: 52px;
        }}
        [data-testid="stFileUploaderFile"] * {{ color: {NAVY} !important; font-weight: 800 !important; }}

        button[kind="secondary"] {{
            background-color: {NAVY} !important; color: white !important; border-radius: 9px !important; border: none !important; font-weight: 800 !important;
        }}

        .summary-row {{ display: grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap: 14px; margin: 16px 0 18px 0; }}
        .metric-card {{
            background: {WHITE}; border: 1px solid #D7DCE5; border-radius: 16px; padding: 16px 18px;
            box-shadow: 0 4px 14px rgba(2,12,61,0.08); display: flex; align-items: center; gap: 14px;
        }}
        .metric-icon {{
            width: 42px; height: 42px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
            color: white; font-size: 24px; font-weight: 950;
        }}
        .metric-icon.added {{ background: #22A95A; }}
        .metric-icon.removed {{ background: #EF4444; }}
        .metric-icon.changed {{ background: {ORANGE}; }}
        .metric-icon.flags {{ background: {NAVY}; }}
        .metric-number {{ font-size: 30px; font-weight: 950; color: {NAVY}; line-height: 1; }}
        .metric-label {{ font-size: 13px; color: {MUTED_TEXT}; margin-top: 4px; font-weight: 700; }}

        .comment-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(310px, 1fr)); gap: 12px; }}
        .comment-card {{ border: 1px solid #D7DCE5; border-radius: 15px; padding: 14px 15px; background: #FAFBFC; }}
        .comment-top {{ display: flex; align-items: center; gap: 9px; margin-bottom: 8px; }}
        .priority-pill {{ color: white; font-size: 12px; font-weight: 950; padding: 4px 9px; border-radius: 999px; }}
        .comment-category {{ color: {NAVY}; font-size: 15px; font-weight: 950; }}
        .comment-text {{ color: {DARK_TEXT}; font-size: 14px; font-weight: 600; line-height: 1.35; }}
        .comment-card ul {{ margin-bottom: 0; padding-left: 18px; font-size: 12px; color: {MUTED_TEXT}; font-weight: 600; }}

        .legend {{ display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }}
        .legend span {{ padding: 6px 12px; border-radius: 999px; font-size: 13px; font-weight: 900; color: {DARK_TEXT}; border: 1px solid rgba(0,0,0,0.12); }}
        .legend-added {{ background-color: #B7F7C1; }}
        .legend-removed {{ background-color: #FFB6BD; }}
        .legend-changed {{ background-color: #FFE88A; }}

        .diff-wrapper {{ background: {WHITE}; border: 1px solid #CCD3DE; border-radius: 14px; overflow-x: auto; max-height: 74vh; overflow-y: auto; }}
        table.diff {{ font-family: Consolas, "Courier New", monospace !important; border-collapse: collapse; width: 100%; font-size: 14px; color: {DARK_TEXT} !important; background: {WHITE} !important; }}
        .diff_header {{ background-color: {NAVY} !important; color: white !important; font-weight: 950 !important; text-align: left; padding: 10px; position: sticky; top: 0; z-index: 2; }}
        td.diff_header {{ width: 50px; text-align: right; color: {NAVY} !important; background-color: #E6EAF0 !important; font-weight: 800; border-right: 1px solid #B5BDCA; position: static; }}
        .diff_next {{ background-color: #F3F5F8 !important; color: {NAVY} !important; font-weight: 900; text-align: center; }}
        .diff_add {{ background-color: #B7F7C1 !important; color: #04210A !important; font-weight: 800; }}
        .diff_sub {{ background-color: #FFB6BD !important; color: #3B060B !important; font-weight: 800; }}
        .diff_chg {{ background-color: #FFE88A !important; color: #332600 !important; font-weight: 900; }}
        span.diff_add {{ background-color: #35D06D !important; color: #052E13 !important; padding: 1px 3px; border-radius: 3px; font-weight: 950; }}
        span.diff_sub {{ background-color: #F35D71 !important; color: #450A0A !important; padding: 1px 3px; border-radius: 3px; font-weight: 950; }}
        span.diff_chg {{ background-color: #FACC15 !important; color: #332600 !important; padding: 1px 3px; border-radius: 3px; font-weight: 950; }}
        table.diff td {{ padding: 7px 9px; vertical-align: top; border-bottom: 1px solid #E8ECF2; color: {DARK_TEXT} !important; white-space: pre-wrap; line-height: 1.38; }}

        div[data-testid="stRadio"] label, div[data-testid="stToggle"] label {{ color: {NAVY} !important; font-weight: 800 !important; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# App layout
# ============================================================
st.markdown(
    """
    <div class="brand-header">
        <div class="brand-row">
            <div class="fallback-logo"><span>smart</span> tech<br>contracting</div>
            <div>
                <div class="brand-title">Engineering Review Suite</div>
                <div class="brand-subtitle">Review SOO changes, points list changes, wiring diagram changes, and package-level consistency.</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="mode-panel"><div class="panel-title">Choose Review Type</div></div>', unsafe_allow_html=True)

mode = st.radio(
    "Review type",
    [
        "SOO Revision Review",
        "Points List Review",
        "Wiring Diagram Review",
        "Full Package Review",
    ],
    horizontal=True,
    label_visibility="collapsed",
)


# ============================================================
# SOO Review
# ============================================================
if mode == "SOO Revision Review":
    col1, col2 = st.columns(2, gap="large")
    with col1:
        original_file = file_uploader_card("Original SOO", "Upload original SOO markdown, text, Word, or PDF.", "soo_original", ["md", "txt", "docx", "pdf"])
    with col2:
        updated_file = file_uploader_card("Updated SOO", "Upload updated SOO markdown, text, Word, or PDF.", "soo_updated", ["md", "txt", "docx", "pdf"])

    if original_file and updated_file:
        original_text = read_text_file(original_file)
        updated_text = read_text_file(updated_file)

        counts = get_diff_counts(original_text, updated_text)
        comments = make_engineering_comments(original_text, updated_text)
        flag_count = sum(1 for c in comments if c["priority"] in ["High", "Review"])

        render_metric_row([
            ("+", counts["added"], "Added Lines", "added"),
            ("−", counts["removed"], "Removed Lines", "removed"),
            ("~", counts["changed"], "Total Changed Lines", "changed"),
            ("!", flag_count, "Review Flags", "flags"),
        ])

        render_comments_panel("Engineering Comments", comments)
        render_diff_panel(original_text, updated_text, "Original SOO", "Updated SOO")


# ============================================================
# Points List Review
# ============================================================
elif mode == "Points List Review":
    col1, col2 = st.columns(2, gap="large")
    with col1:
        original_points_file = file_uploader_card("Original Points List", "Upload original PDF table, Excel, or CSV.", "points_original", ["pdf", "xlsx", "xls", "csv"])
    with col2:
        updated_points_file = file_uploader_card("Updated Points List", "Upload updated PDF table, Excel, or CSV.", "points_updated", ["pdf", "xlsx", "xls", "csv"])

    if original_points_file and updated_points_file:
        original_df = read_points_table(original_points_file)
        updated_df = read_points_table(updated_points_file)
        comparison = compare_points(original_df, updated_df)

        if not comparison["summary"].empty:
            s = comparison["summary"].set_index("Metric")["Count"].to_dict()
            render_metric_row([
                ("+", s.get("Added points", 0), "Added Points", "added"),
                ("−", s.get("Removed points", 0), "Removed Points", "removed"),
                ("~", s.get("Changed fields", 0), "Changed Fields", "changed"),
                ("#", s.get("Updated points", 0), "Updated Points", "flags"),
            ])

        render_comments_panel("Points List Review Comments", comparison["comments"])

        st.markdown('<div class="review-panel"><div class="panel-title">Points List Comparison Tables</div>', unsafe_allow_html=True)
        tab1, tab2, tab3, tab4 = st.tabs(["Added Points", "Removed Points", "Changed Fields", "Extracted Tables"])
        with tab1:
            st.dataframe(comparison["added"], use_container_width=True)
        with tab2:
            st.dataframe(comparison["removed"], use_container_width=True)
        with tab3:
            st.dataframe(comparison["changed"], use_container_width=True)
        with tab4:
            st.markdown("#### Original Extract")
            st.dataframe(original_df, use_container_width=True)
            st.markdown("#### Updated Extract")
            st.dataframe(updated_df, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# Wiring Diagram Review
# ============================================================
elif mode == "Wiring Diagram Review":
    st.info("For AutoCAD files, export the DWG to PDF first when possible. Direct DWG parsing is not reliable in Streamlit without a CAD conversion service.")

    col1, col2 = st.columns(2, gap="large")
    with col1:
        original_wiring_file = file_uploader_card("Original Wiring Diagram", "Upload original wiring PDF or text export.", "wiring_original", ["pdf", "txt"])
    with col2:
        updated_wiring_file = file_uploader_card("Updated Wiring Diagram", "Upload updated wiring PDF or text export. DWG should be exported to PDF first.", "wiring_updated", ["pdf", "txt", "dwg", "dxf"])

    if original_wiring_file and updated_wiring_file:
        unsupported = []
        if original_wiring_file.name.lower().endswith((".dwg", ".dxf")):
            unsupported.append(original_wiring_file.name)
        if updated_wiring_file.name.lower().endswith((".dwg", ".dxf")):
            unsupported.append(updated_wiring_file.name)

        if unsupported:
            st.warning("DWG/DXF files were uploaded. This version can store the upload, but cannot reliably extract CAD geometry/text yet. Export those drawings to PDF for comparison.")
        else:
            original_text = read_text_file(original_wiring_file)
            updated_text = read_text_file(updated_wiring_file)

            summary_df, comments = compare_wiring(original_text, updated_text)
            flag_count = sum(1 for c in comments if c["priority"] in ["High", "Review"])

            render_metric_row([
                ("#", int(summary_df["Original Count"].sum()), "Original References", "flags"),
                ("#", int(summary_df["Updated Count"].sum()), "Updated References", "flags"),
                ("+", int(summary_df["Added"].sum()), "Added References", "added"),
                ("−", int(summary_df["Removed"].sum()), "Removed References", "removed"),
            ])

            render_comments_panel("Wiring Diagram Review Comments", comments)

            st.markdown('<div class="review-panel"><div class="panel-title">Extracted Wiring Entity Summary</div>', unsafe_allow_html=True)
            st.dataframe(summary_df, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

            render_diff_panel(original_text, updated_text, "Original Wiring", "Updated Wiring")


# ============================================================
# Full Package Review
# ============================================================
elif mode == "Full Package Review":
    st.markdown("Upload any files you have. The app will run available reviews and summarize package-level concerns.")

    c1, c2 = st.columns(2, gap="large")
    with c1:
        original_soo = file_uploader_card("Original SOO", "Markdown, Word, text, or PDF.", "pkg_original_soo", ["md", "txt", "docx", "pdf"])
        original_points = file_uploader_card("Original Points List", "PDF table, Excel, or CSV.", "pkg_original_points", ["pdf", "xlsx", "xls", "csv"])
        original_wiring = file_uploader_card("Original Wiring", "PDF or text export.", "pkg_original_wiring", ["pdf", "txt"])
    with c2:
        updated_soo = file_uploader_card("Updated SOO", "Markdown, Word, text, or PDF.", "pkg_updated_soo", ["md", "txt", "docx", "pdf"])
        updated_points = file_uploader_card("Updated Points List", "PDF table, Excel, or CSV.", "pkg_updated_points", ["pdf", "xlsx", "xls", "csv"])
        updated_wiring = file_uploader_card("Updated Wiring", "PDF or text export. Export DWG to PDF first.", "pkg_updated_wiring", ["pdf", "txt", "dwg", "dxf"])

    package_comments = []
    metrics = []

    if original_soo and updated_soo:
        ot = read_text_file(original_soo)
        ut = read_text_file(updated_soo)
        counts = get_diff_counts(ot, ut)
        comments = make_engineering_comments(ot, ut)
        package_comments.extend([{"priority": c["priority"], "category": "SOO - " + c["category"], "comment": c["comment"], "items": c["items"]} for c in comments])
        metrics.append(("SOO changed lines", counts["changed"]))

    if original_points and updated_points:
        odf = read_points_table(original_points)
        udf = read_points_table(updated_points)
        comparison = compare_points(odf, udf)
        package_comments.extend([{"priority": c["priority"], "category": "Points - " + c["category"], "comment": c["comment"], "items": c["items"]} for c in comparison["comments"]])
        if not comparison["summary"].empty:
            s = comparison["summary"].set_index("Metric")["Count"].to_dict()
            metrics.append(("Points added", s.get("Added points", 0)))
            metrics.append(("Points removed", s.get("Removed points", 0)))
            metrics.append(("Point field changes", s.get("Changed fields", 0)))

    if original_wiring and updated_wiring and not updated_wiring.name.lower().endswith((".dwg", ".dxf")):
        ow = read_text_file(original_wiring)
        uw = read_text_file(updated_wiring)
        summary_df, comments = compare_wiring(ow, uw)
        package_comments.extend([{"priority": c["priority"], "category": "Wiring - " + c["category"], "comment": c["comment"], "items": c["items"]} for c in comments])
        metrics.append(("Wiring added refs", int(summary_df["Added"].sum())))
        metrics.append(("Wiring removed refs", int(summary_df["Removed"].sum())))

    if package_comments or metrics:
        metric_cards = []
        for i, (label, value) in enumerate(metrics[:4]):
            metric_cards.append(("#", value, label, "flags"))
        while len(metric_cards) < 4:
            metric_cards.append(("—", 0, "Waiting for files", "flags"))

        render_metric_row(metric_cards)
        render_comments_panel("Full Package Review Comments", package_comments or [{"priority": "Info", "category": "No review yet", "comment": "Upload matching original and updated files to run package review.", "items": []}])
