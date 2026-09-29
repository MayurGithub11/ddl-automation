"""
Pending New Service Connection (LT) list - Streamlit feature
Upload LIST.csv / .zip / .xlsx  ->  choose Division + Sub-division  ->  download clean Excel.
"""
import io
import pandas as pd
import streamlit as st

# Columns to keep, in output order (Division & Sub-dn come first)
KEEP_COLS = [
    "Division", "Sub-dn", "Section", "DTC", "Category", "Workflow Status",
    "Application ID", "Consumer Number", "Name", "Address", "Infra Status",
    "Report Age", "SOP Remark", "Age Days",
]
RENAME = {"Sub-dn": "Sub-Division"}
DEFAULT_DIVISION_KEY = "BHANDARA"
DEFAULT_SUBDN_KEY = "PAONI"


@st.cache_data(show_spinner="Reading file...")
def load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Read CSV / ZIP(of CSV) / Excel. Loads ONLY the needed columns to save memory."""
    name = filename.lower()
    need = ["Sr.No."] + KEEP_COLS
    if name.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
        hdr = _find_header(raw)
        df = raw.iloc[hdr + 1:].copy()
        df.columns = [str(c).strip() for c in raw.iloc[hdr]]
    else:
        comp = "zip" if name.endswith(".zip") else None
        head = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, nrows=15,
                           compression=comp, encoding_errors="replace")
        hdr = _find_header(head)
        df = pd.read_csv(io.BytesIO(file_bytes), header=hdr, dtype=str, low_memory=False,
                         compression=comp, encoding_errors="replace",
                         usecols=lambda c: str(c).strip() in need)
        df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in file: {missing}")
    df = df[df["Sr.No."].notna()]                       # drop blank/footer rows
    df = df[df["Sr.No."].str.strip() != "Sr.No."]       # drop repeated headers
    return df[need].reset_index(drop=True)


def _find_header(raw: pd.DataFrame) -> int:
    for i in range(min(15, len(raw))):
        if raw.iloc[i].astype(str).str.strip().eq("Sr.No.").any():
            return i
    raise ValueError("Could not find the header row (a row containing 'Sr.No.').")


def clean(df: pd.DataFrame) -> pd.DataFrame:
    out = df[KEEP_COLS].copy()
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip().replace({"nan": "", "-": ""})
    # numeric fields: strip thousands commas
    out["Age Days"] = pd.to_numeric(out["Age Days"].str.replace(",", ""), errors="coerce")
    out["Application ID"] = out["Application ID"].str.replace(",", "")
    out = out.rename(columns=RENAME)
    return out


def remove_ag(df: pd.DataFrame) -> pd.DataFrame:
    return df[~df["Category"].str.upper().str.contains(r"\bAG\b", regex=True)]


def to_excel(df: pd.DataFrame, by_section: bool) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        wb = xw.book
        hfmt = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                              "border": 1, "text_wrap": True, "valign": "vcenter",
                              "align": "center"})
        cfmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
        widths = {"Division": 20, "Sub-Division": 16, "Section": 18, "DTC": 10,
                  "Category": 12, "Workflow Status": 26, "Application ID": 14,
                  "Consumer Number": 16, "Name": 26, "Address": 40,
                  "Infra Status": 20, "Report Age": 17, "SOP Remark": 15, "Age Days": 9}

        def write(sheet, data):
            data = data.reset_index(drop=True)
            data.insert(0, "Sr No", range(1, len(data) + 1))
            data.to_excel(xw, sheet_name=sheet, index=False, startrow=0, header=False)
            ws = xw.sheets[sheet]
            for j, col in enumerate(data.columns):
                ws.write(0, j, col, hfmt)
                ws.set_column(j, j, 6 if col == "Sr No" else widths.get(col, 14), cfmt)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(data), len(data.columns) - 1)
            ws.set_landscape(); ws.set_paper(9)          # A4 landscape
            ws.fit_to_pages(1, 0); ws.repeat_rows(0)     # 1 page wide, repeat header

        write("All", df)
        if by_section:
            for sec, g in df.groupby("Section"):
                name = (sec or "No Section")[:28]
                for ch in '[]:*?/\\':
                    name = name.replace(ch, "")
                write(name or "Sheet", g)
    return bio.getvalue()


def main():
    st.title("Pending New Connections - Report")
    up = st.file_uploader("Upload LIST file (CSV, ZIP of CSV, or Excel)", type=["csv", "zip", "xlsx", "xls"])
    if not up:
        return
    try:
        df = load_file(up.getvalue(), up.name)
    except Exception as e:
        st.error(str(e))
        return

    # --- Division / Sub-division selection (defaults: Bhandara / Paoni) ---
    divs = sorted(df["Division"].dropna().unique())
    d_idx = next((i for i, d in enumerate(divs) if DEFAULT_DIVISION_KEY in d.upper()), 0)
    division = st.selectbox("Division", divs, index=d_idx)

    subs = sorted(df.loc[df["Division"] == division, "Sub-dn"].dropna().unique())
    s_idx = next((i for i, s in enumerate(subs) if DEFAULT_SUBDN_KEY in s.upper()), 0)
    sub = st.selectbox("Sub-Division", subs, index=s_idx)

    sel = df[(df["Division"] == division) & (df["Sub-dn"] == sub)]
    out = remove_ag(clean(sel))
    removed = len(sel) - len(out)
    out = out.sort_values(["Section", "Age Days"], ascending=[True, False])

    c1, c2, c3 = st.columns(3)
    c1.metric("Applications", len(out))
    c2.metric("AG removed", removed)
    c3.metric("Beyond SOP", int(out["SOP Remark"].str.contains("BEYOND").sum()))

    by_section = st.checkbox("Also make a separate sheet for each Section", value=True)
    st.dataframe(out, use_container_width=True, hide_index=True)

    st.download_button(
        "Download Excel",
        data=to_excel(out, by_section),
        file_name=f"Pending_NSC_{sub.replace('/', '').replace('.', '').strip()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
