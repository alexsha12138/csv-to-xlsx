# streamlit_app.py
import io
import pandas as pd
import streamlit as st
from charset_normalizer import from_bytes

st.title("CSV → XLSX (Row Range)")

uploaded = st.file_uploader("Upload CSV", type=["csv"])

def detect_encoding(file_obj) -> str:
    raw = file_obj.read()
    result = from_bytes(raw).best()
    encoding = result.encoding if result else None
    # Reset pointer for next read
    file_obj.seek(0)
    return encoding

if uploaded:
    # Detect encoding
    detected = detect_encoding(uploaded)
    st.caption(f"Detected encoding: **{detected or 'unknown'}**")

    # Let user override if needed
    enc_options = ["Auto-detect", "utf-8", "utf-8-sig", "utf-16", "gb18030", "big5", "windows-1252", "latin1"]
    enc_choice = st.selectbox("Encoding", enc_options, index=0)

    # Read CSV with fallback strategy
    tried = []
    df = None

    def try_read(encoding):
        uploaded.seek(0)
        return pd.read_csv(uploaded, sep="\t", encoding=encoding)

    if enc_choice != "Auto-detect" and enc_choice is not None:
        try:
            df = try_read(enc_choice)
        except Exception as e:
            tried.append((enc_choice, str(e)))

    if df is None:
        # Build ordered fallback list
        fallbacks = []
        if detected:
            fallbacks.append(detected)
        fallbacks += ["utf-8", "utf-8-sig", "utf-16", "gb18030", "big5"]
        # De-dup while preserving order
        seen = set()
        fallbacks = [e for e in fallbacks if not (e in seen or seen.add(e))]

        for enc in fallbacks:
            try:
                df = try_read(enc)
                st.caption(f"Loaded with encoding: **{enc}**")
                break
            except Exception as e:
                tried.append((enc, str(e)))

    if df is None:
        st.error("Could not read the CSV with the attempted encodings.")
        with st.expander("Show errors"):
            for enc, err in tried:
                st.write(f"**{enc}** → {err}")
    else:
        st.write(f"Rows: {len(df):,} | Columns: {len(df.columns):,}")
        st.dataframe(df, use_container_width=True)

        start = st.number_input("Start row (1-based)", min_value=1, max_value=max(1, len(df)), value=1, step=1)
        end = st.number_input("End row (1-based)", min_value=1, max_value=max(1, len(df)), value=len(df), step=1)

        # Custom filename input
        file_name = st.text_input("Output file name", value="converted_file.xlsx")

        if start > end:
            st.error("Start row must be ≤ End row.")
        else:
            if st.button("Convert to XLSX"):
                selection = df.iloc[start-1:end]
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    selection.to_excel(writer, index=False, sheet_name="Selection")

                # Ensure .xlsx extension
                final_name = file_name.strip()
                if final_name and not final_name.lower().endswith(".xlsx"):
                    final_name += ".xlsx"
                elif not final_name:
                    final_name = "converted_file.xlsx"

                st.download_button(
                    "Download Excel",
                    data=buf.getvalue(),
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
