import streamlit as st
import pandas as pd
import math
import io
import zipfile

st.set_page_config(page_title="Batch File Splitter", layout="centered")

st.title("📂 Excel / CSV Batch Splitter")
st.write("Upload a file and split it into multiple batches based on number of rows.")

uploaded_file = st.file_uploader(
    "Upload Excel / CSV file",
    type=["csv", "xlsx", "xls"]
)

rows_per_batch = st.number_input(
    "Enter rows per batch",
    min_value=1,
    step=1
)

if uploaded_file is not None and rows_per_batch:

    file_name = uploaded_file.name.lower()

    # Read file
    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)

        elif file_name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")

        elif file_name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")

        else:
            st.error("❌ Unsupported file format")
            st.stop()

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        st.stop()

    # Convert datetime columns to dd-MMM-yyyy format
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%d-%b-%Y')

    total_rows = len(df)
    total_batches = math.ceil(total_rows / rows_per_batch)

    st.success(f"✅ Total Rows: {total_rows}")
    st.info(f"📦 Total Batches: {total_batches}")

    if st.button("🚀 Split File"):

        progress = st.progress(0)

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:

            for i in range(total_batches):

                start = i * rows_per_batch
                end = start + rows_per_batch

                batch_df = df.iloc[start:end]

                output = io.BytesIO()

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    batch_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Sheet1"
                    )

                batch_filename = f"batch_{i+1}.xlsx"

                zf.writestr(
                    batch_filename,
                    output.getvalue()
                )

                progress.progress((i + 1) / total_batches)

        st.success("✅ File successfully split into batches")

        st.download_button(
            label="📥 Download ZIP",
            data=zip_buffer.getvalue(),
            file_name="split_batches.zip",
            mime="application/zip"
        )
