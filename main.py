import streamlit as st
import pandas as pd
import math
import io
import zipfile

st.set_page_config(page_title="File Batch Splitter", layout="centered")

st.title("📂 Excel / CSV Batch Splitter")

st.write("Upload a file and split it into batches based on number of rows.")

# Upload file
uploaded_file = st.file_uploader(
    "Upload Excel or CSV file",
    type=["xlsx", "csv"]
)

rows_per_batch = st.number_input(
    "Enter number of rows per batch",
    min_value=1,
    step=1
)

if uploaded_file is not None and rows_per_batch:

    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    total_rows = len(df)
    total_batches = math.ceil(total_rows / rows_per_batch)

    st.success(f"Total Rows: {total_rows}")
    st.info(f"Total Batches Created: {total_batches}")

    if st.button("🚀 Split and Download"):

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:

            for i in range(total_batches):
                start = i * rows_per_batch
                end = start + rows_per_batch

                batch_df = df.iloc[start:end]

                output = io.BytesIO()

                file_name = f"batch_{i+1}"

                if uploaded_file.name.endswith(".csv"):
                    batch_df.to_csv(output, index=False)
                    zf.writestr(f"{file_name}.csv", output.getvalue())
                else:
                    batch_df.to_excel(output, index=False)
                    zf.writestr(f"{file_name}.xlsx", output.getvalue())

        st.download_button(
            label="📥 Download All Batches (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="batches.zip",
            mime="application/zip"
        )
