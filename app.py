import streamlit as st
import fitz

st.set_page_config(
    page_title="PDF Debug Tool",
    layout="wide"
)

st.title("PDF Debug Tool")

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

if uploaded_file:

    pdf_data = uploaded_file.read()

    doc = fitz.open(
        stream=pdf_data,
        filetype="pdf"
    )

    st.success(
        f"Pages Found: {len(doc)}"
    )

    total_annotations = 0

    for page_num in range(len(doc)):

        page = doc[page_num]

        annot = page.first_annot

        if annot is None:
            continue

        st.subheader(
            f"Page {page_num + 1}"
        )

        while annot:

            total_annotations += 1

            st.write(
                "Type:",
                annot.type
            )

            try:
                st.write(
                    "Info:",
                    annot.info
                )
            except Exception as e:
                st.write(
                    "Info Error:",
                    str(e)
                )

            st.write("----------------")

            annot = annot.next

    st.success(
        f"Total Annotations Found: {total_annotations}"
    )
