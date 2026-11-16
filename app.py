import streamlit as st
import fitz
import pandas as pd
from io import BytesIO

st.title("PDF Comments Extractor")

pdf_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

if pdf_file:

    pdf = fitz.open(
        stream=pdf_file.read(),
        filetype="pdf"
    )

    comments = []

    for page_num in range(len(pdf)):

        page = pdf[page_num]

        annot = page.first_annot

        while annot:

            info = annot.info

            comments.append({
                "Page": page_num + 1,
                "Author": info.get("title", ""),
                "Comment": info.get("content", ""),
                "Type": annot.type[1]
            })

            annot = annot.next

    if comments:

        df = pd.DataFrame(comments)

        output = BytesIO()

        df.to_excel(
            output,
            index=False
        )

        st.success(
            f"Found {len(comments)} comments"
        )

        st.download_button(
            "Download Excel",
            output.getvalue(),
            "PDF_Comments.xlsx"
        )
    else:
        st.warning("No comments found")
