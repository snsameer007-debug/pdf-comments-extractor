import streamlit as st
import fitz
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

st.set_page_config(
    page_title="PDF Comment Extractor",
    layout="wide"
)

st.title("PDF Comment Extractor")

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

    ignore_types = [
        "Stamp",
        "Square"
    ]

    ignore_authors = [
        "AutoCAD SHX Text"
    ]

    all_items = []

    for page_num in range(len(doc)):

        page = doc[page_num]

        annot = page.first_annot

        while annot:

            try:

                annotation_type = annot.type[1]

                if annotation_type in ignore_types:

                    annot = annot.next
                    continue

                info = annot.info or {}

                author = str(
                    info.get("title", "")
                ).strip()

                if author in ignore_authors:

                    annot = annot.next
                    continue

                date = str(
                    info.get("creationDate", "")
                ).strip()

                subject = str(
                    info.get("subject", "")
                ).strip()

                content = str(
                    info.get("content", "")
                ).strip()

                comment = ""

                for value in [
                    content,
                    subject
                ]:

                    if value:

                        comment = value
                        break

                if not comment:

                    comment = str(info)

                parent_xref = None

                try:

                    parent_xref = annot.irt_xref

                except Exception:

                    parent_xref = None

                all_items.append(
                    {
                        "xref": annot.xref,
                        "reply_to": parent_xref,
                        "page": page_num + 1,
                        "author": author,
                        "type": annotation_type,
                        "subject": subject,
                        "comment": comment,
                        "date": date
                    }
                )
                      except Exception as ex:

                all_items.append(
                    {
                        "xref": "",
                        "reply_to": None,
                        "page": page_num + 1,
                        "author": "",
                        "type": "ERROR",
                        "subject": "",
                        "comment": str(ex),
                        "date": ""
                    }
                )

            annot = annot.next
