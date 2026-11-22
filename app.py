import streamlit as st
import fitz
import pandas as pd

from io import BytesIO

from openpyxl import load_workbook

from openpyxl.styles import (
    Alignment,
    Font,
    PatternFill
)

st.set_page_config(
    page_title="PDF Comment Extractor",
    layout="wide"
)

st.title("PDF Comment Extractor")

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------

IGNORE_TYPES = {
    "Stamp",
    "Square"
}

IGNORE_AUTHORS = {
    "AutoCAD SHX Text"
}


def clean_value(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\r", " ")
    value = value.replace("\n", " ")

    return value.strip()


def get_raw_pdf_object(doc, xref):

    try:

        return doc.xref_object(
            xref,
            compressed=False
        )

    except Exception:

        return ""


def extract_comment_text(info, raw_object):

    candidates = []

    candidates.append(
        clean_value(
            info.get("content", "")
        )
    )

    candidates.append(
        clean_value(
            info.get("subject", "")
        )
    )

    candidates.append(
        clean_value(
            info.get("title", "")
        )
    )

    for value in candidates:

        if value:
            return value

    if raw_object:

        for marker in [
            "/Contents",
            "/Subj"
        ]:

            pos = raw_object.find(marker)

            if pos >= 0:

                snippet = raw_object[
                    pos:
                    pos + 500
                ]

                return snippet

    return ""


# --------------------------------------------------
# Main
# --------------------------------------------------

if uploaded_file:

    pdf_bytes = uploaded_file.read()

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    all_annotations = []

    for page_number in range(len(doc)):

        page = doc[page_number]

        annot = page.first_annot

        while annot:

            try:

                annotation_type = annot.type[1]

                if annotation_type in IGNORE_TYPES:

                    annot = annot.next
                    continue

                info = annot.info or {}

                author = clean_value(
                    info.get("title", "")
                )

                if author in IGNORE_AUTHORS:

                    annot = annot.next
                    continue

                raw_object = get_raw_pdf_object(
                    doc,
                    annot.xref
                )

                parent_xref = None

                try:

                    parent_xref = annot.irt_xref

                except Exception:

                    parent_xref = None

                subject = clean_value(
                    info.get(
                        "subject",
                        ""
                    )
                )

                creation_date = clean_value(
                    info.get(
                        "creationDate",
                        ""
                    )
                )

                comment_text = extract_comment_text(
                    info,
                    raw_object
                )

                all_annotations.append({

                    "xref": annot.xref,

                    "reply_to": parent_xref,

                    "page": page_number + 1,

                    "author": author,

                    "type": annotation_type,

                    "subject": subject,

                    "comment": comment_text,

                    "raw_object": raw_object,

                    "date": creation_date

                })

            except Exception as ex:

                all_annotations.append({

                    "xref": None,

                    "reply_to": None,

                    "page": page_number + 1,

                    "author": "",

                    "type": "ERROR",

                    "subject": "",

                    "comment": str(ex),

                    "raw_object": "",

                    "date": ""

                })

            annot = annot.next

    # --------------------------------------
    # Build reply chains
    # --------------------------------------

    results = []

    for item in all_annotations:

        if item["reply_to"] is not None:
            continue

        reply_comments = []

        reply_authors = []

        reply_dates = []

        for child in all_annotations:

            if child["reply_to"] == item["xref"]:

                if child["comment"]:

                    reply_comments.append(
                        child["comment"]
                    )

                if child["author"]:

                    reply_authors.append(
                        child["author"]
                    )

                if child["date"]:

                    reply_dates.append(
                        child["date"]
                    )

        results.append({

            "Page":
                item["page"],

            "Author":
                item["author"],

            "Type":
                item["type"],

            "Subject":
                item["subject"],

            "Comment":
                item["comment"],

            "Response Count":
                len(reply_comments),

            "Responses":
                "\n\n====================\n\n".join(
                    reply_comments
                ),

            "Responded By":
                "\n".join(
                    reply_authors
                ),

            "Response Dates":
                "\n".join(
                    reply_dates
                ),

            "Date":
                item["date"],

            "XREF":
                item["xref"]

        })

    df = pd.DataFrame(results)

    st.success(
        f"{len(df)} comments extracted"
    )

    st.dataframe(
        df,
        use_container_width=True
    )

    # --------------------------------------
    # Export Excel
    # --------------------------------------

    excel_buffer = BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Comments",
            index=False
        )

    excel_buffer.seek(0)

    wb = load_workbook(
        excel_buffer
    )

    ws = wb["Comments"]

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.alignment = Alignment(
            wrap_text=True,
            horizontal="center",
            vertical="center"
        )

    for row in ws.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )

    for col in ws.columns:

        max_len = 0

        col_letter = (
            col[0].column_letter
        )

        for cell in col:

            try:

                length = len(
                    str(cell.value)
                )

                if length > max_len:

                    max_len = length

            except Exception:

                pass

        ws.column_dimensions[
            col_letter
        ].width = min(
            max_len + 5,
            100
        )

    ws.freeze_panes = "A2"

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    st.download_button(
        label="Download Excel",
        data=output,
        file_name="PDF_Comments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
