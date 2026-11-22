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

# --------------------------------------------------------
# PAGE
# --------------------------------------------------------

st.set_page_config(
    page_title="PDF Comment Extractor",
    layout="wide"
)

st.title("PDF Comment Extractor")

uploaded_file = st.file_uploader(
    "Upload PDF File",
    type=["pdf"]
)

# --------------------------------------------------------
# SETTINGS
# --------------------------------------------------------

IGNORE_TYPES = {
    "Stamp",
    "Square"
}

IGNORE_AUTHORS = {
    "AutoCAD SHX Text"
}

# --------------------------------------------------------
# HELPERS
# --------------------------------------------------------


def safe_text(value):

    if value is None:
        return ""

    return str(value).replace(
        "\r",
        " "
    ).replace(
        "\n",
        " "
    ).strip()


def get_comment_text(info):

    content = safe_text(
        info.get("content", "")
    )

    subject = safe_text(
        info.get("subject", "")
    )

    if content:
        return content

    if subject:
        return subject

    return ""


# --------------------------------------------------------
# MAIN
# --------------------------------------------------------

if uploaded_file:

    pdf_bytes = uploaded_file.read()

    try:

        doc = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

    except Exception as ex:

        st.error(
            f"Cannot open PDF: {ex}"
        )

        st.stop()

    comments = []

    for page_index in range(len(doc)):

        page = doc[page_index]

        annot = page.first_annot

        while annot:

            next_annot = annot.next

            try:

                annotation_type = annot.type[1]

                if annotation_type in IGNORE_TYPES:

                    annot = next_annot
                    continue

                info = annot.info or {}

                author = safe_text(
                    info.get(
                        "title",
                        ""
                    )
                )

                if author in IGNORE_AUTHORS:

                    annot = next_annot
                    continue

                subject = safe_text(
                    info.get(
                        "subject",
                        ""
                    )
                )

                comment = get_comment_text(
                    info
                )

                creation_date = safe_text(
                    info.get(
                        "creationDate",
                        ""
                    )
                )

                reply_to = None

                try:

                    if annot.irt_xref > 0:

                        reply_to = annot.irt_xref

                except Exception:

                    pass

                comments.append({

                    "xref":
                        annot.xref,

                    "reply_to":
                        reply_to,

                    "page":
                        page_index + 1,

                    "author":
                        author,

                    "type":
                        annotation_type,

                    "subject":
                        subject,

                    "comment":
                        comment,

                    "date":
                        creation_date

                })

            except Exception as ex:

                comments.append({

                    "xref": 0,

                    "reply_to": None,

                    "page":
                        page_index + 1,

                    "author": "",

                    "type": "ERROR",

                    "subject": "",

                    "comment": str(ex),

                    "date": ""

                })

            annot = next_annot

    results = []

    for parent in comments:

        if parent["reply_to"] is not None:

            continue

        replies = []

        reply_authors = []

        for child in comments:

            if child["reply_to"] == parent["xref"]:

                if child["comment"]:

                    replies.append(
                        child["comment"]
                    )

                if child["author"]:

                    reply_authors.append(
                        child["author"]
                    )

        results.append({

            "Page":
                parent["page"],

            "Author":
                parent["author"],

            "Type":
                parent["type"],

            "Subject":
                parent["subject"],

            "Comment":
                parent["comment"],

            "Reply Count":
                len(replies),

            "Responses":
                "\n\n-----------------\n\n".join(
                    replies
                ),

            "Responded By":
                "\n".join(
                    reply_authors
                ),

            "Date":
                parent["date"]

        })

    df = pd.DataFrame(
        results
    )

    st.success(
        f"{len(df)} comments extracted"
    )

    st.dataframe(
        df,
           use_container_width=True
    )

 # ----------------------------------------------------
    # EXCEL
    # ----------------------------------------------------

    excel_buffer = BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Comments"
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
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for row in ws.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )

    for column in ws.columns:

        maximum = 0

        letter = column[0].column_letter

        for cell in column:

            try:

                length = len(
                    str(cell.value)
                )

                if length > maximum:

                    maximum = length

            except:

                pass

        ws.column_dimensions[
            letter
        ].width = min(
            maximum + 5,
            80
        )

    ws.freeze_panes = "A2"

    output = BytesIO()

    wb.save(
        output
    )

    output.seek(0)

    st.download_button(

        label="Download Excel",

        data=output,

        file_name="PDF_Comments.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )
