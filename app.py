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

    IGNORE_TYPES = {
        "Stamp",
        "Square"
    }

    IGNORE_AUTHORS = {
        "AutoCAD SHX Text"
    }

    all_comments = []

    for page_num in range(len(doc)):

        page = doc[page_num]

        annot = page.first_annot

        while annot:

            try:

                annot_type = annot.type[1]

                if annot_type in IGNORE_TYPES:
                    annot = annot.next
                    continue

                info = annot.info or {}

                author = str(
                    info.get("title", "")
                ).strip()

                if author in IGNORE_AUTHORS:
                    annot = annot.next
                    continue

                subject = str(
                    info.get("subject", "")
                ).strip()

                content = str(
                    info.get("content", "")
                ).strip()

                creation_date = str(
                    info.get("creationDate", "")
                ).strip()

                comment_text = ""

                if content:
                    comment_text = content

                elif subject:
                    comment_text = subject

                else:
                    comment_text = str(info)

                parent_xref = None

                try:
                    parent_xref = annot.irt_xref
                except:
                    pass

                all_comments.append({
                    "xref": annot.xref,
                    "reply_to": parent_xref,
                    "page": page_num + 1,
                    "author": author,
                    "type": annot_type,
                    "subject": subject,
                    "comment": comment_text,
                    "date": creation_date
                })

            except Exception as ex:

                all_comments.append({
                    "xref": None,
                    "reply_to": None,
                    "page": page_num + 1,
                    "author": "",
                    "type": "ERROR",
                    "subject": "",
                    "comment": str(ex),
                    "date": ""
                })

            annot = annot.next

    comment_map = {}

    for item in all_comments:

        comment_map[item["xref"]] = item

    results = []

    for item in all_comments:

        if item["reply_to"] is not None:
            continue

        replies = []
        replied_by = []

        for child in all_comments:

            if child["reply_to"] == item["xref"]:

                replies.append(
                    child["comment"]
                )

                replied_by.append(
                    child["author"]
                )

        results.append({

            "Page": item["page"],

            "Author": item["author"],

            "Type": item["type"],

            "Subject": item["subject"],

            "Comment": item["comment"],

            "Response Count": len(replies),

            "Responses":
                "\n\n-------------------\n\n".join(replies),

            "Responded By":
                "\n".join(replied_by),

            "Date": item["date"]

        })

    df = pd.DataFrame(results)

    st.success(
        f"{len(df)} comments extracted"
    )

    st.dataframe(
        df,
        use_container_width=True
    )

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

    wb = load_workbook(excel_buffer)

    ws = wb["Comments"]

    dark_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    for cell in ws[1]:

        cell.fill = dark_fill

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

    for column in ws.columns:

        max_length = 0

        letter = column[0].column_letter

        for cell in column:

            try:

                value = str(cell.value)

                if len(value) > max_length:
                    max_length = len(value)

            except:
                pass

        ws.column_dimensions[
            letter
        ].width = min(
            max_length + 5,
            100
        )

    ws.freeze_panes = "A2"

    final_excel = BytesIO()

    wb.save(final_excel)

    final_excel.seek(0)

    st.download_button(
        label="Download Excel",
        data=final_excel,
        file_name="PDF_Comments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
