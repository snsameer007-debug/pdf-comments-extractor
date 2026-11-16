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

                for value in [content, subject]:

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

    results = []

    for item in all_items:

        if item["reply_to"] is not None:
            continue

        replies = []
        reply_authors = []

        for child in all_items:

            if child["reply_to"] == item["xref"]:

                if child["comment"]:
                    replies.append(
                        child["comment"]
                    )

                if child["author"]:
                    reply_authors.append(
                        child["author"]
                    )

        results.append(
            {
                "Page": item["page"],
                "Author": item["author"],
                "Type": item["type"],
                "Subject": item["subject"],
                "Comment": item["comment"],
                "Response": "\n\n".join(replies),
                "Responded By": "\n".join(reply_authors),
                "Date": item["date"]
            }
        )

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

    for column in ws.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            value = str(cell.value)

            if len(value) > max_length:
                max_length = len(value)

        ws.column_dimensions[
            column_letter
        ].width = min(
            max_length + 5,
            80
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
