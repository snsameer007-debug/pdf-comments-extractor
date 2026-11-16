import streamlit as st
import fitz
import pandas as pd
from io import BytesIO
import base64
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

st.set_page_config(
    page_title="PDF Comments Extractor",
    layout="wide"
)

st.title("PDF Comments Extractor")

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

if uploaded_file:

    pdf_bytes = uploaded_file.read()

    # PDF Preview
    st.subheader("PDF Preview")

    try:
        pdf_base64 = base64.b64encode(pdf_bytes).decode()

        pdf_viewer = f"""
        <iframe
        src="data:application/pdf;base64,{pdf_base64}"
        width="100%"
        height="900">
        </iframe>
        """

        st.markdown(
            pdf_viewer,
            unsafe_allow_html=True
        )

    except Exception:
        st.warning("Unable to preview PDF")

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    IGNORE_TYPES = [
        "Stamp",
        "Square"
    ]

    annotations = []

    annotation_counter = 0

    # First Pass
    for page_no in range(len(doc)):

        page = doc[page_no]

        annot = page.first_annot

        while annot:

            annotation_type = annot.type[1]

            if annotation_type not in IGNORE_TYPES:

                info = annot.info

                annotation_counter += 1

                record = {
                    "internal_id": annotation_counter,
                    "page": page_no + 1,
                    "author": info.get("title", ""),
                    "comment": info.get("content", ""),
                    "type": annotation_type,
                    "date": info.get("creationDate", ""),
                    "xref": annot.xref,
                    "reply_to": None
                }

                try:

                    parent = annot.irt_xref

                    if parent:
                        record["reply_to"] = parent

                except Exception:
                    pass

                annotations.append(record)

            annot = annot.next

    # Build lookup
    xref_map = {}

    for item in annotations:
        xref_map[item["xref"]] = item

    output_rows = []

    for item in annotations:

        if item["reply_to"] is not None:
            continue

        replies = []

        reply_authors = []

        for candidate in annotations:

            if candidate["reply_to"] == item["xref"]:

                if candidate["comment"]:

                    replies.append(
                        candidate["comment"]
                    )

                if candidate["author"]:

                    reply_authors.append(
                        candidate["author"]
                    )

        output_rows.append(

            {
                "Page": item["page"],
                "Author": item["author"],
                "Type": item["type"],
                "Comment": item["comment"],
                "Response":
                    "\n------------------\n".join(replies),
                "Responded By":
                    "\n".join(reply_authors),
                "Date": item["date"]
            }

        )

    st.success(
        f"Extracted {len(output_rows)} primary comments"
    )

    if len(output_rows) > 0:

        df = pd.DataFrame(output_rows)

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
            "solid",
            fgColor="1F4E78"
        )

        for cell in ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = header_fill

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

        for col in ws.columns:

            max_len = 0

            column = col[0].column_letter

            for cell in col:

                try:

                    value_len = len(
                        str(cell.value)
                    )

                    if value_len > max_len:
                        max_len = value_len

                except Exception:
                    pass

            ws.column_dimensions[
                column
            ].width = min(
                max_len + 5,
                70
            )

        ws.freeze_panes = "A2"

        final_excel = BytesIO()

        wb.save(final_excel)

        final_excel.seek(0)

        st.download_button(
            "Download Excel",
            data=final_excel,
            file_name="PDF_Comments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(
            df,
            use_container_width=True
        )

    else:

        st.warning(
            "No comments found."
        )
