import streamlit as st
import fitz
import pandas as pd
from io import BytesIO
from PIL import Image
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

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    st.subheader("PDF Preview")

    page_to_show = st.number_input(
        "Page Number",
        min_value=1,
        max_value=len(doc),
        value=1
    )

    page = doc[page_to_show - 1]

    pix = page.get_pixmap(
        matrix=fitz.Matrix(2, 2)
    )

    img = Image.frombytes(
        "RGB",
        [pix.width, pix.height],
        pix.samples
    )

    st.image(
        img,
        caption=f"Page {page_to_show}",
        use_container_width=True
    )

    ignore_types = [
        "Stamp",
        "Square"
    ]

    all_annotations = []

    for page_no in range(len(doc)):

        page = doc[page_no]

        annot = page.first_annot

        while annot:

            try:

                annotation_type = annot.type[1]

                if annotation_type in ignore_types:
                    annot = annot.next
                    continue

                info = annot.info

                author = info.get(
                    "title",
                    ""
                )

                comment = info.get(
                    "content",
                    ""
                )

                date = info.get(
                    "creationDate",
                    ""
                )

                parent_xref = None

                try:

                    parent_xref = annot.irt_xref

                except:

                    parent_xref = None

                all_annotations.append(
                    {
                        "xref": annot.xref,
                        "reply_to": parent_xref,
                        "Page": page_no + 1,
                        "Author": author,
                        "Type": annotation_type,
                        "Comment": comment,
                        "Date": date
                    }
                )

            except Exception as ex:

                st.warning(str(ex))

            annot = annot.next

    final_rows = []

    for item in all_annotations:

        if item["reply_to"] is not None:
            continue

        replies = []

        repliers = []

        for candidate in all_annotations:

            if candidate["reply_to"] == item["xref"]:

                replies.append(
                    candidate["Comment"]
                )

                repliers.append(
                    candidate["Author"]
                )

        final_rows.append(
            {
                "Page": item["Page"],
                "Author": item["Author"],
                "Type": item["Type"],
                "Comment": item["Comment"],
                "Response": "\n\n".join(replies),
                "Responded By": "\n".join(repliers),
                "Date": item["Date"]
            }
        )

    df = pd.DataFrame(final_rows)

    st.success(
        f"Found {len(df)} comments"
    )

    st.dataframe(
        df,
        use_container_width=True
    )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Comments"
        )

    output.seek(0)

    wb = load_workbook(output)

    ws = wb["Comments"]

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = header_fill

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

        max_length = 0

        column = col[0].column_letter

        for cell in col:

            try:

                length = len(
                    str(cell.value)
                )

                if length > max_length:
                    max_length = length

            except:
                pass

        ws.column_dimensions[
            column
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
