import io
import streamlit as st
import openpyxl

def acenet_image_url(sku: str) -> str:
    return f"https://images.aceservices.com/{sku[:2]}/{sku[2:4]}/{sku}_A.eps_High.jpg"

def process_file(input_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(input_bytes))
    ws = wb.active
    ws.sheet_format.defaultRowHeight = 60
    count = 0
    for row in range(4, ws.max_row + 1):
        sku = ws.cell(row=row, column=11).value  # Column K = Article No
        if sku is not None:
            url = acenet_image_url(str(sku))
            ws.cell(row=row, column=10).value = f'=_xlfn.IMAGE("{url}")'
            count += 1
    if count == 0:
        raise ValueError(
            "No SKUs were found in column K starting at row 4. "
            "Please make sure you are uploading the correct WTM Excel file with Article Numbers in column K."
        )
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, count


st.set_page_config(page_title="WTM Image URL Generator", layout="centered")
st.title("WTM Image URL Generator")
st.write("Upload a WTM Excel file to populate Article Image URLs.")
st.info("**Expected format:** Article Numbers (SKUs) should be in **column K**, starting at row 4.")

uploaded_file = st.file_uploader("Select WTM Excel File", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Generate"):
        try:
            result, count = process_file(uploaded_file.getvalue())
            base_name = uploaded_file.name.rsplit(".", 1)[0]
            st.success(f"Populated {count} image URLs.")
            st.download_button(
                label="Download Result",
                data=result,
                file_name=f"{base_name} - with images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.warning(
                "**After downloading:** When you open the file, Excel may block active content. "
                "Click **Enable Content** (or **Enable Editing**) when prompted. "
                "If images do not appear, close the file and reopen it."
            )
        except Exception as e:
            st.error(str(e))
