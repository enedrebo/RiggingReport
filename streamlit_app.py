import streamlit as st
import zipfile
import io
import hashlib
import tempfile
import os
from pathlib import Path

from excel_reader import read_excel, read_page_list
from report_template import build_pdf

st.set_page_config(page_title="Rigging Report Generator", layout="wide")
st.title("Rigging Calculations Report Generator")

# Initialise session state defaults — must happen before anything else
if "uploader_versions" not in st.session_state:
    st.session_state.uploader_versions = {}   # {page_index: int}
if "page_images" not in st.session_state:
    st.session_state.page_images = []
if "page_list" not in st.session_state:
    st.session_state.page_list = []

# ── Description ───────────────────────────────────────────────────────────────
st.markdown(
    "This tool generates a structured PDF report from a rigging calculations Excel workbook. "
    "Fill in the template with your project data, upload it here (optionally together with "
    "images in a ZIP), and download a ready-to-use report."
)
st.divider()

# ── Template download ──────────────────────────────────────────────────────────
TEMPLATE_FILE = os.path.join(os.path.dirname(__file__), "RiggingCalculations_template.xlsx")
if os.path.isfile(TEMPLATE_FILE):
    with open(TEMPLATE_FILE, "rb") as f:
        st.download_button(
            label="Download Excel template",
            data=f.read(),
            file_name="RiggingCalculations_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
st.divider()

# ── Step 1: Upload ─────────────────────────────────────────────────────────────
st.subheader("1. Upload file")
uploaded = st.file_uploader(
    "Upload your Excel file (.xlsx) or a ZIP containing the Excel and images (.zip)",
    type=["xlsx", "zip"],
)

if not uploaded:
    st.info("Upload an Excel or ZIP file to get started.")
    st.stop()

# ── Parse upload — reset session state when a new file is uploaded ─────────────
file_bytes = uploaded.getvalue()
file_hash  = hashlib.md5(file_bytes).hexdigest()

if st.session_state.get("file_hash") != file_hash:
    excel_bytes = None
    zip_images  = {}

    if uploaded.name.lower().endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            for name in zf.namelist():
                basename = Path(name).name
                if basename.lower().endswith(".xlsx") and not basename.startswith("~"):
                    if excel_bytes is None:
                        excel_bytes = zf.read(name)
                elif basename.lower().endswith((".jpg", ".jpeg", ".png")):
                    zip_images[basename] = zf.read(name)
    else:
        excel_bytes = file_bytes

    if excel_bytes is None:
        st.error("No Excel file found in the ZIP.")
        st.stop()

    try:
        page_list = read_page_list(io.BytesIO(excel_bytes))
    except Exception as e:
        st.error(f"Could not read Excel: {e}")
        st.stop()

    # Clear any old image uploader widget states from a previous session
    for key in list(st.session_state.keys()):
        if key.startswith("img_") and "_v" in key:
            del st.session_state[key]

    st.session_state.file_hash         = file_hash
    st.session_state.excel_bytes       = excel_bytes
    st.session_state.page_list         = page_list
    st.session_state.page_images       = [zip_images.get(p["image_filename"]) for p in page_list]
    st.session_state.uploader_versions = {}


# ── Image uploader callback — fires immediately when a new image is selected ───
def _on_image_upload(i):
    v   = st.session_state.uploader_versions.get(i, 0)
    val = st.session_state.get(f"img_{i}_v{v}")
    if val is not None:
        st.session_state.page_images[i] = val.getvalue()


# ── Step 2: Image management ───────────────────────────────────────────────────
st.subheader("2. Images")
st.caption(
    "Images found in your ZIP are shown automatically. "
    "Use the upload buttons to add or replace images. "
    "Click Remove to clear an image."
)

hcols = st.columns([4, 2, 2, 3])
hcols[0].markdown("**Page**")
hcols[1].markdown("**Image filename**")
hcols[2].markdown("**Preview**")
hcols[3].markdown("**Upload / replace**")
st.divider()

for i, page in enumerate(st.session_state.page_list):
    cols = st.columns([4, 2, 2, 3])

    with cols[0]:
        st.write(page["title"])

    with cols[1]:
        fname = page["image_filename"]
        st.write(fname if fname else "—")

    with cols[2]:
        img_bytes = st.session_state.page_images[i]
        if img_bytes:
            st.image(img_bytes, width=100)
            if st.button("Remove", key=f"remove_{i}"):
                st.session_state.page_images[i] = None
                st.session_state.uploader_versions[i] = (
                    st.session_state.uploader_versions.get(i, 0) + 1
                )
                st.rerun()
        else:
            st.caption("No image")

    with cols[3]:
        v = st.session_state.uploader_versions.get(i, 0)
        st.file_uploader(
            "upload",
            type=["jpg", "jpeg", "png"],
            key=f"img_{i}_v{v}",
            label_visibility="collapsed",
            on_change=_on_image_upload,
            args=(i,),
        )

    st.divider()

# ── Summary table ─────────────────────────────────────────────────────────────
item_pages = [p for p in st.session_state.page_list if "item_no" in p]
if item_pages:
    st.subheader("Summary table")
    hdr = st.columns([2, 4, 4, 2])
    hdr[0].markdown("**Item number**")
    hdr[1].markdown("**Item description**")
    hdr[2].markdown("**Specification**")
    hdr[3].markdown("**Utilization**")
    st.divider()
    for p in item_pages:
        row = st.columns([2, 4, 4, 2])
        row[0].write(p.get("item_no", ""))
        row[1].write(p.get("item_desc", ""))
        row[2].write(p.get("det_desc", ""))
        uf = p.get("uf")
        row[3].write(f"{uf:.5g}" if uf is not None else "—")
    st.divider()

# ── Step 3: Generate ───────────────────────────────────────────────────────────
st.subheader("3. Generate report")

if st.button("Generate PDF", type="primary"):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(st.session_state.excel_bytes)
        tmp_path = f.name

    try:
        with st.spinner("Reading Excel..."):
            loads_data, items_data = read_excel(tmp_path)
    except ValueError as e:
        st.error(f"Error reading Excel: {e}")
        st.stop()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    with st.spinner("Building PDF..."):
        pdf_buffer = io.BytesIO()
        build_pdf(loads_data, items_data, pdf_buffer,
                  page_images=st.session_state.page_images)

    st.success("Report ready!")
    st.download_button(
        label="Download PDF",
        data=pdf_buffer.getvalue(),
        file_name="rigging_report.pdf",
        mime="application/pdf",
        type="primary",
    )
