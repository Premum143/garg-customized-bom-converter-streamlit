from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

from garg_customized_converter import (
    DEFAULT_BOM_NUMBER,
    DEFAULT_STORES,
    FG_HEADERS,
    MASTER_HEADERS,
    RM_HEADERS,
    ROUTING_HEADERS,
    ConversionError,
    build_tranzact_rows,
    build_workbook,
    parse_customized_bom,
    workbook_to_bytes,
)


st.set_page_config(page_title="GARG Customized BOM Converter", layout="wide")
st.title("GARG Customized BOM Converter")
st.caption("Assembly BOM only. This tool assumes cut parts already exist in inventory and builds FG/SFG BOMs that consume those cut parts.")


with st.sidebar:
    st.subheader("Conversion Settings")
    bom_number = st.text_input("BOM Number", value=DEFAULT_BOM_NUMBER)
    fg_store = st.text_input("FG Store", value=DEFAULT_STORES["fg_store"])
    rm_store = st.text_input("RM Store", value=DEFAULT_STORES["rm_store"])
    scrap_store = st.text_input("Reject Store", value=DEFAULT_STORES["scrap_store"])
    wip_store = st.text_input("WIP Store", value=DEFAULT_STORES["wip_store"])

    st.divider()
    st.subheader("Model")
    st.markdown(
        "\n".join(
            [
                "1. `Item ID` for cut parts = `DRAWING NO`.",
                "2. `Item Name` = `DESCRIPTION + MATERIAL CODE + size`.",
                "3. Flat files become one FG consuming cut parts directly.",
                "4. Section-based files create SFGs from `GEL_..._A01/A02/...` markers.",
                "5. Final FG consumes SFGs; each SFG consumes its own cut parts.",
                "6. Cut-part BOM and nesting are handled outside this tool.",
            ]
        )
    )


uploaded_files = st.file_uploader(
    "Upload customized GARG BOM files",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload one or more customized BOM workbooks to begin.")
    st.stop()

upload_signature = tuple(sorted(uploaded_file.name for uploaded_file in uploaded_files))
if st.session_state.get("customized_upload_signature") != upload_signature:
    st.session_state["customized_upload_signature"] = upload_signature
    st.session_state.pop("customized_result", None)
    st.session_state.pop("customized_output", None)

parsed_products = []
errors = []
for uploaded_file in uploaded_files:
    try:
        parsed = parse_customized_bom(BytesIO(uploaded_file.getvalue()), filename=uploaded_file.name)
        parsed_products.append(parsed)
    except ConversionError as exc:
        errors.append({"File": uploaded_file.name, "Error": str(exc)})
    except Exception as exc:  # pragma: no cover
        errors.append({"File": uploaded_file.name, "Error": f"Unexpected error: {exc}"})

summary_rows = []
for parsed in parsed_products:
    summary_rows.append(
        {
            "File": parsed["path"].name,
            "FG Code": parsed["metadata"]["fg_code"],
            "FG Name": parsed["metadata"]["description"],
            "Mode": parsed["mode"],
            "Root Components": len(parsed["root_components"]),
            "Sections": len(parsed["sections"]),
            "Skipped Markers": len(parsed["skipped_rows"]),
        }
    )

if errors:
    st.error("Some files could not be parsed.")
    st.dataframe(pd.DataFrame(errors), width="stretch", hide_index=True)
    st.caption("If a valid multi-level BOM fails here, the error table above is the source of truth. It will show the exact parse reason for that file.")

if summary_rows:
    st.subheader("Detected Workbooks")
    st.dataframe(pd.DataFrame(summary_rows), width="stretch", hide_index=True)

detail_tabs = st.tabs(["Section Detail", "Skipped Rows"])

with detail_tabs[0]:
    detail_rows = []
    for parsed in parsed_products:
        if parsed["mode"] == "multi_level":
            for section in parsed["sections"]:
                detail_rows.append(
                    {
                        "File": parsed["path"].name,
                        "Parent FG": parsed["metadata"]["fg_code"],
                        "SFG Code": section["code"],
                        "SFG Name": section["name"],
                        "Components": len(section["rows"]),
                    }
                )
    if detail_rows:
        st.dataframe(pd.DataFrame(detail_rows), width="stretch", hide_index=True)
    else:
        st.info("No section-based SFG structures detected in the uploaded files.")

with detail_tabs[1]:
    skipped_rows = []
    for parsed in parsed_products:
        for row in parsed["skipped_rows"]:
            skipped_rows.append(
                {
                    "File": parsed["path"].name,
                    "Drawing No": row[0],
                    "Description": row[1],
                    "Material Code": row[2],
                    "Qty": row[3],
                }
            )
    if skipped_rows:
        st.dataframe(pd.DataFrame(skipped_rows), width="stretch", hide_index=True)
    else:
        st.info("No skipped non-component rows were recorded.")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Uploaded Files", len(uploaded_files))
col2.metric("Parsed Files", len(parsed_products))
col3.metric("Multi-level Files", sum(1 for parsed in parsed_products if parsed["mode"] == "multi_level"))
col4.metric("Flat Files", sum(1 for parsed in parsed_products if parsed["mode"] == "flat"))

generate_disabled = bool(errors or not parsed_products)
if st.button("Generate Assembly BOM Workbook", type="primary", disabled=generate_disabled):
    result = build_tranzact_rows(
        parsed_products,
        bom_number=bom_number.strip() or DEFAULT_BOM_NUMBER,
        stores={
            "fg_store": fg_store,
            "rm_store": rm_store,
            "scrap_store": scrap_store,
            "wip_store": wip_store,
        },
    )
    workbook = build_workbook(result)
    st.session_state["customized_result"] = result
    st.session_state["customized_output"] = workbook_to_bytes(workbook)

result = st.session_state.get("customized_result")
output_bytes = st.session_state.get("customized_output")

if result and output_bytes:
    st.success("Workbook generated.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("FG Rows", len(result["fg_rows"]))
    col2.metric("RM Rows", len(result["rm_rows"]))
    col3.metric("Routing Rows", len(result["routing_rows"]))
    col4.metric("Master Rows", len(result["master_rows"]))

    filename = f"GARG_CUSTOMIZED_ASSEMBLY_BOM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "Download Workbook",
        data=output_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    preview_tabs = st.tabs(["FG", "RM", "Routing", "Master File"])
    preview_rows = [
        (result["fg_rows"], FG_HEADERS),
        (result["rm_rows"], RM_HEADERS),
        (result["routing_rows"], ROUTING_HEADERS),
        (result["master_rows"], MASTER_HEADERS),
    ]
    for tab, (rows, headers) in zip(preview_tabs, preview_rows):
        with tab:
            sample = rows[:100]
            st.caption(f"Showing first {len(sample)} rows.")
            st.dataframe(pd.DataFrame(sample, columns=headers), width="stretch", hide_index=True)
