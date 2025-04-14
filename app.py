import sys
import platform
from time import perf_counter


import streamlit as st

import polars as pl

from mergedata import (
    prepare_data,
    group_data,
    unique_rows,
    extract_distributor_data,
    extract_exhibitor_data,
    extract_annexure_data,
)
from utils import tpl_suffix, get_fname
from htmlmerge import md_html_mergefields, get_jinja2_template
from docxmerge import docx_mergefields, docx2pdf_linux, docx2pdf_windows


t_start = perf_counter()
st.title("Agreement Document Generator App")

# Data files
distributors_fname = "distributors.xlsx"
exhibitors_fname = "exhibitors.xlsx"
theatres_fname = "chhaava_theatres.xlsx"
template_fname = "agreement.html.jinja"
# template_fname = "agreement_template.docx"
tpl_type = tpl_suffix(template_fname)
css_fname = "agreement.css" if tpl_type in ["html", "md"] else None


# Read data
distributors = pl.read_excel(distributors_fname)
exhibitors = pl.read_excel(exhibitors_fname)
theatres = pl.read_excel(theatres_fname)

msg = f"""### Input files

* Distributors data: {distributors_fname}
* Exhibitors data: {exhibitors_fname}
* Theatres data: {theatres_fname}
* Template file: {template_fname}
* Template type: {tpl_type}"""
if css_fname:
    msg += f"""\n* CSS file: {css_fname}"""
st.markdown(msg)

if tpl_type in ["md", "html"]:
    st.write(f"CSS file: {css_fname}")

# Clean and prepare data for use
distributors, df = prepare_data(distributors, exhibitors, theatres)

group_cols = [
    "exhibitor",
    "exhibitor_place",
    "movie",
    "release_date",
    "agreement_date",
]
grouped_df = group_data(df, group_cols)
num_groups = unique_rows(df, group_cols)
st.write(f"Number of unique groups: {num_groups}")

fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

distributor_data = extract_distributor_data(distributors)

count = 0
flist = []
if tpl_type in ["md", "html"]:
    jinja_template = get_jinja2_template(template_fname, "")

# --------------------------
with st.status("Generating agreement document files...", expanded=True) as status:
    for g_exhibitor, g_theatres in grouped_df:
        count += 1
        exhibitor_data = extract_exhibitor_data(g_exhibitor, g_theatres)
        annexure = extract_annexure_data(g_theatres)

        output_fname = get_fname(
            fname_tpl,
            count=count,
            movie=exhibitor_data["movie"].lower(),
            exhibitor=exhibitor_data["exhibitor"],
            release_date=exhibitor_data["release_date"],
        )

        if tpl_type == "docx":
            output_fname = f"{output_fname}.docx"
            docx_mergefields(
                template_fname,
                output_fname,
                distributor_data,
                exhibitor_data,
                annexure,
            )
            st.write(f"Generating {output_fname}")
        elif tpl_type in ["html", "md"]:
            output_fname = f"{output_fname}.pdf"
            md_html_mergefields(
                jinja_template,
                tpl_type,
                css_fname,
                output_fname,
                distributor_data=distributor_data,
                exhibitor_data=exhibitor_data,
                annexure=annexure,
            )
            st.write(f"Generating {output_fname}")
        else:
            sys.exit(1)
        flist.append(output_fname)
    t2 = perf_counter()
    status.update(
        label=f"Generation of agreement files is complete in {t2 - t_start:.2f}s at {(t2 - t_start) / num_groups:.2f}s per file",
        state="complete",
        expanded=False,
    )
    # --------------------------
if tpl_type in ["md", "html"]:
    t_stop = perf_counter()
elif tpl_type == "docx":
    with st.status(
        "Converting Microsoft Word documents to PDF...", expanded=True
    ) as status:
        match platform.system():
            case "Linux":
                docx2pdf_linux(flist)
            case "Windows":
                docx2pdf_windows(flist)
            case _:
                raise OSError
        t_stop = perf_counter()
    status.update(
        label=f"Generation of PDF files is complete in {t_stop - t2:.2f}s at {(t_stop - t_start) / num_groups:.2f}s per file",
        state="complete",
        expanded=False,
    )
    st.write(
        f"Total time taken: {t_stop - t_start:.2f}s at {(t_stop - t_start) / num_groups:.2f}s per file"
    )
