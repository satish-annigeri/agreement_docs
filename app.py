import sys
import os
from time import perf_counter
import argparse


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
from utils import tpl_suffix, get_fname, with_suffix
from htmlmerge import md_html_mergefields, get_jinja2_template
from docxmerge import (
    docx_mergefields,
    detect_soffice_path,
    soffice_docx2pdf,
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Agreement Document Generator App",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("theatres", help="Theatres data file in .xlsx format")
    parser.add_argument(
        "-e",
        "--exhibitors",
        default="exhibitors.xlsx",
        help="Exhibitors data file in .xlsx format",
    )
    parser.add_argument(
        "-d",
        "--distributors",
        default="distributors.xlsx",
        help="Distributor data file in .xlsx format",
    )
    parser.add_argument(
        "-t",
        "--template",
        default="agreement.html.jinja",
        help="Template for agreement document in .html.jinja, .md.jinja or .docx format",
    )
    parser.add_argument(
        "-c", "--css", default="agreement.css", help="CSS file for HTML template"
    )
    return parser.parse_args()


# ---- Streamlit App ----


t_start = perf_counter()
st.title("Agreement Document Generator")

args = parse_args()
print(parse_args())

# Data files
distributors_fname = args.distributors  # "distributors.xlsx"
exhibitors_fname = args.exhibitors  # "exhibitors.xlsx"
theatres_fname = args.theatres  # "chhaava_theatres.xlsx"
template_fname = args.template  # "agreement.html.jinja"
# template_fname = "agreement_template.docx"
tpl_type = tpl_suffix(template_fname)
css_fname = args.css  # "agreement.css" if tpl_type in ["html", "md"] else ""


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
elif tpl_type == "docx":
    pass
else:
    st.error(
        f"Unknown template type: {tpl_type}. Supported types are: md, html, docx\nProgram aborted"
    )
    sys.exit(1)

# Read data
distributors = pl.read_excel(distributors_fname)
exhibitors = pl.read_excel(exhibitors_fname)
theatres = pl.read_excel(theatres_fname)

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
num_exhibitors = unique_rows(df, group_cols)
st.write(f"Number of exhibitors: {num_exhibitors}")

fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

distributor_data = extract_distributor_data(distributors)


if tpl_type in ["md", "html"]:
    jinja_template = get_jinja2_template(template_fname, "")
elif tpl_type == "docx":
    soffice_path, cmd_list, shell = detect_soffice_path()
else:
    print(
        f"Unknown template type: {tpl_type}. Supported types are: md, html, docx\nProgram aborted"
    )
    sys.exit(1)

# --------------------------
with st.status("Generating agreement document files...", expanded=True) as status:
    count = 0
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
            st.write(f"Generating {with_suffix(output_fname, '.pdf')}")
            soffice_docx2pdf(output_fname, cmd_list, shell)
            os.remove(output_fname)
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
            st.error(
                f"Unknown template type: {tpl_type}. Supported types are: md, html, docx"
            )
            sys.exit(1)
    t2 = t_stop = perf_counter()

    status.update(
        label=f"Generation of agreement files is complete in {t2 - t_start:.2f}s at {(t2 - t_start) / num_exhibitors:.2f}s per file",
        state="complete",
        expanded=False,
    )
    st.success(
        f"Generation of PDF files is complete in {t_stop - t2:.2f}s at {(t_stop - t_start) / num_exhibitors:.2f}s per file",
    )
st.success(
    f"Total time taken: {t_stop - t_start:.2f}s at {(t_stop - t_start) / num_exhibitors:.2f}s per file"
)
