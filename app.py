import sys
import os
from os.path import isfile
from time import perf_counter
from datetime import datetime
import typer
from typing_extensions import Annotated
import io
import zipfile


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


def create_zip(flist: list[str], zip_name: str = "output.zip"):
    with zipfile.ZipFile(zip_name, "w") as zipf:
        for fname in flist:
            zipf.write(fname)


def st_read_data(
    distributors_fname, exhibitors_fname, theatres_fname, template_fname, css_fname
):
    if not theatres_fname:
        st.session_state.run_mode = "interactive"
        distributors = st.file_uploader("Upload distributors data file", type=["xlsx"])
        exhibitors = st.file_uploader("Upload exhibitors data file", type=["xlsx"])
        theatres = st.file_uploader("Upload theatres data file", type=["xlsx"])
        st.session_state.distributors = (
            pl.read_excel(distributors) if distributors else pl.DataFrame()
        )
        st.session_state.exhibitors = (
            pl.read_excel(exhibitors) if exhibitors else pl.DataFrame()
        )
        st.session_state.theatres = (
            pl.read_excel(theatres) if theatres else pl.DataFrame()
        )
        tpl = st.file_uploader("Upload agreement template file", type=["jinja", "docx"])
        st.session_state.tpl_fname = tpl.name if tpl else ""
        if tpl and tpl_suffix(tpl.name) in ["html", "md"]:
            css = st.file_uploader(
                "Upload CSS file for HTML/Markdown template", type=["css"]
            )
            st.session_state.css_fname = css.name if css else ""
        if st.session_state.run_mode == "interactive":
            button_enabled = (
                st.session_state.distributors.height > 0
                and st.session_state.exhibitors.height > 0
                and st.session_state.theatres.height > 0
                and st.session_state.tpl_fname
            )
            if tpl_suffix(st.session_state.tpl_fname) in ["html", "md"]:
                button_enabled = button_enabled and st.session_state.css_fname
            button = st.button(
                "Continue",
                disabled=not button_enabled,
                help="Select files and click to generate agreement documents",
            )
            if button:
                st.session_state.read_data = "continue"
    else:
        st.session_state.run_mode = "batch"
        st.session_state.distributors = pl.read_excel(distributors_fname)
        st.session_state.exhibitors = pl.read_excel(exhibitors_fname)
        st.session_state.theatres = pl.read_excel(theatres_fname)
        st.session_state.tpl_fname = template_fname if isfile(template_fname) else ""
        st.session_state.css_fname = css_fname if isfile(css_fname) else ""
        st.session_state.read_data = "continue"


if "read_data" not in st.session_state:
    st.session_state.read_data = ""
if "distributors" not in st.session_state:
    st.session_state.distributors = pl.DataFrame()
if "exhibitors" not in st.session_state:
    st.session_state.exhibitors = pl.DataFrame()
if "theatres" not in st.session_state:
    st.session_state.theatres = pl.DataFrame()
if "tpl_fname" not in st.session_state:
    st.session_state.tpl_fname = ""
if "css_fname" not in st.session_state:
    st.session_state.css_fname = ""
if "zip_downloaded" not in st.session_state:
    st.session_state.zip_downloaded = False


# ---- Streamlit App ----


def st_app():
    if st.session_state.read_data == "continue":
        t_start = perf_counter()
        st.title("Agreement Document Generator")

        tpl_type = tpl_suffix(st.session_state.tpl_fname)
        msg = f"""### Input\n
* Distributors data: {st.session_state.distributors.height} rows
* Exhibitors data: {st.session_state.exhibitors.height} rows
* Theatres data: {st.session_state.theatres.height} rows 
* Template file: {st.session_state.tpl_fname}
* Template type: {tpl_type}"""
        if st.session_state.css_fname:
            msg += f"\n* CSS file: {st.session_state.css_fname}"
        st.markdown(msg)

        if tpl_type not in ["docx", "md", "html"]:
            st.error(
                f"Unknown template type: {tpl_type}. Supported types are: md, html, docx\nProgram aborted"
            )
            st.error(
                "**Template file must be in .html.jinja, .md.jinja or .docx format**"
            )
            st.stop()

        # Clean and prepare data for use

        distributors, df = prepare_data(
            st.session_state.distributors,
            st.session_state.exhibitors,
            st.session_state.theatres,
        )

        group_cols = [
            "exhibitor",
            "exhibitor_place",
            "movie",
            "release_date",
            "agreement_date",
        ]
        grouped_df = group_data(df, group_cols)
        num_exhibitors = unique_rows(df, group_cols)
        st.markdown(f"**Number of exhibitors: {num_exhibitors}**")

        fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

        distributor_data = extract_distributor_data(distributors)

        if tpl_type in ["md", "html"]:
            jinja_template = get_jinja2_template(st.session_state.tpl_fname, "")
        elif tpl_type == "docx":
            soffice_path, cmd_list, shell = detect_soffice_path()
        else:
            print(
                f"Unknown template type: {tpl_type}. Supported types are: md, html, docx\nProgram aborted"
            )
            sys.exit(1)

        # --------------------------
        with st.status(
            "Generating agreement document files...", expanded=True
        ) as status:
            count = 0
            flist = []
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
                        st.session_state.tpl_fname,
                        output_fname,
                        distributor_data,
                        exhibitor_data,
                        annexure,
                    )
                    st.write(f"Generating {with_suffix(output_fname, '.pdf')}")
                    soffice_docx2pdf(output_fname, cmd_list, shell)
                    os.remove(output_fname)
                elif tpl_type in ["html", "md"]:
                    # st.write(f"{st.session_state.css_fname}")
                    output_fname = f"{output_fname}.pdf"
                    md_html_mergefields(
                        jinja_template,
                        tpl_type,
                        st.session_state.css_fname,
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
                flist.append(output_fname)
            t2 = t_stop = perf_counter()

            status.update(
                label=f"Generation of agreement files is complete in {t2 - t_start:.2f}s at {(t2 - t_start) / num_exhibitors:.2f}s per file",
                state="complete",
                expanded=False,
            )
            st.success(
                f"Generation of PDF files is complete in {t_stop - t2:.2f}s at {(t_stop - t_start) / num_exhibitors:.2f}s per file",
            )

        # Zip PDF files for download
        today = datetime.today().strftime("%Y-%m-%d")
        zip_fname = f"agreement_docs_{today}.zip"
        create_zip(flist, zip_fname)
        with open(zip_fname, "rb") as f:
            st.download_button(
                label="Download ZIP file",
                data=f,
                file_name=zip_fname,
                mime="application/zip",
                help="Download the generated agreement documents as a ZIP file",
            )
        st.session_state.zip_downloaded = True
        st.success(f"ZIP file '{zip_fname}' downloaded successfully.")
        # Clean up files
        for file in flist:
            os.remove(file)

        st.success(
            f"Total time taken: {t_stop - t_start:.2f}s at {(t_stop - t_start) / num_exhibitors:.2f}s per file"
        )


def main(
    theatres: Annotated[str | None, typer.Argument()] = "",
    distributors: Annotated[
        str, typer.Option("--distributors", "-d")
    ] = "distributors.xlsx",
    exhibitors: Annotated[str, typer.Option("--exhibitors", "-e")] = "exhibitors.xlsx",
    template: Annotated[str, typer.Option("--template", "-t")] = "agreement.html.jinja",
    css: Annotated[str, typer.Option("--css", "-c")] = "agreement.css",
):
    st.write(f"Theatres: '{theatres}'")
    st_read_data(distributors, exhibitors, theatres, template, css)
    st_app()
    if st.session_state.zip_downloaded:
        st.stop()


if __name__ == "__main__":
    typer.run(main)
