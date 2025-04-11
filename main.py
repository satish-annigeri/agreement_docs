import sys
import time
import platform
from typing_extensions import Annotated


from rich.console import Console
import typer


from utils import tpl_suffix
from mergedata import read_data, prepare_data, group_data, unique_rows
from docxmerge import (
    docx_merge,
    docx2pdf_linux,
    docx2pdf_windows,
    print_header,
)  # get_docx_mergefields,
from htmlmerge import md_html_merge


app = typer.Typer()
con = Console()


def main(
    theatre: Annotated[str, typer.Argument(help="Theatres data in .xlsx format")],
    template: Annotated[
        str,
        typer.Option(
            "--template",
            "-t",
            help="Template file in .docx, .md.jinja or .html.jinja format",
        ),
    ] = "agreement_template.docx",
    css: Annotated[
        str,
        typer.Option(
            "--css",
            "-c",
            help="CSS stylesheet file for Markdown and HTML template files",
        ),
    ] = "agreement.css",
    distributor: Annotated[
        str,
        typer.Option("--distributor", "-d", help="Distributor data in .xlsx format"),
    ] = "distributors.xlsx",
    exhibitor: Annotated[
        str, typer.Option("--exhibitor", "-e", help="Exhibitors data in .xlsx format")
    ] = "exhibitors.xlsx",
):
    t_start = time.perf_counter()
    t1 = t_start
    print_header("Preparing Agreement Documents")

    distributor_fname = distributor
    exhibitor_fname = exhibitor
    theatre_fname = theatre
    template_fname = template

    distributors, exhibitors, theatres = read_data(
        distributor_fname, exhibitor_fname, theatre_fname, verbose=True
    )
    tpl_type = tpl_suffix(template_fname)
    print(f"Template file: {template_fname}")
    print(f"Template type: {tpl_type}")
    if tpl_type in ["md", "html"]:
        print(f"Stylesheet: {css}")
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
    t2 = time.perf_counter()
    con.log(f"Data preparation complete {t2 - t1:.2f}s")
    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

    if tpl_type == "docx":
        flist = docx_merge(
            distributors, grouped_df, num_groups, template_fname, fname_tpl
        )
        t3 = time.perf_counter()
        con.log(f"Generating Microsoft Word agreement documents took {t3 - t2:.2f}s")

        match platform.system():
            case "Linux":
                docx2pdf_linux(flist)
            case "Windows":
                docx2pdf_windows(flist)
            case _:
                raise OSError
        t4 = time.perf_counter()
        con.log(f"Converting Microsoft Word files to PDF took {t4 - t3:.2f}s")
        t_stop = t4
    elif tpl_type in ["md", "html"]:
        flist = md_html_merge(
            distributors,
            grouped_df,
            num_groups,
            template_fname,
            "",
            css,
            fname_tpl,
        )
        t3 = time.perf_counter()
        t_stop = t3
        con.log(f"Converting Markdown/HTML files to PDF took {t3 - t2:.2f}s")
    else:
        print(
            f"Invalid template file {template_fname}. Must be a '.md.jinja' or '.html.jinja'"
        )
        sys.exit(1)
    t_total = t_stop - t_start
    con.print(
        f"\nTotal execution time: {t_total:.2f}s for {len(flist)} files. Average: {t_total / len(flist):.2f}s per file."
    )


if __name__ == "__main__":
    typer.run(main)
