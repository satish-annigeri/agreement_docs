import sys
import time
import platform
from typing_extensions import Annotated


from rich.console import Console
import typer


from utils import tpl_suffix
from mergedata import read_data, prepare_data, group_data
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
    theatre: str,
    template: Annotated[
        str, typer.Option("--template", "-t")
    ] = "agreement_template.docx",
    distributor: Annotated[
        str, typer.Option("--distributor", "-d")
    ] = "distributors.xlsx",
    exhibitor: Annotated[str, typer.Option("--exhibitor", "-e")] = "exhibitors.xlsx",
):
    t1 = time.perf_counter()
    print_header("Preparing Agreement Documents")

    distributor_fname = distributor
    exhibitor_fname = exhibitor
    theatre_fname = theatre
    template_fname = template

    distributors, exhibitors, theatres = read_data(
        distributor_fname, exhibitor_fname, theatre_fname, verbose=True
    )
    distributors, df = prepare_data(distributors, exhibitors, theatres)

    group_cols = [
        "exhibitor",
        "exhibitor_place",
        "movie",
        "release_date",
        "agreement_date",
    ]
    grouped_df = group_data(df, group_cols)

    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

    tpl_type = tpl_suffix(template_fname)

    if tpl_type == "docx":
        print("\nPreparng Microsoft Word agreement files...")
        flist = docx_merge(distributors, grouped_df, template_fname, fname_tpl)

        print("\nConverting Microsoft Word files to PDF and deleting them...")
        match platform.system():
            case "Linux":
                docx2pdf_linux(flist)
            case "Windows":
                docx2pdf_windows(flist)
            case _:
                raise OSError
    elif tpl_type in ["md", "html"]:
        flist = md_html_merge(
            distributors, grouped_df, template_fname, "", "style.css", fname_tpl
        )
        for pdf_fname in flist:
            print(pdf_fname)
    else:
        print(
            f"Invalid template file {template_fname}. Must be a '.md.jinja' or '.html.jinja'"
        )
        sys.exit(1)

    t2 = time.perf_counter()
    t3 = time.perf_counter()
    print(f"\nGenerating Microsoft Word documents: {t2 - t1:.4f}s")
    print(f"Converting Microsoft Word documents to PDF: {t3 - t2:.4f}s")
    print(
        f"Total execution time: {t3 - t1:.4f}s for {len(flist)} files. Average: {(t3 - t1) / len(flist):.4f}s per file."
    )


if __name__ == "__main__":
    typer.run(main)
