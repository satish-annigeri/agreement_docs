import sys
import time
import platform
from os.path import splitext
from typing_extensions import Annotated


from rich.console import Console


import typer


from docxmerge import (
    docx_merge,
    docx2pdf_linux,
    docx2pdf_windows,
    print_header,
)  # get_docx_mergefields,

from mergedata import read_data, prepare_data, group_data

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

    _, suffix = splitext(template_fname)
    if suffix == ".docx":
        print("\nPreparng Microsoft Word agreement files...")
        docx_flist = docx_merge(distributors, grouped_df, template_fname, fname_tpl)
    else:
        print(f"Invalid template file {template_fname}. Must be a .docx")
        sys.exit(1)
    t2 = time.perf_counter()
    print(f"\nGenerating Microsoft Word documents took: {t2 - t1:.4f} s")

    print("\nConverting Microsoft Word files to PDF and deleting them...")
    match platform.system():
        case "Linux":
            docx2pdf_linux(docx_flist)
        case "Windows":
            docx2pdf_windows(docx_flist)
        case _:
            raise OSError
    t3 = time.perf_counter()
    print(f"\nGenerating Microsoft Word documents took: {t2 - t1:.4f} s")
    print(f"Converting Microsoft Word documents to PDF took: {t3 - t2:.4f} s")
    print(
        f"Total execution time: {t3 - t1:.4f} s for {len(docx_flist)} files. Average: {(t3 - t1) / len(docx_flist):.4f} s per file."
    )


if __name__ == "__main__":
    typer.run(main)
