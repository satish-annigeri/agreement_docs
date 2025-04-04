import platform
from rich.console import Console
from typing_extensions import Annotated


import typer


from docxmerge import (
    docx_merge,
    docx2pdf_linux,
    docx2pdf_windows,
    print_header,
)  # get_docx_mergefields,

from mergedata import read_data, prepare_data, extract_distributor_data, group_data

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
    print_header("Preparing Agreement Documents")

    distributor_fname = distributor
    exhibitor_fname = exhibitor
    theatre_fname = theatre

    distributors, exhibitors, theatres = read_data(
        distributor_fname, exhibitor_fname, theatre_fname, verbose=True
    )
    distributors, df = prepare_data(distributors, exhibitors, theatres)

    distributor_data = extract_distributor_data(distributors)

    group_cols = [
        "exhibitor",
        "exhibitor_place",
        "movie",
        "release_date",
        "agreement_date",
    ]
    grouped_df = group_data(df, group_cols)

    print("\nPreparng Microsoft Word agreement files...")
    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"
    docx_flist = docx_merge(distributor_data, grouped_df, template, fname_tpl)
    print("\nConverting Microsoft Word files to PDF and deleting them...")
    match platform.system():
        case "Linux":
            docx2pdf_linux(docx_flist)
        case "Windows":
            docx2pdf_windows(docx_flist)
        case _:
            raise OSError


if __name__ == "__main__":
    typer.run(main)
