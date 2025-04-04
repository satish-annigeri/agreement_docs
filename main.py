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

    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"
    docx_flist = docx_merge(distributor, exhibitor, theatre, template, fname_tpl)
    print("\nConverting Word files to PDF and deleting Word files...")
    match platform.system():
        case "Linux":
            docx2pdf_linux(docx_flist)
        case "Windows":
            docx2pdf_windows(docx_flist)
        case _:
            raise OSError


if __name__ == "__main__":
    typer.run(main)
