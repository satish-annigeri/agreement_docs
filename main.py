import platform
from rich.console import Console

from docxmerge import get_docx_mergefields, docx_merge, docx2pdf_linux, docx2pdf_windows


con = Console()


def main(distributor_fname, exhibitor_fname, theatre_fname, docx_tpl_fname):
    print(f"\n\n{'=' * 70}\n\n")

    con.print(get_docx_mergefields("agreement_template.docx"))

    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"
    docx_flist = docx_merge(
        distributor_fname, exhibitor_fname, theatre_fname, docx_tpl_fname, fname_tpl
    )
    match platform.system():
        case "Linux":
            docx2pdf_linux(docx_flist)
        case "Windows":
            docx2pdf_windows(docx_flist)
        case _:
            raise OSError


if __name__ == "__main__":
    distributor_fname = "distributors.xlsx"
    exhibitor_fname = "exhibitors.xlsx"
    theatre_fname = "chhaava_theatres.xlsx"
    docx_tpl_fname = "agreement_template.docx"
    main(distributor_fname, exhibitor_fname, theatre_fname, docx_tpl_fname)
