from os.path import isfile, splitext
import subprocess
import platform


from mailmerge import MailMerge
from rich.console import Console


from mergedata import (
    read_excel,
    clean_distributors_data,
    clean_exhibitors_data,
    clean_theatres_data,
    join_data,
    group_data,
    extract_distributor_data,
    extract_exhibitor_data,
    extract_annexure_data,
)

con = Console()


# ---- Utility Functions ----


def with_suffix(fpath: str, suffix: str) -> str:
    base, _ = splitext(fpath)
    return base + suffix


def print_header(header: str, underline: str = "-"):
    print()
    print(header)
    print(f"{underline * len(header)}")
    print()


# ---- Functions to merge prepared data with template ----


def get_docx_mergefields(docx_fname: str):
    if not isfile(docx_fname):
        raise FileNotFoundError
    else:
        tpl = MailMerge(docx_fname)
        return tpl.get_merge_fields()


def docx_mergefields(
    docx_tpl: str,
    docx_output_fname: str,
    distributor_data,
    exhibitor_data,
    annexure,
):
    tpl = MailMerge(docx_tpl)
    tpl.merge(**distributor_data)
    tpl.merge(**exhibitor_data)
    tpl.merge_rows("slno", annexure)
    tpl.write(docx_output_fname)


def docx_merge(
    distributors_fname, exhibitors_fname, theatres_fname, docx_tpl_fname, fname_tpl
):
    print(f"Reading: {distributors_fname}")
    distributors = read_excel(distributors_fname)
    print(f"Reading: {exhibitors_fname}")
    exhibitors = read_excel(exhibitors_fname)
    print(f"Reading: {theatres_fname}")
    theatres = read_excel(theatres_fname)
    distributors = clean_distributors_data(distributors)
    exhibitors = clean_exhibitors_data(exhibitors)
    theatres = clean_theatres_data(theatres)

    print("\nPreparing data for merging...")
    df = join_data(exhibitors, theatres)

    distributor_data = extract_distributor_data(distributors)

    group_cols = [
        "exhibitor",
        "exhibitor_place",
        "movie",
        "release_date",
        "agreement_date",
    ]
    grouped_df = group_data(df, group_cols)
    count = 0
    flist = []
    for g_exhibitors, g_theatres in grouped_df:
        count += 1
        exhibitor_data = extract_exhibitor_data(g_exhibitors, g_theatres)
        annexure = extract_annexure_data(g_theatres)
        docx_output_fname = f"{
            fname_tpl.format(
                count=count,
                movie=exhibitor_data['movie'].lower(),
                exhibitor=exhibitor_data['exhibitor']
                .replace('/', '')
                .replace(' ', '_')
                .lower(),
                release_date=exhibitor_data['release_date'],
            )
        }.docx"
        print(f"\t{docx_output_fname}")

        docx_mergefields(
            docx_tpl_fname,
            docx_output_fname,
            distributor_data,
            exhibitor_data,
            annexure,
        )
        flist.append(docx_output_fname)
    return flist


def docx2pdf_linux(docx_flist):
    if platform.system() == "Linux":
        res = subprocess.run("which soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            SOFFICE_PATH = res.stdout.decode("utf-8").replace("\n", "")
        else:
            SOFFICE_PATH = "/usr/bin/soffice"
        if not isfile(SOFFICE_PATH):
            raise FileNotFoundError

        for docx_fname in docx_flist:
            subprocess.run(
                [
                    f"{SOFFICE_PATH}",
                    "--headless",
                    "--convert-to",
                    "pdf:writer_pdf_Export",
                    f"{docx_fname}",
                    ">",
                    "/dev/null",
                    "2>&1",
                ],
                shell=True,
                # capture_output=True,
            )
            subprocess.run(f"rm {docx_fname}", shell=True)
            print(f"\t{with_suffix(docx_fname, '.pdf')}")
    else:
        raise OSError


def docx2pdf_windows(docx_flist):
    if platform.system() == "Windows":
        res = subprocess.run("where soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            SOFFICE_PATH = res.stdout.decode("utf-8").replace("\n", "")
        else:
            SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
        if not isfile(SOFFICE_PATH):
            raise FileNotFoundError

        for docx_fname in docx_flist:
            res = subprocess.run(
                [
                    f"{SOFFICE_PATH}",
                    "--headless",
                    "--convert-to",
                    "pdf:writer_pdf_Export",
                    f"{docx_fname}",
                ],
                shell=True,
                capture_output=True,
            )
            print(f"\t{with_suffix(docx_fname, '.pdf')}")
            subprocess.run(f"del {docx_fname}", shell=True)
    else:
        raise OSError
