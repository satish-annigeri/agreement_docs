from os.path import isfile
import subprocess
import platform


from mailmerge import MailMerge
from rich.console import Console
from rich.progress import (
    Progress,
    TextColumn,
    SpinnerColumn,
    MofNCompleteColumn,
    TaskProgressColumn,
    TimeElapsedColumn,
)

from mergedata import (
    extract_distributor_data,
    extract_exhibitor_data,
    extract_annexure_data,
)
from utils import with_suffix

con = Console()


# ---- Utility Functions ----


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


def docx_merge(distributors, grouped_df, num_groups: int, docx_tpl_fname, fname_tpl):
    distributor_data = extract_distributor_data(distributors)

    count = 0
    flist = []

    progress = Progress(
        TaskProgressColumn(),
        SpinnerColumn(),
        MofNCompleteColumn(),
        TextColumn("[cyan]{task.fields[progress_description]}"),
        TextColumn("[bold cyan]{task.fields[task_description]}"),
    )
    with progress:
        task = progress.add_task(
            "[cyan]Generating Word file:",
            total=num_groups,
            progress_description="[cyan]Generating Word file:",
            task_description="Filename",
        )
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
            progress.update(
                task, task_description=f"{with_suffix(docx_output_fname, '.pdf')}"
            )

            docx_mergefields(
                docx_tpl_fname,
                docx_output_fname,
                distributor_data,
                exhibitor_data,
                annexure,
            )
            progress.advance(task)

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

        progress = Progress(
            TaskProgressColumn(),
            SpinnerColumn(),
            TimeElapsedColumn(),
            MofNCompleteColumn(),
            TextColumn("[cyan]{task.fields[progress_description]}"),
            TextColumn("[bold cyan]{task.fields[task_description]}"),
        )
        with progress:
            task = progress.add_task(
                "",
                total=len(docx_flist),
                progress_description="",
                task_description="Filename",
            )

            for docx_fname in docx_flist:
                progress.update(
                    task, task_description=f"{with_suffix(docx_fname, '.pdf')}"
                )
                res = subprocess.run(
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
                    capture_output=True,
                )
                subprocess.run(f"rm {docx_fname}", shell=True)
                progress.advance(task)
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

        progress = Progress(
            TaskProgressColumn(),
            SpinnerColumn(),
            TimeElapsedColumn(),
            MofNCompleteColumn(),
            TextColumn("[cyan]{task.fields[progress_description]}"),
            TextColumn("[bold cyan]{task.fields[task_description]}"),
        )
        with progress:
            task = progress.add_task(
                "",
                total=len(docx_flist),
                progress_description="",
                task_description="Filename",
            )

            for docx_fname in docx_flist:
                progress.update(
                    task, task_description=f"{with_suffix(docx_fname, '.pdf')}"
                )
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
                subprocess.run(f"del {docx_fname}", shell=True)
                progress.advance(task)
    else:
        raise OSError
