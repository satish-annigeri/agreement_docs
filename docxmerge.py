from os.path import isfile
import subprocess
import platform


from mailmerge import MailMerge
from rich.console import Console


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
    else:
        raise OSError


def detect_soffice_path():
    if platform.system() == "Windows":
        res = subprocess.run("where soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            return res.stdout.decode("utf-8").replace("\n", "")
        else:
            return r"C:\Program Files\LibreOffice\program\soffice.exe"
    elif platform.system() == "Linux":
        res = subprocess.run("which soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            return res.stdout.decode("utf-8").replace("\n", "")
        else:
            return "/usr/bin/soffice"
    elif platform.system() == "Darwin":
        res = subprocess.run("which soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            return res.stdout.decode("utf-8").replace("\n", "")
        else:
            return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        raise OSError("Unsupported OS")


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
            subprocess.run(f"del {docx_fname}", shell=True)
    else:
        raise OSError


if __name__ == "__main__":
    soffice_path = detect_soffice_path()
    if soffice_path:
        print(f"LibreOffice path detected: {soffice_path}")
    else:
        print("LibreOffice path not detected. Please install LibreOffice.")
