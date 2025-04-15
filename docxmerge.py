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


def detect_soffice_path(suggested_path: str = ""):
    if platform.system() == "Windows":
        res = subprocess.run("where soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            soffice_path = res.stdout.decode("utf-8").replace("\n", "")
        elif isfile(r"C:\Program Files\LibreOffice\program\soffice.exe"):
            soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        elif suggested_path and isfile(suggested_path):
            soffice_path = suggested_path
        else:
            raise FileNotFoundError(
                "LibreOffice not found. Please install LibreOffice."
            )
        cmd_list = [
            f"{soffice_path}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "{docx_fname}",
        ]
    elif platform.system() == "Linux":
        res = subprocess.run("which soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            soffice_path = res.stdout.decode("utf-8").replace("\n", "")
        elif isfile("/usr/bin/soffice"):
            soffice_path = "/usr/bin/soffice"
        elif suggested_path and isfile(suggested_path):
            soffice_path = suggested_path
        else:
            raise FileNotFoundError(
                "LibreOffice not found. Please install LibreOffice."
            )
        cmd_list = [
            f"{soffice_path}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "{docx_fname}",
            ">",
            "/dev/null",
            "2>&1",
        ]
    elif platform.system() == "Darwin":
        res = subprocess.run("which soffice", shell=True, capture_output=True)
        if res.returncode == 0:
            soffice_path = res.stdout.decode("utf-8").replace("\n", "")
        elif isfile("/Applications/LibreOffice.app/Contents/MacOS/soffice"):
            soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        elif suggested_path and isfile(suggested_path):
            soffice_path = suggested_path
        else:
            raise FileNotFoundError(
                "LibreOffice not found. Please install LibreOffice."
            )
        cmd_list = [
            f"{soffice_path}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "{docx_fname}",
            ">",
            "/dev/null",
            "2>&1",
        ]
    else:
        raise OSError("Unsupported OS")
    return soffice_path, cmd_list


def soffice_docx2pdf(docx_fname: str, cmd_list: list[str], verbose: bool = False):
    cmd_list[4] = cmd_list[4].format(docx_fname=docx_fname)
    res = subprocess.run(cmd_list, shell=True, capture_output=True)
    print(res)
    if verbose and res.returncode == 0:
        con.log(f"Converted {docx_fname} to PDF successfully.")


# def docx2pdf_linux(docx_flist):
#     if platform.system() == "Linux":
#         res = subprocess.run("which soffice", shell=True, capture_output=True)
#         if res.returncode == 0:
#             SOFFICE_PATH = res.stdout.decode("utf-8").replace("\n", "")
#         else:
#             SOFFICE_PATH = "/usr/bin/soffice"
#         if not isfile(SOFFICE_PATH):
#             raise FileNotFoundError
#         for docx_fname in docx_flist:
#             res = subprocess.run(
#                 [
#                     f"{SOFFICE_PATH}",
#                     "--headless",
#                     "--convert-to",
#                     "pdf:writer_pdf_Export",
#                     f"{docx_fname}",
#                     ">",
#                     "/dev/null",
#                     "2>&1",
#                 ],
#                 capture_output=True,
#             )
#             subprocess.run(f"rm {docx_fname}", shell=True)
#     else:
#         raise OSError


# def docx2pdf_windows(docx_flist):
#     if platform.system() == "Windows":
#         res = subprocess.run("where soffice", shell=True, capture_output=True)
#         if res.returncode == 0:
#             SOFFICE_PATH = res.stdout.decode("utf-8").replace("\n", "")
#         else:
#             SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
#         if not isfile(SOFFICE_PATH):
#             raise FileNotFoundError

#         for docx_fname in docx_flist:
#             res = subprocess.run(
#                 [
#                     f"{SOFFICE_PATH}",
#                     "--headless",
#                     "--convert-to",
#                     "pdf:writer_pdf_Export",
#                     f"{docx_fname}",
#                 ],
#                 shell=True,
#                 capture_output=True,
#             )
#             subprocess.run(f"del {docx_fname}", shell=True)
#     else:
#         raise OSError


if __name__ == "__main__":
    soffice_path, cmd_list = detect_soffice_path()
    if soffice_path:
        print(f"LibreOffice path: {soffice_path}")
        print(f"Command list: {cmd_list}")
        soffice_docx2pdf("agreement_template.docx", cmd_list)
    else:
        print("LibreOffice path not detected. Please install LibreOffice.")
