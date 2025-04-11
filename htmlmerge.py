import re
import time
from os.path import abspath, splitext


from rich.console import Console
from rich.progress import (
    Progress,
    TextColumn,
    SpinnerColumn,
    MofNCompleteColumn,
    TaskProgressColumn,
    TimeElapsedColumn,
)
import pendulum
from jinja2 import Environment, FileSystemLoader
import mistune
from weasyprint import HTML, __version__ as wezp_ver

from utils import tpl_suffix

from mergedata import (
    extract_distributor_data,
    extract_exhibitor_data,
    extract_annexure_data,
)

re_html_fname = re.compile(r".*[.]html$", re.I)


con = Console()


def is_html_fname(s: str) -> bool:
    if re_html_fname.match(s):
        return True
    else:
        return False


def get_jinja2_template(tpl_fname: str, tpl_dir: str):
    tpl_dir = abspath(tpl_dir)
    env = Environment(loader=FileSystemLoader(tpl_dir))
    tpl = env.get_template(tpl_fname)
    return tpl


def md2html(md_fname: str) -> str:
    with open(md_fname, "r", encoding="utf-8") as md_file:
        md_content = md_file.read()
    html = mistune.html(md_content)

    if isinstance(html, str):
        return html
    else:
        return ""


def md_html_mergefields(
    jinja_tpl,
    tpl_type: str,
    # tpl_fname: str,
    # tpl_dir: str,
    css: str,
    pdf_fname: str,
    distributor_data,
    exhibitor_data,
    annexure,
):
    time_now = pendulum.now().format("YYYY-MM-DDTHH:MM:SSZ")
    if tpl_type == "md":
        md_content = jinja_tpl.render(
            **distributor_data,
            **exhibitor_data,
            annexure=annexure,
            time_now=time_now,
            weasyprint_ver=wezp_ver,
        )
        html_content = mistune.html(md_content)
    elif tpl_type == "html":
        html_content = jinja_tpl.render(
            **distributor_data,
            **exhibitor_data,
            annexure=annexure,
            time_now=time_now,
            weasyprint_ver=wezp_ver,
        )
    HTML(string=html_content).write_pdf(pdf_fname, stylesheets=[css])


def md_html_merge(
    distributors, grouped_df, num_groups: int, tpl_fname, tpl_dir, css, fname_tpl
):
    distributor_data = extract_distributor_data(distributors)

    count = 0
    flist = []

    progress = Progress(
        TaskProgressColumn(),
        SpinnerColumn(),
        TimeElapsedColumn(),
        MofNCompleteColumn(),
        TextColumn("[cyan]{task.fields[progress_description]}"),
        TextColumn("[bold cyan]{task.fields[task_description]}"),
    )
    tpl_type = tpl_suffix(tpl_fname)
    jinja_tpl = get_jinja2_template(tpl_fname, tpl_dir)

    with progress:
        task = progress.add_task(
            "",
            total=num_groups,
            progress_description="",
            task_description="Filename",
        )
        for g_exhibitors, g_theatres in grouped_df:
            count += 1
            exhibitor_data = extract_exhibitor_data(g_exhibitors, g_theatres)
            annexure = extract_annexure_data(g_theatres)

            pdf_fname = f"{
                fname_tpl.format(
                    count=count,
                    movie=exhibitor_data['movie'].lower(),
                    exhibitor=exhibitor_data['exhibitor']
                    .replace('/', '')
                    .replace(' ', '_')
                    .lower(),
                    release_date=exhibitor_data['release_date'],
                )
            }.pdf"
            progress.update(task, task_description=f"{pdf_fname}")
            flist.append(pdf_fname)

            md_html_mergefields(
                jinja_tpl,
                tpl_type,
                css,
                pdf_fname,
                distributor_data=distributor_data,
                exhibitor_data=exhibitor_data,
                annexure=annexure,
            )
            progress.advance(task)
    return flist


if __name__ == "__main__":
    t1 = time.perf_counter()
    title = "Generating Agreement Documents"
    headings = {
        "heading1": "Installing the Components and Requirements of the Program",
        "hading2": "Install Source Code of the Program",
    }
    tpl_fname = "README.md.jinja"
    basename, _ = splitext(tpl_fname)
    basename, suffix = splitext(basename)
    pdf_fname = f"{basename}.pdf"
    css = "agreement.css"

    print(f"Jinja2 template: {tpl_fname}")
    tpl = get_jinja2_template(tpl_fname, "")
    if suffix.lower() == ".md":
        md_content = tpl.render(title=title, headings=headings)
        html_content = mistune.html(md_content)
    elif suffix.lower() == ".html":
        html_content = tpl.render(title=title, headigns=headings)
    if html_content:
        print(f"Writing: {pdf_fname}")
        HTML(string=html_content).write_pdf(pdf_fname, stylesheets=[css])
    t2 = time.perf_counter()
    print(f"Total time: {t2 - t1:.4f}s")
