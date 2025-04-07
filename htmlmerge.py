import re
import time
from os.path import abspath, splitext
from rich.console import Console

from jinja2 import Environment, FileSystemLoader
import mistune
from weasyprint import HTML

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


# def mdmerge(tpl_fname: str, tpl_dir: str = ".", **kwargs):
#     if isfile(tpl_fname):
#         tpl = get_jinja2_template(tpl_fname, tpl_dir)
#         return tpl.render(**kwargs)
#     else:
#         raise FileNotFoundError


# def htmlmerge(template_fname: str, tpl_dir: str = "", **kwargs):
#     if isfile(tpl_fname):
#         tpl = get_jinja2_template(tpl_fname, tpl_dir)
#         return tpl.render(**kwargs)
#     else:
#         raise FileNotFoundError


def md2html(md_fname: str) -> str:
    with open(md_fname, "r", encoding="utf-8") as md_file:
        md_content = md_file.read()
    html = mistune.html(md_content)

    if isinstance(html, str):
        return html
    else:
        return ""


def md_html_mergefields(
    tpl_fname: str,
    tpl_dir: str,
    css: str,
    pdf_fname: str,
    distributor_data,
    exhibitor_data,
    annexure,
):
    tpl_type = tpl_suffix(tpl_fname)
    if tpl_type in ["md", "html"]:
        tpl = get_jinja2_template(tpl_fname, tpl_dir)
    if tpl_type == "md":
        md_content = tpl.render(
            **distributor_data,
            **exhibitor_data,
            annexure=annexure,
        )
        html_content = mistune.html(md_content)
    elif tpl_type == "html":
        html_content = tpl.render(
            distributor_data=distributor_data,
            exhibitor_data=exhibitor_data,
            annexure=annexure,
        )
    HTML(string=html_content).write_pdf(pdf_fname, stylesheets=[css])


def md_html_merge(distributors, grouped_df, tpl_fname, tpl_dir, css, fname_tpl):
    distributor_data = extract_distributor_data(distributors)

    count = 0
    flist = []
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
        print(pdf_fname)
        flist.append(pdf_fname)
        # con.print(distributor_data)
        # con.print(exhibitor_data)
        # con.print(annexure)
        md_html_mergefields(
            tpl_fname,
            tpl_dir,
            css,
            pdf_fname,
            distributor_data=distributor_data,
            exhibitor_data=exhibitor_data,
            annexure=annexure,
        )
    return flist


# def html2pdf(input_html: str, stylesheet: str, pdf_fname: str):
#     if is_html_fname(input_html):
#         with open(input_html, "r") as f:
#             input_html = f.read()
#     HTML(string=input_html).write_pdf(pdf_fname, stylesheets=[stylesheet])


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
    css = "style.css"

    print(f"Jinja2 template: {tpl_fname}")
    jinja_tpl = get_jinja2_template(tpl_fname, "")
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
