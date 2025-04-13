import re
import time
from os.path import abspath, splitext


from rich.console import Console

# from rich.progress import (
#     Progress,
#     TextColumn,
#     SpinnerColumn,
#     MofNCompleteColumn,
#     TaskProgressColumn,
#     TimeElapsedColumn,
# )
import pendulum
from jinja2 import Environment, FileSystemLoader
import mistune
from weasyprint import HTML, __version__ as wezp_ver
from weasyprint.text.fonts import FontConfiguration


re_html_fname = re.compile(r".*[.]html$", re.I)

font_config = FontConfiguration()
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
    css_fname: str,
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
    HTML(string=html_content).write_pdf(
        pdf_fname, stylesheets=[css_fname], font_config=font_config
    )


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
    # css = "agreement.css"

    print(f"Jinja2 template: {tpl_fname}")
    tpl = get_jinja2_template(tpl_fname, "")
    if suffix.lower() == ".md":
        md_content = tpl.render(title=title, headings=headings)
        html_content = mistune.html(md_content)
    elif suffix.lower() == ".html":
        html_content = tpl.render(title=title, headigns=headings)
    if html_content:
        print(f"Writing: {pdf_fname}")
        HTML(string=html_content).write_pdf(
            pdf_fname, stylesheets=["agreement.css"], font_config=font_config
        )
    t2 = time.perf_counter()
    print(f"Total time: {t2 - t1:.4f}s")
