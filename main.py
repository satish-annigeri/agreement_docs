import sys
import os
import time
from typing_extensions import Annotated


from rich.console import Console
from rich.progress import (
    Progress,
    TextColumn,
    SpinnerColumn,
    MofNCompleteColumn,
    TaskProgressColumn,
    TimeElapsedColumn,
)
import typer


from utils import tpl_suffix, get_fname, with_suffix
from mergedata import (
    read_data,
    prepare_data,
    group_data,
    unique_rows,
    extract_distributor_data,
    extract_exhibitor_data,
    extract_annexure_data,
)
from docxmerge import (
    docx_mergefields,
    print_header,
    detect_soffice_path,
    soffice_docx2pdf,
)
from htmlmerge import md_html_mergefields, get_jinja2_template


app = typer.Typer()
con = Console()


def main(
    theatre: Annotated[str, typer.Argument(help="Theatres data in .xlsx format")],
    template: Annotated[
        str,
        typer.Option(
            "--template",
            "-t",
            help="Template file in .docx, .md.jinja or .html.jinja format",
        ),
    ] = "agreement_template.docx",
    css: Annotated[
        str,
        typer.Option(
            "--css",
            "-c",
            help="CSS stylesheet file for Markdown and HTML template files",
        ),
    ] = "agreement.css",
    distributor: Annotated[
        str,
        typer.Option("--distributor", "-d", help="Distributor data in .xlsx format"),
    ] = "distributors.xlsx",
    exhibitor: Annotated[
        str, typer.Option("--exhibitor", "-e", help="Exhibitors data in .xlsx format")
    ] = "exhibitors.xlsx",
):
    t_start = time.perf_counter()
    t1 = t_start
    print_header("Preparing Agreement Documents")

    distributor_fname = distributor
    exhibitor_fname = exhibitor
    theatre_fname = theatre
    template_fname = template
    css_fname = css

    distributors, exhibitors, theatres = read_data(
        distributor_fname, exhibitor_fname, theatre_fname, verbose=True
    )
    tpl_type = tpl_suffix(template_fname)
    if tpl_type in ["md", "html"]:
        print(f"Stylesheet: {css}")
    distributors, df = prepare_data(distributors, exhibitors, theatres)

    group_cols = [
        "exhibitor",
        "exhibitor_place",
        "movie",
        "release_date",
        "agreement_date",
    ]
    grouped_df = group_data(df, group_cols)
    num_groups = unique_rows(df, group_cols)
    t2 = time.perf_counter()
    con.log(f"Data preparation complete {t2 - t1:.2f}s")
    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"

    distributor_data = extract_distributor_data(distributors)

    count = 0
    if tpl_type in ["md", "html"]:
        jinja_template = get_jinja2_template(template_fname, "")
    elif tpl_type == "docx":
        soffice_path, cmd_list, shell = detect_soffice_path()
    else:
        print(
            f"Unknown template type: {tpl_type}. Supported types are: md, html, docx\nProgram aborted"
        )
        sys.exit(1)

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
            total=num_groups,
            progress_description="Generating",
            task_description="",
        )
        # --------------------------
        for g_exhibitor, g_theatres in grouped_df:
            count += 1
            exhibitor_data = extract_exhibitor_data(g_exhibitor, g_theatres)
            annexure = extract_annexure_data(g_theatres)

            output_fname = get_fname(
                fname_tpl,
                count=count,
                movie=exhibitor_data["movie"].lower(),
                exhibitor=exhibitor_data["exhibitor"],
                release_date=exhibitor_data["release_date"],
            )

            if tpl_type == "docx":
                output_fname = f"{output_fname}.docx"
                progress.update(
                    task, task_description=f"=== {with_suffix(output_fname, '.pdf')}"
                )
                docx_mergefields(
                    template_fname,
                    output_fname,
                    distributor_data,
                    exhibitor_data,
                    annexure,
                )
                soffice_docx2pdf(output_fname, cmd_list, shell)
                os.remove(output_fname)
            elif tpl_type in ["html", "md"]:
                output_fname = f"{output_fname}.pdf"
                progress.update(task, task_description=f"{output_fname}")
                md_html_mergefields(
                    jinja_template,
                    tpl_type,
                    css_fname,
                    output_fname,
                    distributor_data=distributor_data,
                    exhibitor_data=exhibitor_data,
                    annexure=annexure,
                )
            else:
                con.print("Unknown template type. Exiting...")
                sys.exit(1)
            progress.advance(task)
        t3 = time.perf_counter()
    # --------------------------

    t_stop = t3
    t_total = t_stop - t_start
    con.print(
        f"\nTotal execution time: {t_total:.2f}s for {num_groups} files. Average: {t_total / num_groups:.2f}s per file."
    )


if __name__ == "__main__":
    typer.run(main)
