import marimo

__generated_with = "0.12.8"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def _():
    from platform import system
    platform = system()
    return platform, system


@app.cell
def _():
    from rich.console import Console
    import polars as pl
    from jinja2 import Environment, FileSystemLoader

    con = Console()
    env = Environment(loader=FileSystemLoader(""))
    return Console, Environment, FileSystemLoader, con, env, pl


@app.cell
def _():
    from mergedata import (read_data, prepare_data, group_data, unique_rows, extract_distributor_data, extract_exhibitor_data, extract_annexure_data)
    return (
        extract_annexure_data,
        extract_distributor_data,
        extract_exhibitor_data,
        group_data,
        prepare_data,
        read_data,
        unique_rows,
    )


@app.cell
def _():
    from htmlmerge import md_html_mergefields
    return (md_html_mergefields,)


@app.cell
def _():
    from utils import tpl_suffix
    return (tpl_suffix,)


@app.cell
def _(mo):
    d_file = mo.ui.file_browser(filetypes=['.xlsx'], multiple=False)
    d_file
    return (d_file,)


@app.cell
def _(mo):
    e_file = mo.ui.file_browser(filetypes=['.xlsx'], multiple=False)
    e_file
    return (e_file,)


@app.cell
def _(mo):
    t_file = mo.ui.file_browser(filetypes=['.xlsx'], multiple=False)
    t_file
    return (t_file,)


@app.cell
def _(d_file, e_file, t_file, tpl_suffix):
    distributors_fname = d_file.path()
    exhibitors_fname = e_file.path()
    theatres_fname = t_file.path()
    tpl_fname = "agreement.html.jinja"
    tpl_type = tpl_suffix(tpl_fname)
    print(f"Template type: {tpl_type}")
    return (
        distributors_fname,
        exhibitors_fname,
        theatres_fname,
        tpl_fname,
        tpl_type,
    )


@app.cell
def _(distributors_fname, exhibitors_fname, read_data, theatres_fname):
    distributors, exhibitors, theatres = read_data(distributors_fname, exhibitors_fname, theatres_fname, verbose=False)
    return distributors, exhibitors, theatres


@app.cell
def _(
    con,
    distributors,
    exhibitors,
    extract_distributor_data,
    group_data,
    prepare_data,
    theatres,
    unique_rows,
):
    distributors_df, df = prepare_data(distributors, exhibitors, theatres)
    distributor_data = extract_distributor_data(distributors_df)

    group_cols = [
        "exhibitor",
        "exhibitor_place",
        "movie",
        "release_date",
        "agreement_date",
    ]
    grouped_df = group_data(df, group_cols)
    num_groups = unique_rows(df, group_cols)
    print(f"Number of groups: {num_groups}")
    con.print("Distributor data:")
    con.print(distributor_data)
    return (
        df,
        distributor_data,
        distributors_df,
        group_cols,
        grouped_df,
        num_groups,
    )


@app.cell
def _(
    distributor_data,
    env,
    extract_annexure_data,
    extract_exhibitor_data,
    grouped_df,
    md_html_mergefields,
    tpl_fname,
    tpl_type,
):
    count = 0
    flist = []
    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"


    if tpl_type in ['html', 'md']:
        jinja_tpl = env.get_template(tpl_fname)
        css_fname = "agreement.css"
    
    for g_exhibitors, g_theatres in grouped_df:
        exhibitor_data = extract_exhibitor_data(g_exhibitors, g_theatres)
        annexure = extract_annexure_data(g_theatres)
        count += 1
        output_fname = fname_tpl.format(
            count=count,
            movie=exhibitor_data["movie"].lower(),
            exhibitor=exhibitor_data["exhibitor"].replace("/", "").replace(" ", "_").lower(),
            release_date=exhibitor_data['release_date']
        )
        match tpl_type:
            case 'docx':
                print(f"Microsoft Word: {output_fname}.docx")
            case 'html' | 'md':
                print(f"HTML/Markdown: {output_fname}.pdf")
                md_html_mergefields(
                    jinja_tpl,
                    tpl_type,
                    css_fname,
                    f"{output_fname}.pdf",
                    distributor_data=distributor_data,
                    exhibitor_data=exhibitor_data,
                    annexure=annexure,
                )
            case _:
                print('Unknown template type')
        flist.append(output_fname)
    return (
        annexure,
        count,
        css_fname,
        exhibitor_data,
        flist,
        fname_tpl,
        g_exhibitors,
        g_theatres,
        jinja_tpl,
        output_fname,
    )


if __name__ == "__main__":
    app.run()
