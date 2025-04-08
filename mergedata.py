from os.path import isfile

import polars as pl
import pendulum
from rich.console import Console


con = Console()


# ---- General purpose functions ----


def fmt_indian(n: float, trunc: bool = True):
    if not n:
        return "Rs. -NIL-"
    s = str(int(float(n))) if trunc else str(n)
    d = [3, 2, 2, 2]
    i = 0
    c = []
    i2 = len(s)
    while i < len(d):
        i1 = i2 - d[i]
        if i1 < 0:
            i1 = 0
        c.append(s[i1:i2])
        i += 1
        i2 = i1
    c = [cc for cc in c if cc]
    c = ",".join(c[::-1])
    if trunc:
        c += "/-"
    c = "Rs. " + c
    return c


# ---- Functions for preparing data to be merged into docx MergeFields ----


def read_excel(fname: str) -> pl.DataFrame:
    if isfile(fname):
        df = pl.read_excel(fname)
        return df
    else:
        raise FileNotFoundError(f"{fname}: File not found")


def read_csv(fname: str) -> pl.DataFrame:
    if isfile(fname):
        df = pl.read_excel(fname)
        return df
    else:
        raise FileNotFoundError(f"{fname}: File not found")


def read_data(
    distributor_fname: str,
    exhibitor_fname: str,
    theatre_fname: str,
    verbose: bool = False,
):
    # Read data
    if verbose:
        print(f"Reading: {distributor_fname}")
    distributors = read_excel(distributor_fname)
    if verbose:
        print(f"Reading: {exhibitor_fname}")
    exhibitors = read_excel(exhibitor_fname)
    if verbose:
        print(f"Reading: {theatre_fname}")
    theatres = read_excel(theatre_fname)

    return distributors, exhibitors, theatres


def clean_distributors_data(df: pl.DataFrame) -> pl.DataFrame:
    return df[0]


def clean_exhibitors_data(df: pl.DataFrame) -> pl.DataFrame:
    return df.with_columns(
        pl.col("exhibitor").str.to_uppercase(),
        pl.col("exhibitor_place").str.to_uppercase(),
    )


def clean_theatres_data(df: pl.DataFrame) -> pl.DataFrame:
    df = df.with_columns(
        pl.col("theatre").str.to_uppercase(),
        pl.col("station").str.to_uppercase(),
        pl.col("release_date")
        .map_elements(
            lambda dt: pendulum.instance(dt).format("Do [day of] MMMM, YYYY"),
            return_dtype=pl.String,
        )
        .alias("release_date_long"),
        pl.col("agreement_date")
        .map_elements(
            lambda dt: pendulum.instance(dt).format("Do [day of] MMMM, YYYY"),
            return_dtype=pl.String,
        )
        .alias("agreement_date_long"),
        pl.col("mg").map_elements(fmt_indian, return_dtype=pl.String).alias("mg_str"),
        pl.col("theatre_share").fill_null("-Nil-"),
        pl.col("release_date").dt.strftime("%d-%m-%Y"),
        pl.col("agreement_date").dt.strftime("%d-%m-%Y"),
    )
    return df


def clean_data(distributors, exhibitors, theatres):
    distributors = clean_distributors_data(distributors)
    exhibitors = clean_exhibitors_data(exhibitors)
    theatres = clean_theatres_data(theatres)
    return distributors, exhibitors, theatres


def join_data(exhibitors: pl.DataFrame, theatres: pl.DataFrame) -> pl.DataFrame:
    exhibitors = exhibitors.with_columns(
        pl.col("exhibitor").str.to_lowercase().alias("ex_b"),
    )
    theatres = theatres.with_columns(
        pl.col("exhibitor")
        .str.to_lowercase()
        .map_elements(
            lambda x: next((ex for ex in exhibitors["ex_b"] if ex.startswith(x)), None),
            return_dtype=pl.String,
        )
        .alias("ex_a")
    )
    df = theatres.join(exhibitors, left_on="ex_a", right_on="ex_b", how="inner")
    df = df.drop("exhibitor", "ex_a")
    df = df.rename({"exhibitor_right": "exhibitor"})
    return df


def prepare_data(distributors, exhibitors, theatres):
    distributors, exhibitors, theatres = clean_data(distributors, exhibitors, theatres)
    df = join_data(exhibitors, theatres)
    return distributors, df


def group_data(df: pl.DataFrame, group_cols: list[str]):
    return df.group_by(group_cols, maintain_order=True)


def unique_rows(df: pl.DataFrame, cols: str | list[str]) -> int:
    return df.select(cols).unique().height


def extract_distributor_data(distributors):
    return distributors.to_dicts()[0]


def extract_exhibitor_data(key, group):
    advance_amt = group["advance_amt"].sum()
    data = {
        "exhibitor": key[0],
        "exhibitor_place": key[1],
        "movie": key[2],
        "release_date": key[3],
        "agreement_date": key[4],
        "release_date_long": group.select("release_date_long")[0, 0],
        "agreement_date_long": group.select("agreement_date_long")[0, 0],
        "advance_amt": fmt_indian(advance_amt),
        "exhibitor_gst": group.select("exhibitor_gst")[0, 0],
        "movie_description": group.select("movie_description")[0, 0],
        "daily_shows": group.select("daily_shows")[0, 0],
    }
    return data


def extract_annexure_data(g_theatres):
    annexure = g_theatres.sort("station", "theatre").select(
        "theatre", "station", "mg_str", "theatre_share"
    )
    annexure = annexure.with_columns(pl.arange(1, annexure.height + 1).alias("slno"))
    return annexure.to_dicts()


if __name__ == "__main__":
    distributors_fname = "distributors.xlsx"
    exhibitors_fname = "exhibitors.xlsx"
    theatres_fname = "chhaava_theatres.xlsx"
    try:
        distributors = read_excel(distributors_fname)
        exhibitors = read_excel(exhibitors_fname)
        theatres = read_excel(theatres_fname)
        # print(f"Distributors: {distributors.height}")
        # print(f"Exhibitors: {exhibitors.height}")
        # print(f"Theatres: {theatres.height}")
        distributors = clean_distributors_data(distributors)
        exhibitors = clean_exhibitors_data(exhibitors)
        theatres = clean_theatres_data(theatres)
        # print(distributors)
        # print(exhibitors)
        # print(
        #     theatres["theatre", "station", "release_date_long", "agreement_date_long"]
        # )
        df = join_data(exhibitors, theatres)
        # print(df["exhibitor", "theatre", "station", "release_date", "agreement_date"])
        group_cols = [
            "exhibitor",
            "exhibitor_place",
            "movie",
            "release_date",
            "agreement_date",
        ]

        num_exhibitors = df.select(group_cols).unique()
        print(f"Number of files: {num_exhibitors.height}")
        distributor_data = extract_distributor_data(distributors)
        # con.print(distributor_data)
        grouped_df = group_data(df, group_cols)
        for g_exhibitor, g_theatres in grouped_df:
            exhibitor_data = extract_exhibitor_data(g_exhibitor, g_theatres)
            # con.print(exhibitor_data)
            annexure = extract_annexure_data(g_theatres)
            # print(annexure)
            # print()
    except Exception as e:
        print(f"{e}")
