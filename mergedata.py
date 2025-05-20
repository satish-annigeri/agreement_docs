from os.path import isfile
from math import isclose
import polars as pl
import pendulum
from rich.console import Console


con = Console()


# ---- General purpose functions ----


def fmt_indian(n: float, currency: str = "₹", trunc: bool = True):
    """
    Format a number in Indian currency format.

    Args:
        n (float): The number to format.
        currency (str): The currency symbol to use. Default is "₹".
        trunc (bool): If True, truncate the number to an integer. Default is True.

    Returns:
        str: The formatted number as a string.
    """
    if not n or isclose(n, 0.0):
        c = "-NIL-"
    else:
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
    c = f"{currency} {c}"
    return c


# ---- Functions for preparing data to be merged into docx MergeFields ----


def read_excel(fname: str) -> pl.DataFrame:
    """
    Read an Excel file and return a DataFrame.

    Args:
        fname (str): The path to the Excel file.

    Returns:
        pl.DataFrame: A DataFrame containing the data from the Excel file.

    Raises:
        FileNotFoundError: If the file does not exist."""
    if isfile(fname):
        df = pl.read_excel(fname)
        return df
    else:
        raise FileNotFoundError(f"{fname}: File not found")


def read_csv(fname: str) -> pl.DataFrame:
    """
    Read a CSV file and return a DataFrame.

    Args:
        fname (str): The path to the CSV file.

    Returns:
        pl.DataFrame: A DataFrame containing the data from the CSV file.

    Raises:
        FileNotFoundError: If the file does not exist."""
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
    """
    Read data from Excel files and return DataFrames.

    Args:
        distributor_fname (str): The path to the distributor Excel file.
        exhibitor_fname (str): The path to the exhibitor Excel file.
        theatre_fname (str): The path to the theatre Excel file.
        verbose (bool): If True, print the names of the files being read.
            Default is False.

    Returns:
        distributors (pl.DataFrame): A DataFrame containing distributor data.
        exhibitors (pl.DataFrame): A DataFrame containing exhibitor data.
        theatres (pl.DataFrame): A DataFrame containing theatre data.

    Raises:
        FileNotFoundError: If any of the files do not exist.
    """
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
    """
    Returns all the columns of the initial row of the DataFrame,
        assuming it is the only one we want..

    Args:
        df (pl.DataFrame): The DataFrame containing distributor data.

    Returns:
        pl.DataFrame: The cleaned DataFrame with uppercase columns.
    """
    return df[0]


def clean_exhibitors_data(df: pl.DataFrame) -> pl.DataFrame:
    """
    Clean the exhibitor data by converting data in columns to uppercase.

    Args:
        df (pl.DataFrame): The DataFrame containing exhibitor data.

    Returns:
        pl.DataFrame: The cleaned DataFrame with uppercase columns.
    """
    return df.with_columns(
        pl.col("exhibitor").str.to_uppercase(),
        pl.col("exhibitor_place").str.to_uppercase(),
    )


def clean_theatres_data(df: pl.DataFrame) -> pl.DataFrame:
    """
    Clean the theatre data by converting data in columns to uppercase, and long format of date.

    Args:
        df (pl.DataFrame): The DataFrame containing theatre data.

    Returns:
        pl.DataFrame: The cleaned DataFrame with uppercase columns and formatted dates.
    """
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


def clean_data(
    distributors: pl.DataFrame, exhibitors: pl.DataFrame, theatres: pl.DataFrame
):
    """
    Clean the data by applying the appropriate cleaning functions to each DataFrame.

    Args:
        distributors (pl.DataFrame): The DataFrame containing distributor data.
        exhibitors (pl.DataFrame): The DataFrame containing exhibitor data.
        theatres (pl.DataFrame): The DataFrame containing theatre data.

    Returns:
        distributors (pl.DataFrame): The cleaned DataFrame with distributor data.
        exhibitors (pl.DataFrame): The cleaned DataFrame with exhibitor data.
        theatres (pl.DataFrame): The cleaned DataFrame with theatre data.
    """
    distributors = clean_distributors_data(distributors)
    exhibitors = clean_exhibitors_data(exhibitors)
    theatres = clean_theatres_data(theatres)
    return distributors, exhibitors, theatres


def join_data(exhibitors: pl.DataFrame, theatres: pl.DataFrame) -> pl.DataFrame:
    """
    Join the exhibitors and theatres DataFrames on the exhibitor column.

    Args:
        exhibitors (pl.DataFrame): The DataFrame containing exhibitor data.
        theatres (pl.DataFrame): The DataFrame containing theatre data.

    Returns:
        pl.DataFrame: A DataFrame containing the joined data.
    """
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


def prepare_data(
    distributors: pl.DataFrame, exhibitors: pl.DataFrame, theatres: pl.DataFrame
):
    """
    Prepare the data by cleaning and joining the DataFrames.

    Args:
        distributors (pl.DataFrame): The DataFrame containing distributor data.
        exhibitors (pl.DataFrame): The DataFrame containing exhibitor data.
        theatres (pl.DataFrame): The DataFrame containing theatre data.

    Returns:
        distributors (pl.DataFrame): The cleaned DataFrame with distributor data.
        df (pl.DataFrame): A DataFrame containing the joined data.
    """
    distributors, exhibitors, theatres = clean_data(distributors, exhibitors, theatres)
    df = join_data(exhibitors, theatres)
    return distributors, df


def group_data(df: pl.DataFrame, group_cols: list[str]):
    """
    Group the DataFrame by the specified columns and maintain the order.

    Args:
        df (pl.DataFrame): The DataFrame to group.
        group_cols (list[str]): The columns to group by.

    Returns:
        pl.DataFrame: A DataFrame with the grouped data."""
    return df.group_by(group_cols, maintain_order=True)


def unique_rows(df: pl.DataFrame, cols: str | list[str]) -> int:
    """
    Count the number of unique rows in the DataFrame based on the specified columns.

    Args:
        df (pl.DataFrame): The DataFrame to count unique rows in.
        cols (str | list[str]): The columns to consider for uniqueness.

    Returns:
        int: The number of unique rows in the DataFrame.
    """
    return df.select(cols).unique().height


def extract_distributor_data(distributors) -> list[dict[str, str]]:
    """
    Extract distributor data from the DataFrame.

    Args:
        distributors (pl.DataFrame): The DataFrame containing distributor data.

    Returns:
        list[dict[str, str]]: A list of dictionaries containing distributor data.
    """
    return distributors.to_dicts()[0]


def extract_exhibitor_data(key, group) -> dict[str, str]:
    """
    Extract exhibitor data from the grouped DataFrame.

    Args:
        key (tuple): The key for the group.
        group (pl.DataFrame): The grouped DataFrame.

    Returns:
        dict[str, str]: A dictionary containing exhibitor data.
    """
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


def extract_annexure_data(g_theatres) -> list[dict[str, str]]:
    """
    Extract annexure data from the grouped DataFrame to fill the table of annexures in the template.

    Args:
        g_theatres (pl.DataFrame): The grouped DataFrame containing theatre data.

    Returns:
        list[dict[str, str]]: A list of dictionaries containing annexure data.
    """
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

        df = join_data(exhibitors, theatres)

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

        grouped_df = group_data(df, group_cols)
        for g_exhibitor, g_theatres in grouped_df:
            exhibitor_data = extract_exhibitor_data(g_exhibitor, g_theatres)

            annexure = extract_annexure_data(g_theatres)

    except Exception as e:
        print(f"{e}")
