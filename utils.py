from os.path import splitext


def with_suffix(fpath: str, suffix: str) -> str:
    """Add suffix to file name.

    Args:
        fpath (str): File path.
        suffix (str): Suffix to add.

    Returns:
        str: File path with added suffix.
    """
    base, _ = splitext(fpath)
    return base + suffix


def tpl_suffix(tpl_fname: str):
    """Determine the type of template from the filename suffix. Ignore the .jinja at the end.

    Args:
        tpl_fname (str): Template file name.
    Returns:
        str: Template type (e.g., "html", "docx").
    """
    basename, suffix = splitext(tpl_fname)
    if suffix == ".jinja":
        basename, suffix = splitext(basename)
    return suffix.replace(".", "").lower()


def get_fname(fname_tpl: str, **kwargs) -> str:
    """Generate a file name based on a template and keyword arguments.

    Args:
        fname_tpl (str): File name template.
        **kwargs: Keyword arguments to fill in the template.

    Returns:
        str: Generated file name.
    """
    fname = fname_tpl.format(**kwargs).replace("/", "").replace(" ", "_").lower()
    return fname


if __name__ == "__main__":
    fname_tpl = "{count:02}_{movie}_{exhibitor}_{release_date}"
    print(
        get_fname(
            fname_tpl,
            count=1,
            movie="CHHAVA",
            exhibitor="M/s Jayanna Films",
            release_date="12-04-2025",
        )
    )
