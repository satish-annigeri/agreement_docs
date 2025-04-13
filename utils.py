from os.path import splitext


def with_suffix(fpath: str, suffix: str) -> str:
    base, _ = splitext(fpath)
    return base + suffix


def tpl_suffix(tpl_fname: str):
    basename, suffix = splitext(tpl_fname)
    if suffix == ".jinja":
        basename, suffix = splitext(basename)
    return suffix.replace(".", "").lower()


def get_fname(fname_tpl: str, **kwargs) -> str:
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
