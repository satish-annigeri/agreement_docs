from os.path import splitext


def with_suffix(fpath: str, suffix: str) -> str:
    base, _ = splitext(fpath)
    return base + suffix


def tpl_suffix(tpl_fname: str):
    basename, suffix = splitext(tpl_fname)
    if suffix == ".jinja":
        basename, suffix = splitext(basename)
    return suffix.replace(".", "").lower()
