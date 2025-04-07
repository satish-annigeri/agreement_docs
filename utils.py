from os.path import splitext


def tpl_suffix(tpl_fname: str):
    basename, suffix = splitext(tpl_fname)
    if suffix == ".jinja":
        basename, suffix = splitext(basename)
    return suffix.replace(".", "").lower()
