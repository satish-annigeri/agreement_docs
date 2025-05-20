# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = "DocGen"
copyright = "2025, Satish Annigeri"
author = "Satish Annigeri"
release = "0.1.0"

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = []

templates_path = ["_templates"]
exclude_patterns = []


# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

# html_theme = 'alabaster'
html_theme = "sphinx_rtd_theme"
html_theme_options = {
    "collapse_navigation": False,
    "navigation_depth": 4,
    "sticky_navigation": True,
}
html_static_path = ["_static"]
extensions = ["sphinx.ext.autodoc", "sphinx.ext.napoleon"]
autodoc_default_options = {
    "members": True,
    "undoc-members": True,  # Include undocumented members
}

napoleon_google_docstring = True
napoleon_numpy_docstring = False
autoapi_root = "docs/source/api"
