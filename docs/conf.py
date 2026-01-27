# Configuration file for the Sphinx documentation builder.

project = "genro-office"
copyright = "2025, Softwell S.r.l."
author = "Genropy Team"
release = "0.1.0"

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx.ext.viewcode",
    "sphinx.ext.intersphinx",
    "myst_parser",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]

html_theme = "furo"
html_static_path = ["_static"]

# Napoleon settings
napoleon_google_docstring = True
napoleon_numpy_docstring = False
napoleon_include_init_with_doc = True

# Intersphinx
intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
}

# MyST settings
myst_enable_extensions = [
    "colon_fence",
    "deflist",
]

# Autodoc settings
autodoc_member_order = "bysource"
autodoc_typehints = "description"
