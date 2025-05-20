Usage
-----

DocGen is a command line interface (CLI) tool that generates PDF files from templates defined either in Jinja2 HTML files or Microsoft Word files with MergeFields. The data to populate the 
templates is generated from Microsoft Excel files. The current application is specifically designed to generate Agreement Documents for a business house, but the portion of the application that populates the templates with the prepared data is generic and can be used to generate other documents as long as the data is prepared appropriate to the template.

Preparing the Template file
~~~~~~~~~~~~~~~~~~~~~~~~~~~

The template file can be in either of the following formats:

1. Microsoft Word file with MergeFields
2. Jinja2 HTML file
3. Jinja2 Markdown file

Except for the MergeFields, rest of the template is not affected by this application. The MergeFields in a Microsoft Word file are inserted using the `Insert > Quick Parts > Fields` option in Microsoft Word. The MergeFields are replaced with the data from the Excel files when the document is generated. The Jinja2 HTML and Markdown templates are processed using the Jinja2 templating engine, which allows for more complex logic and formatting. See `Jinja2 <https://jinja.palletsprojects.com/en/stable/>`_ for more information on how to use Jinja2 templates.

In case of templates in Jinja2 HTML or Markdown format, the CSS stylesheet file is used to style the document. The default CSS stylesheet filename is *agreement.css*. If you use a different filename, it must be furinished as an option to the CLI. The CSS file is not required for Microsoft Word templates.

Preparing the Data files
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The data files are in Microsoft Excel format and are used to populate the template files. The data files are:

1. Theatres data file
2. Distributors data file
3. Exhibitors data file

Usage
~~~~~

The command line interface (CLI) is used to generate the PDF files from the template files. The CLI takes the following arguments:

1. Theatres data filename (required)
2. Template filename (assumed to be *agreement_template.docx* if not provided)
3. CSS stylesheet file (assumed to be *agreement.css* if not provided)
4. Distributors data file (assumed to be *distributors.xlsx* if not provided)
5. Exhibitors data file (assumed to be *exhibitors.xlsx* if not provided)

.. code-block:: shell

    Usage: main.py [OPTIONS] THEATRE

    Arguments
    theatre      TEXT  Theatres data in .xlsx format [default: None] [required]

    --template     -t      TEXT  Template file in .docx, .md.jinja or .html.jinja format [default: agreement_template.docx]
    --css          -c      TEXT  CSS stylesheet file for Markdown and HTML template files [default: agreement.css]
    --distributor  -d      TEXT  Distributor data in .xlsx format [default: distributors.xlsx]
    --exhibitor    -e      TEXT  Exhibitors data in .xlsx format [default: exhibitors.xlsx]
    --help                       Show this message and exit.