# Generating Agreement Documents from Templates

This program is designed to generate PDF Agreement Documents from a suitably prepared template file and populated with data read from Microsoft Excel. The data is suitably processed and inserted into the template file appropriately before converting it to a PDF file.

Template files can be prepared in one of the following formats:

1. Microsoft Word files containing MergeFields that will be replaced by data from Microsoft Excel files
2. Jinja template files in Markdown format
3. Jinja template files in HTML format

Template files in Microsoft Word format are populated with data, written to a Microsoft Word file and converted to a PDF file by the LibreOffice `soffice` program. LibreOffice is therefore a requirement for this to work.

Template files in Jinja Markdown format are populated with data and written to HTML format. Template files in Jinja HTML format are populated with data and written to HTML format. The HTML file so obtained is  and then converted to a PDF file using [Weasyprint](https://weasyprint.org/). Weasyprint depends on Pango and is therefore a requirement for this to work. Refer to Weasyprint documentation to learn how to [install Pango](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html) for your operating system.

## External Dependencies

The following external programs must be installed:

1. [`git`](https://git-scm.com) to clone the source code from GitHub
2. [Astral-sh `uv`](https://docs.astral.sh/uv/) to manage Python projects
3. [`LibreOffice`](https://www.libreoffice.org/) to convert Microsoft Word documents to PDF when using Microsoft Word templates
4. [`Pango`](https://github.com/GNOME/pango) to convert HTML to PDF when using Markdown or HTML Jinja templates

## Install `uv`: A Python Package and Project Management Tool

### About `uv`

[Astral-sh `uv`](https://docs.astral.sh/uv/) is a single tool that manages a host of tasks which were previously managed by individual tools. For example, it can list, download and manage different versions of Python on a PC and allow you to choose your required version of Python per project, similar to `pyenv`.

It can create a *Virtual Environment* similar to `venv` and use it to execute an application without having to explicitly activate it, although you can activate it if you prefer to.

It can install and manage Python packages similar to `pip`.

It can manage Python projects by creating and managing `pyproject.toml` project configuration file similar to `poetry`.

Here are the steps to carry out the same tasks as before, using `uv` instead of `pip`:

1. [Install `uv` from Astral-sh website](#install-uv).
3. [Clone the Source Code from GitHub](clone-the-source-code-from-github)
4. [Install required Python packages](#install-required-python-packages).
5. [Copy Data and Template Files to the Project Directory](copy-data-and-template-files-to-the-project-directory)
6. [Run the program](#run-the-program).

### Install `uv`

See the Astral-sh `uv`  [installation page](https://docs.astral.sh/uv/getting-started/installation/) for different ways to download and install `uv`. The simplest way is to Open `Windows Powershell` by typing `powershell` in the Windows search bar and clicking on `Windows Powershell`. If you don't see `Windows Powershell` listed, you can download it from Microsoft web page by searching using your favourite search engine.

Once `Windows Poweshell` is open, type the following command at the prompt:

```bash
>powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

This will download and install Astral-sh `uv` on your computer. To verify this, close the `Windows Powershell`, open the `Command Prompt` and type `uv -V` at the prompt. If installation is successfull, you will see the version of `uv` installed on your computer.

### Clone the Source Code from GitHub

Program source code and its project files can be cloned from GitHub. Open `Command prompt` and change over to the folder within which the source code is to be installed. The following command assumes that you have `git` installed. If not, head over to the [`git` download page](https://git-scm.com/doc) for instructions to download and install `git`.

Check to ensure that you are in the project folder and have **not** activate the *Virtual Environment*. Then type the following command to clone the source code repo from GitHub:

```bash
>git clone https://github.com/satish-annigeri/agreement_docs.git
```

This will create the folder `agreement_docs` and clone project files into it. The `agreement_docs` folder has the following contents:

```bash
agreement_docs/
├── agreement_template.html.jinja
├── agreement_template.md.jinja
├── docxmerge.py
├── htmlmerge.py
├── main.py
├── mergedata.py
├── pyproject.toml
├── style.css
├── utils.py
└── uv.lock
```

### Install required Python packages

Install the required packages with the command:

```bash
agreement_docs>uv sync
```

This will:

* Rread the `pyproject.toml` file and determine the Python packages required to be installed
* Create a *Virtual Environment* named `.venv`
* Download and install all required Python packages and their dependencies into the *Virtual Environment*

You can verify that all required packages are installed by listing all installed packages with the command `uv pip list`.

### Copy Data and Template Files to the Project Directory

With the project folder created and the *Virtual Environment* configured, it is now time to copy the Microsoft Word template file `agreement_template.docx`. This must be prepared using MergeFields.

The Jinja template files `agreement_template.md.jinja` and `agreement_template.html.jinja` have already been cloned from the GitHub repository. They are to be prepared based on Jinja2 templating language.

The Microsoft Excel data files `distributors.xlsx`, `exhibitors.xlsx` and a suitably prepared theatres file must be copied into the folder. Refer the sample theatre file `chhaava_theatres.xlsx` for guidance. 

### Run the program

Verify that you are in the project folder `agreement_docs` and the *Virtual Environment* is **not** activated. You can now run the program with the following command at the prompt:

```bash
agreement_docs>uv run -- python main.py run chhaava_theatres.xlsx
```

where `chhaava_theatres.xlsx` is the data file with the list of theatres for the movie **Chhaava**. This file could be replaced with another theatres list file.

### Preparing to Run the Program
Before running the program, check the following:

1. You have opened the Command Prompt in the folder where your program source code is, namely, `agreement_docs` within the `My Documents` folder.
2. You have activated the *Virtual Environment* with the command `>.venv\Scripts\activate`.
3. Check the versions and of `python.exe` and `pip.exe` and the path of `pip.exe`.
4. The following data files are available in the current folder: `agreement_template.docx` (or `agreement_template.md.jinja`), `distributors.xlsx`, `exhibitors.xlsx` and the `.xlsx` file containing the list of theatres for which the agreement documents are to be generated. The name of this file is your choice, but it is recommeneded that it begin with the name of the movie for which the agreement documents are being generated.

### Preparing Input Files

The program expects the following input files:

1. `agreement_template.docx` is a Microsoft Word file containing the text of the agreement document  with place holders which change for each exhibitor. This must be created once and will be used as a template for all the agreement documents. At present, the template is common for all exhibitors, but in the future, this program will be modified to use different template file for different exhibitors.
2. `distributors.xlsx` is a Microsoft Excel file with the column names as shown in the file sent. At present, it uses only the first row as the only distributors, irrespective of the number of rows in the file. In future, the program can be modified to handle multiple distributors, if necessary. This file need be changed only if any of the data about the distributor changes. Otherwise, this file remains fairly constant.
3. `exhibitors.xlsx` is a Microsoft Excel file with the column names as shown in the file sent. There is one row for each Exhibitor. No two rows of this file must be identical. New rows can be added to this file when a new exhibitor is appointed. Otherwise this file remains fairly constant.
4. `abcd_theatres.xlsx` is a Microsoft Excel file with the column names as shown in the file sent. The first part `abcd` of the filename may be changed to reflect the name of the movie, but that is only for our convenience and not a requirement. This file must be prepared afresh for each set of new agreement files to be generated. The fist column of this file must be the first few unique letters from the name of the exhibitor as defined in the `exhibitors.xlsx` file. Note that the case of the letters is unimportant, but spaces, punctuation etc. must be identical as in the `exhibitors.xlsx` file.
