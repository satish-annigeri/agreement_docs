Installation
--------------------------

This prpject uses `uv`_ to manage the project and its dependencies. See here for instructions to install `uv`_ in case you aren't using it already. If you wish to use only ``pip``, note that the package is not available on PyPI and you can install it from GitHub using ``pip`` like any other Python package from GitHub. However, be warae that you must have installed `git <https://git-scm.com/>`_ on your machine for this command to work.

Install Package from PyPI
~~~~~~~~~~~~~~~~~~~~~~~~~

To install the package, you can use pip. The package is available on PyPI and can be installed using the following command:

.. code-block:: shell

   $uv pip install git+https://github.com/satish-annigeri/agreement_docs.git

Install Source Code from GitHub
~~~~~~~~~~~~~~~~~~~~~

Alternatively, you can clone the repository and install it manually, and you will need both `uv`_ and `git`_. To do this, follow these steps:

1. Clone the repository
2. Navigate to the cloned directory
3. Install the package using ``pip``
4. Install the package in editable mode (optional)

   .. code-block:: text

      $git clone https://github.com/satish-annigeri/agreement_docs.git
      $cd agreement_docs
      $uv sync
      $uv pip install -e .
In case you wish to explore and modify the source code, and/or generate the documentation, install the developmwnt dependencies as well:

   .. code-block:: shell

      $uv sync --dev
This will install the development dependencies listed in the `pyproject.toml` file.

.. _uv: https://docs.astral.sh/uv/