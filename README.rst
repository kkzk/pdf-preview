pdf-preview
===========

Preview Excel/Word documents via PDF.

Install
-------
::

  uv tool install pdf-preview
  pdf-preview -i

or with pip::

  py -m pip install pdf-preview
  py -m pdf_preview -i

This command adds a registry KEY
``HKEY_CLASSES_ROOT\Directory\shell\pdf_Preview.convert\command``
and default VALUE ``<python executable> -m pdf_preview "%1"``.
You will see context menu for directory. Select it for execute PDF Preview application.
When you select a Word / Excel document in the tree, the result converted to pdf is displayed.

Requirements
------------

- Windows
- Python 3.12 or later
- Microsoft Excel / Word (used for PDF conversion)

Dependencies
------------

- pywin32
- pypdf
- PySide6
- openpyxl
- PyYAML

and their requirement packages.

Development
-----------
::

  uv sync
  uv run pytest -v
  uv build

Uses
----

- PDF.js 6.4.299 legacy build (included)
