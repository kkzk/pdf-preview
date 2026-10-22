import tomllib
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from pdf_preview import __version__
from pdf_preview.main_window import merge_pdfs

PYPROJECT = Path(__file__).parent.parent / 'pyproject.toml'


def get_project_version():
    with open(PYPROJECT, 'rb') as f:
        pyproject_data = tomllib.load(f)
    return pyproject_data['project']['version']


def test_version():
    project_version = get_project_version()
    assert __version__ == project_version


def test_merge_pdfs(tmp_path):
    sources = []
    for i in range(2):
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        path = tmp_path / f"src{i}.pdf"
        writer.write(path)
        sources.append(path)

    output = tmp_path / "merged.pdf"
    merge_pdfs(sources, output)

    assert len(PdfReader(output).pages) == 2
