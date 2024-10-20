import toml
from pdf_preview import __version__

def get_project_version():
    pyproject_file = 'pyproject.toml'
    pyproject_data = toml.load(pyproject_file)
    return pyproject_data['tool']['poetry']['version']

def test_version():
    project_version = get_project_version()
    assert __version__ == project_version
