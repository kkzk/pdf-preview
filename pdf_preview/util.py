from pathlib import Path


def get_pdfjs():
    if not Path("pdfjs-dist.zip").exists():
        import urllib.request
        url = "https://github.com/mozilla/pdf.js/releases/download/v2.2.228/pdfjs-2.2.228-dist.zip"
        with urllib.request.urlopen(url) as response:
            open("pdfjs-dist.zip", "wb").write(response.read())
    if not Path("pdfjs-dist/web/viewer.html").exists():
        import zipfile
        with zipfile.ZipFile("pdfjs-dist.zip", "r") as existing_zip:
            existing_zip.extractall("pdfjs-dist")

