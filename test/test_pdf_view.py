import os

from pypdf import PdfWriter
from pytestqt.qtbot import QtBot

from pdf_preview.pdf_view import PdfViewer


def write_pdf(path, pages):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=595, height=842)
    writer.write(path)


def make_viewer(qtbot: QtBot):
    viewer = PdfViewer()
    qtbot.addWidget(viewer)
    viewer.resize(800, 600)
    viewer.show()
    qtbot.waitExposed(viewer)
    return viewer


def test_file_is_not_locked(qtbot: QtBot, tmp_path):
    pdf = tmp_path / "out.pdf"
    write_pdf(pdf, 2)
    viewer = make_viewer(qtbot)
    viewer.reload(pdf)
    assert viewer.document.pageCount() == 2

    # 表示中でも置き換え・削除ができる
    other = tmp_path / "other.pdf"
    write_pdf(other, 3)
    os.replace(other, pdf)
    os.remove(pdf)
    assert viewer.document.pageCount() == 2


def test_reload_keeps_position(qtbot: QtBot, tmp_path):
    pdf = tmp_path / "out.pdf"
    write_pdf(pdf, 6)
    viewer = make_viewer(qtbot)
    viewer.reload(pdf)
    bar = viewer.view.verticalScrollBar()
    qtbot.waitUntil(lambda: bar.maximum() > 0)

    bar.setValue(bar.maximum() // 2 + 137)
    position = bar.value()
    page = viewer.view.pageNavigator().currentPage()
    assert page > 0

    write_pdf(pdf, 7)
    viewer.reload(pdf)
    qtbot.waitUntil(lambda: viewer.page_count_label.text() == " / 7 ")
    qtbot.waitUntil(lambda: bar.value() == position)
    assert viewer.view.pageNavigator().currentPage() == page
    assert viewer.page_spin.value() == page + 1


def test_reload_missing_file_keeps_document(qtbot: QtBot, tmp_path):
    pdf = tmp_path / "out.pdf"
    write_pdf(pdf, 3)
    viewer = make_viewer(qtbot)
    viewer.reload(pdf)
    viewer.reload(tmp_path / "missing.pdf")
    assert viewer.document.pageCount() == 3


def test_page_navigation_and_zoom(qtbot: QtBot, tmp_path):
    pdf = tmp_path / "out.pdf"
    write_pdf(pdf, 4)
    viewer = make_viewer(qtbot)
    viewer.reload(pdf)
    qtbot.waitUntil(lambda: viewer.view.verticalScrollBar().maximum() > 0)

    viewer.jump_to_page(3)
    qtbot.waitUntil(lambda: viewer.view.pageNavigator().currentPage() == 2)
    assert viewer.page_spin.value() == 3

    before = viewer.view.zoomFactor()
    viewer.zoom_in()
    assert viewer.view.zoomFactor() > before
