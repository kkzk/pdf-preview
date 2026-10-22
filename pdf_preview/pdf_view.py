# -*- coding: utf-8 -*-
import logging
from pathlib import Path

from PySide6 import QtCore, QtWidgets
from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QTimer, Slot
from PySide6.QtGui import QKeySequence
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView

LOGGER = logging.getLogger(__name__)

ZOOM_STEP = 1.25


class PdfViewer(QtWidgets.QWidget):
    """PDF 表示ウィジェット

    PDF はメモリに読み込んでから表示するため、表示中もファイルをロックしない。
    reload() は表示位置とズームを維持したまま再表示する。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.document = QPdfDocument(self)
        self._buffer = None  # QPdfDocument が読む間は保持しておく必要がある

        self.view = QPdfView(self)
        self.view.setDocument(self.document)
        self.view.setPageMode(QPdfView.PageMode.MultiPage)
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)

        # ツールバー
        toolbar = QtWidgets.QToolBar(self)
        prev_action = toolbar.addAction("◀", self.prev_page)
        prev_action.setToolTip(self.tr("Previous page"))
        prev_action.setShortcut(QKeySequence(QKeySequence.StandardKey.MoveToPreviousPage))
        self.page_spin = QtWidgets.QSpinBox(self)
        self.page_spin.setMinimum(1)
        self.page_spin.setKeyboardTracking(False)
        self.page_spin.valueChanged.connect(self.jump_to_page)
        toolbar.addWidget(self.page_spin)
        self.page_count_label = QtWidgets.QLabel(" / 0 ", self)
        toolbar.addWidget(self.page_count_label)
        next_action = toolbar.addAction("▶", self.next_page)
        next_action.setToolTip(self.tr("Next page"))
        next_action.setShortcut(QKeySequence(QKeySequence.StandardKey.MoveToNextPage))
        toolbar.addSeparator()
        zoom_out_action = toolbar.addAction("−", self.zoom_out)
        zoom_out_action.setShortcut(QKeySequence(QKeySequence.StandardKey.ZoomOut))
        zoom_in_action = toolbar.addAction("+", self.zoom_in)
        zoom_in_action.setShortcut(QKeySequence(QKeySequence.StandardKey.ZoomIn))
        toolbar.addAction(self.tr("Fit width"), self.fit_width)
        toolbar.addAction(self.tr("Fit page"), self.fit_page)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(toolbar)
        layout.addWidget(self.view)

        self.view.pageNavigator().currentPageChanged.connect(self.on_current_page_changed)
        self.document.pageCountChanged.connect(self.on_page_count_changed)

    def reload(self, path: Path):
        """PDF を読み込み直す。表示位置は可能な限り維持する。"""
        try:
            data = Path(path).read_bytes()
        except OSError as e:
            LOGGER.warning("PDF を読み込めません {}: {}".format(path, e))
            return

        h_value = self.view.horizontalScrollBar().value()
        v_value = self.view.verticalScrollBar().value()

        buffer = QBuffer(self)
        buffer.setData(QByteArray(data))
        buffer.open(QIODevice.OpenModeFlag.ReadOnly)
        self.document.load(buffer)
        if self._buffer is not None:
            self._buffer.deleteLater()
        self._buffer = buffer

        # スクロール範囲は再レイアウト後に確定するので、イベントループに戻ってから復元する
        QTimer.singleShot(0, lambda: self._restore_scroll(h_value, v_value))

    def _restore_scroll(self, h_value, v_value):
        for bar, value in ((self.view.horizontalScrollBar(), h_value),
                           (self.view.verticalScrollBar(), v_value)):
            bar.setValue(min(value, bar.maximum()))

    @Slot(int)
    def on_page_count_changed(self, count):
        # 再読み込み中は一時的に 0 ページになり値が切り詰められるので、現在ページで合わせ直す
        with QtCore.QSignalBlocker(self.page_spin):
            self.page_spin.setMaximum(max(count, 1))
            self.page_spin.setValue(self.view.pageNavigator().currentPage() + 1)
        self.page_count_label.setText(" / {} ".format(count))

    @Slot(int)
    def on_current_page_changed(self, page):
        with QtCore.QSignalBlocker(self.page_spin):
            self.page_spin.setValue(page + 1)

    @Slot(int)
    def jump_to_page(self, page_number):
        page = page_number - 1
        if 0 <= page < self.document.pageCount():
            self.view.pageNavigator().jump(page, QtCore.QPointF(), self.view.zoomFactor())

    @Slot()
    def prev_page(self):
        self.jump_to_page(self.view.pageNavigator().currentPage())

    @Slot()
    def next_page(self):
        self.jump_to_page(self.view.pageNavigator().currentPage() + 2)

    def _set_zoom(self, factor):
        self.view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.view.setZoomFactor(max(0.1, min(factor, 10.0)))

    @Slot()
    def zoom_in(self):
        self._set_zoom(self.view.zoomFactor() * ZOOM_STEP)

    @Slot()
    def zoom_out(self):
        self._set_zoom(self.view.zoomFactor() / ZOOM_STEP)

    @Slot()
    def fit_width(self):
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)

    @Slot()
    def fit_page(self):
        self.view.setZoomMode(QPdfView.ZoomMode.FitInView)
