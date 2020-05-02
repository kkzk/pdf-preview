# -*- coding: utf-8 -*-
import os
import json
import logging
from datetime import timedelta, datetime
from glob import glob
from pathlib import Path

import openpyxl
from PyPDF2 import PdfFileMerger
from PySide2 import QtCore
from PySide2 import QtWidgets
from PySide2.QtCore import QUrl, Slot, Qt
from PySide2.QtGui import QGuiApplication, QDesktopServices
from PySide2.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PySide2.QtWidgets import QApplication, QMainWindow, QWidget, QFileSystemModel, QTreeView, QSplitter, \
    QListWidgetItem
from PySide2.QtWidgets import QVBoxLayout

from . import saveAsPDF

LOGGER = logging.getLogger(__name__)


def merge_pdfs(paths, output):
    LOGGER.debug("merge from {}".format(paths))
    LOGGER.debug("merge to {}".format(output))
    merger = PdfFileMerger()
    for path in paths:
        merger.append(open(path, "rb"))
    merger.write(str(output))
    merger.close()
    LOGGER.debug("merged")


class SignalHolder(QtCore.QObject):
    threadFinished = QtCore.Signal()


class ConvertThread(QtCore.QRunnable):
    def __init__(self, root: str, output_path: Path, all_books: list, force_files: list, sheet_selection: dict):
        super(ConvertThread, self).__init__()
        self.root = root
        self.obj_connection = SignalHolder()
        self.all_books = all_books
        self.output_path = output_path
        self.force_files = force_files
        self.sheet_selection = sheet_selection.copy()

    def run(self):
        LOGGER.debug("PDF変換開始")
        dest_dir = os.path.expandvars(r'$LOCALAPPDATA\pdf-preview\cache')
        pdfs = []
        cached_file = glob(dest_dir + r"\*")
        for f in cached_file:
            if datetime.fromtimestamp(os.stat(f).st_ctime) < datetime.now() - timedelta(days=2):
                LOGGER.debug("purge cache:{}".format(f))
                os.unlink(f)
        # PDF 作成
        for book_filename in self.all_books:
            sheets = self.sheet_selection.get(book_filename, None)
            force = True if book_filename in self.force_files else False
            converter = saveAsPDF.Converter()
            r = converter.convert(str(Path(self.root) / book_filename), sheets, force, dest_dir)
            if r is not None:
                pdfs.append(r)

        # PDF 結合
        merge_pdfs(pdfs, self.output_path)

        # 結合が終わったことを通知。 WebEngineView での再描画を期待する。
        self.obj_connection.threadFinished.emit()
        return


class CheckableFileSystemModel(QFileSystemModel):
    """チェックボックス付きのファイルツリー用モデル

    self.filepath(index) -> fullpath
    self.data(index, Qt.CheckStateRole) -> check state
    """

    updateCheckState = QtCore.Signal(str, int)

    def __init__(self, parent=None, single_file_mode=None):
        super(CheckableFileSystemModel, self).__init__(parent)
        self.file_order_widget: QtWidgets.QListWidget = None
        self.single_file = single_file_mode
        if self.single_file:
            self.setNameFilters([self.single_file])
        else:
            self.setNameFilters(["*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.pdf"])

    def setBookListWidget(self, widget):
        """Book の一覧を保持する ListWidget を設定する"""
        self.file_order_widget = widget

    def checkState(self, index):
        if self.filePath(index) in self.checks:
            return QtCore.Qt.Checked
        else:
            return QtCore.Qt.Unchecked

    def relativePath(self, index):
        """index位置の相対パスを取得"""
        return str(Path(self.filePath(index)).relative_to(self.rootPath()))

    #
    # override
    #
    def flags(self, index):
        """チェックボックス付きであるフラグを追加"""
        return QFileSystemModel.flags(self, index) | Qt.ItemIsUserCheckable

    def data(self, index, role=Qt.DisplayRole):
        if role != Qt.CheckStateRole:
            return QFileSystemModel.data(self, index, role)
        else:
            if index.column() == 0:
                items = self.file_order_widget.findItems(self.relativePath(index), Qt.MatchExactly)
                if len(items) > 0:
                    return Qt.Checked
                else:
                    return Qt.Unchecked

    def setData(self, index, value, role=None):
        if role == QtCore.Qt.CheckStateRole and index.column() == 0:
            LOGGER.debug("SELECT:{}".format(self.relativePath(index)))
            self.dataChanged.emit(index, index)
            self.updateCheckState.emit(self.relativePath(index), value)
            return True
        return QFileSystemModel.setData(self, index, value, role)


class FileOrderWidget(QtWidgets.QListWidget):
    """選択したファイルの順番を変更するリスト

    signal: fileOrderChanged(list) list: 変更結果のファイルの一覧
    """
    fileOrderChanged = QtCore.Signal()

    def __init__(self, parent, root, single_file):
        super(FileOrderWidget, self).__init__(parent)
        self.root = root
        self.single_file = single_file
        self.model().rowsInserted.connect(self.fileOrderChanged)
        self.model().rowsInserted.connect(self.addWatchPath)
        self.model().rowsRemoved.connect(self.fileOrderChanged)
        self.model().rowsMoved.connect(self.fileOrderChanged)
        # ファイル変更時にPDFを再作成する
        self.watcher = QtCore.QFileSystemWatcher()
        self.watcher.fileChanged.connect(self.on_file_changed)
        self.watcher.directoryChanged.connect(self.on_directory_changed)
        self.iconProvider = QtWidgets.QFileIconProvider()

        self.itemDoubleClicked.connect(self.open_file)

    def open_file(self, item):
        LOGGER.debug("ファイルを開きます:{}".format(item.text()))
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.abspath(item.text())))

    def abspath(self, path):
        return str(Path(self.root) / path)

    def abspath_parent(self, path):
        return str((Path(self.root) / path).parent)

    def addWatchPath(self, parent: QtCore.QModelIndex, first: int, last: int):
        LOGGER.debug("ファイルの変更を監視します:{}".format(self.abspath(self.item(first).text())))
        self.watcher.addPath(self.abspath(self.item(first).text()))
        self.watcher.addPath(self.abspath_parent(self.item(first).text()))

    def on_directory_changed(self, path: str):
        """ディレクトリの状態が変わったときに、選択一覧のファイルが有効かどうか更新する"""
        update = False
        for i in range(self.count()):
            current = not Path(self.abspath(self.item(i).text())).exists()
            if self.item(i).isHidden() != current:
                LOGGER.debug("ファイルの状態が {} に変わりました。:{})".format(current, self.item(i).text()))
                self.item(i).setHidden(current)
                update = True

        if update:
            # 本当に PDF を作り直さないといけないかどうかは、
            # 後続処理がファイルのタイムスタンプから判断します。
            self.on_rows_changed()

    def on_file_changed(self, path: str):
        LOGGER.debug("ファイルが変更されました:{}".format(path))
        r_path = str(Path(path).relative_to(self.root))
        items = self.findItems(r_path, Qt.MatchExactly)
        if len(items) > 0:
            r = self.watcher.addPath(path)
            if not r:
                LOGGER.debug("ファイルが削除されたようです。一覧から消します。:{}".format(path))
                items[0].setHidden(True)
            else:
                LOGGER.debug("一覧に変更はないのですがファイルが更新されたようなのでPDFを作り直してもらいます。")
                self.on_rows_changed()

    def addItem(self, filename):
        if isinstance(filename, str):
            item = QtWidgets.QListWidgetItem(filename)
            item.setIcon(self.iconProvider.icon(QtCore.QFileInfo(self.abspath(filename))))
            if not self.single_file or self.single_file == filename:
                item.setFlags(item.flags() | Qt.ItemIsEnabled)
            else:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
            self.addItem(item)
            super().addItem(item)
        else:
            super().addItem(filename)

    @Slot(str, int)
    def updateFileList(self, filename, value):
        """ツリービューでファイルのチェック状態が変更されたとき。

        ファイル一覧のファイルを追加または削除する。"""
        if value == QtCore.Qt.Checked:
            self.addItem(filename)
            self.setCurrentRow(self.count() - 1)
        elif value == QtCore.Qt.Unchecked:
            items = self.findItems(filename, Qt.MatchExactly)
            for item in items:
                self.takeItem(self.row(item))

    @Slot()
    def on_rows_changed(self):
        """ファイル一覧への追加・削除・順番変更"""
        LOGGER.debug("ファイル一覧が変更されました。有効なファイルは次の通りです。")
        paths = [self.item(i).text() for i in range(self.count()) if not self.item(i).isHidden()]
        LOGGER.debug("{}".format(paths))
        self.fileOrderChanged.emit(paths)


class ExcelSheetsView(QtWidgets.QListWidget):
    """選択したファイルがExcelだった時にシートの一覧を表示するView

    signal:
        sheetSelectionChanged(filename: str, sheetname: str, state: Qt.CheckState)
    """
    sheetSelectionChanged = QtCore.Signal(str, str, Qt.CheckState)

    def __init__(self, parent=None):
        super(ExcelSheetsView, self).__init__(parent)
        self.sheet_selection = {}
        self.currentBookName = None
        self.itemChanged.connect(self.on_itemChanged)

    def setSheetList(self, root, book_name):
        self.clear()
        self.currentBookName = book_name
        if not book_name.endswith((".xlsx", "xlsm")):
            return

        f_path = str((Path(root) / book_name).absolute())
        book = openpyxl.open(f_path, read_only=False)
        for sheet_name in book.sheetnames:
            item = QtWidgets.QListWidgetItem()
            item.setText(sheet_name)
            if book_name not in self.sheet_selection:
                item.setCheckState(Qt.Checked)
            elif self.sheet_selection[book_name].get(sheet_name, True):
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)

            # 非表示のシートを無効にする
            if book[sheet_name].sheet_state == 'hidden':
                item.setCheckState(Qt.Unchecked)
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)

            self.addItem(item)
        book.close()

    @Slot(QListWidgetItem)
    def on_itemChanged(self, item: QListWidgetItem):
        sheet_name = item.text()
        check_state = item.checkState()
        if self.currentBookName not in self.sheet_selection:
            self.sheet_selection[self.currentBookName] = {}
        if check_state == Qt.Checked:
            self.sheet_selection[self.currentBookName][sheet_name] = True
        else:
            self.sheet_selection[self.currentBookName][sheet_name] = False
        self.sheetSelectionChanged.emit(self.currentBookName, sheet_name, check_state)


class LeftPane(QWidget):
    """Windowの左半分

    PDF化対象のファイル選択、シート選択の状態は、中段のブック一覧の QListWidget に追加する
    QListWidgetItem とその data として保持します。ツリービューでは順番を変えられないので。
    """
    file_selection_changed = QtCore.Signal(list)
    sheet_selection_changed = QtCore.Signal(list, str, str, Qt.CheckState)

    @Slot(str, str, Qt.CheckState)
    def on_sheetSelectionUpdated(self, filename, sheet_name, state):
        paths = [self.book_list.item(i).text() for i in range(self.book_list.count())]
        self.sheet_selection_changed.emit(paths, filename, sheet_name, state)

    def on_update_check_state(self, filename, value):
        """ブック一覧を更新する"""
        self.book_list.updateFileList(filename, value)

    def on_currentItemChanged(self, newitem: QtWidgets.QListWidgetItem, olditem: QtWidgets.QListWidgetItem):
        """シート一覧を初期化し、新しいブックのシート一覧を表示する"""
        if newitem:
            r_path = newitem.text()
            root = self.model.rootPath()
            self.sheet_list.setSheetList(root, r_path)

    def __init__(self, parent, root, single_file):
        super(LeftPane, self).__init__(parent)
        self.sheet_selected_dict = {}

        # ファイルのツリービュー
        self.model = CheckableFileSystemModel(self, single_file)
        self.tv = QTreeView(self)
        self.tv.setModel(self.model)
        self.tv.setRootIndex(self.model.setRootPath(root))
        self.tv.header().setStretchLastSection(False)  # 一番右のカラムをストレッチする→False
        self.tv.setColumnWidth(0, 200)
        self.tv.resizeColumnToContents(1)
        self.tv.resizeColumnToContents(2)
        self.tv.resizeColumnToContents(3)

        # ファイル一覧のビュー
        self.book_list = FileOrderWidget(self, root, single_file)
        self.book_list.setAcceptDrops(True)
        self.book_list.setDragEnabled(True)
        self.book_list.setDragDropMode(self.book_list.InternalMove)
        # Excel のシート一覧のビュー
        self.sheet_list = ExcelSheetsView()

        # ツリービューのチェックボックスはブック一覧のリスト内容と連動
        self.model.setBookListWidget(self.book_list)

        # 上段のチェックボックスの変更に合わせて中段のファイルの一覧を更新する
        self.model.updateCheckState.connect(self.on_update_check_state)
        # 中段のブック一覧の順番が変更・チェックが変更になった時
        self.book_list.fileOrderChanged.connect(self.on_fileOrderChanged)
        self.book_list.currentItemChanged.connect(self.on_currentItemChanged)
        # 下段のシート一覧のチェック状態が変更になった時
        self.sheet_list.sheetSelectionChanged.connect(self.on_sheetSelectionUpdated)

        self.tv.doubleClicked.connect(self.open_file)

        s = QSplitter(Qt.Vertical)
        lo = QVBoxLayout()
        s.addWidget(self.tv)
        s.addWidget(self.book_list)
        s.addWidget(self.sheet_list)
        s.setStretchFactor(0, 2)  # 上から順に 2:1:1
        s.setStretchFactor(1, 1)
        s.setStretchFactor(2, 1)
        lo.addWidget(s)
        self.setLayout(lo)

    def open_file(self, index):
        LOGGER.debug("ファイルを開きます:{}".format(index))
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.model.filePath(index)))

    @Slot()
    def on_fileOrderChanged(self):
        LOGGER.debug("ファイル一覧が変更されました。有効なファイルは次の通りです。")
        paths = [self.book_list.item(i).text() for i in range(self.book_list.count()) if
                 not self.book_list.item(i).isHidden()]
        LOGGER.debug("{}".format(paths))
        self.tv.resizeColumnToContents(0)
        self.file_selection_changed.emit(paths)


class MainWindow(QMainWindow):
    def load_sheet_selection(self):
        """シート選択の状態を復元する"""
        try:
            json_data = json.load(open(self.sheet_selection_filename, "r"))
        except IOError:
            return
        try:
            self.left_pane.sheet_list.sheet_selection = json_data["sheets"]
            blocker = QtCore.QSignalBlocker(self.left_pane.book_list)
            for book_name in json_data["files"]:
                self.left_pane.book_list.addItem(book_name)
                self.left_pane.book_list.watcher.addPath(str(Path(self.source) / book_name))
            del blocker
        except KeyError:
            pass

    def save_sheet_selection(self):
        """シート選択の状態をファイルに保存する

        :return: None

        ファイルの選択は並び順を保持したいので book一覧の text を使用します。
        ファイルに対するシートの選択は、ファイルの選択有無を変更しても記憶して
        おきたいので MainWindow クラスが値を保持します。
        """
        json_data = {"files": [], "sheets": {}}

        for i in range(self.left_pane.book_list.count()):
            item = self.left_pane.book_list.item(i)
            path = item.text()  # path(relative)
            json_data["files"].append(path)
        json_data["sheets"] = self.left_pane.sheet_list.sheet_selection
        LOGGER.debug("save setting: {}: {}".format(self.sheet_selection_filename, json_data))
        json.dump(json_data, open(self.sheet_selection_filename, "w"), indent=4, ensure_ascii=False)

    def __init__(self, source):
        """

        :param source: 対象のファイルまたはディレクトリ
        """
        super(MainWindow, self).__init__()
        self.source = source
        # 対象がファイルの場合はファイルのあるディレクトリをツリーに表示する
        if Path(source).is_file():
            self.root = str(Path(source).parent)
            self.single_file = Path(source).name
        else:
            self.root = source
            self.single_file = None

        self.output_path = Path(self.source).absolute().with_suffix(".PDF")
        self.sheet_selection_filename = Path(self.root) / "PDF.json"
        self.setWindowTitle(str(self.output_path))

        self.left_pane = LeftPane(self, self.root, self.single_file)
        self.left_pane.model.updateCheckState.connect(self.save_sheet_selection)  # ツリーでチェックされたら保存
        self.left_pane.file_selection_changed.connect(self.convertToPdf)  # ファイル選択の変更
        self.left_pane.sheet_selection_changed.connect(self.on_sheet_selection_changed)  # シート選択の変更
        self.web = QWebEngineView()
        self.web.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
        self.viewer_html = str(Path(__file__).parent / Path('pdfjs-dist/web/viewer.html'))
        LOGGER.debug(self.viewer_html)

        layout = QVBoxLayout()
        layout.addWidget(self.web)
        self.right_pane = QWidget(self)
        self.right_pane.setLayout(layout)

        base = QSplitter()
        base.addWidget(self.left_pane)
        base.addWidget(self.right_pane)
        base.setStretchFactor(0, 0)  # 左はウインドウサイズ変更に追随させない
        base.setStretchFactor(1, 1)

        self.load_sheet_selection()
        self.left_pane.book_list.fileOrderChanged.emit()
        self.setCentralWidget(base)
        self.resize(QtWidgets.QApplication.screens()[0].size() * 0.7)

    @Slot()
    def reload(self):
        url = QUrl.fromLocalFile(str(self.output_path.absolute())).toString()
        LOGGER.debug("PDF表示を更新します {}".format(url))
        self.web.load(QUrl.fromUserInput(QUrl.fromLocalFile(self.viewer_html).toString() + "?file={}".format(url)))
        return

    # シートの選択を変えたら、変えたブックだけPDF変換してすべて結合
    @Slot(list, str, str, Qt.CheckState)
    def on_sheet_selection_changed(self, paths: list, filename: str, sheet_name: str, state: Qt.CheckState):
        LOGGER.debug("シートの選択変更：{} / {} / {}".format(filename, sheet_name, state))
        self.convertToPdf(paths, [filename])
        return

    # ブックを並び替えたら（最初のブックを選択したときも含む）PDF変換してすべて結合
    def convertToPdf(self, book_names, recreate_file: list = None):
        if recreate_file is None:
            recreate_file = []
        LOGGER.debug("PDF作成:{}".format(book_names))
        self.save_sheet_selection()
        if self.single_file:
            book_names = [self.single_file]
        p = ConvertThread(self.root, self.output_path, book_names, recreate_file,
                          self.left_pane.sheet_list.sheet_selection)
        p.obj_connection.threadFinished.connect(self.reload)
        QtCore.QThreadPool.globalInstance().start(p)


def main(source):
    LOGGER.debug("source:{}".format(source))
    QGuiApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication()
    main_window = MainWindow(source)
    main_window.show()
    app.exec_()
