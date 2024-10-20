# -*- coding: utf-8 -*-
import os
import json
import logging
from datetime import timedelta, datetime
from glob import glob
from pathlib import Path

import openpyxl
from pypdf import PdfMerger
from PySide6 import QtCore
from PySide6 import QtWidgets
from PySide6.QtCore import QUrl, Slot, Qt
from PySide6.QtGui import QGuiApplication, QDesktopServices, QKeySequence
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QFileSystemModel, QTreeView, QSplitter, \
    QListWidgetItem, QAbstractItemView
from PySide6.QtWidgets import QVBoxLayout

from . import saveAsPDF, util
import shutil

LOGGER = logging.getLogger(__name__)


def merge_pdfs(paths, output):
    """PDFファイルを結合する"""
    if LOGGER.isEnabledFor(logging.DEBUG):
        for path in paths:
            LOGGER.debug("merge from {}".format(path))
        LOGGER.debug("merge to {}".format(output))

    merger = PdfMerger()
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
        cache_dir = util.cache_dir()
        pdfs = []
        cached_file = glob(str(cache_dir) + r"\*")
        for f in cached_file:
            if datetime.fromtimestamp(os.stat(f).st_birthtime) < datetime.now() - timedelta(days=2):
                LOGGER.debug("purge cache:{} ({})".format(os.stat(f).st_birthtime, f))
                try:
                    os.unlink(f)
                except PermissionError:
                    pass
        # PDF 作成
        for book_filename in self.all_books:
            sheets = self.sheet_selection.get(book_filename, None)
            force = True if book_filename in self.force_files else False
            converter = saveAsPDF.Converter()
            r = converter.convert(str(Path(self.root) / book_filename), sheets, force, cache_dir)
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

    def __init__(self, parent=None):
        super(CheckableFileSystemModel, self).__init__(parent)
        self.file_order_widget: QtWidgets.QListWidget = None
        self.setNameFilters(["*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx"])

    def setBookListWidget(self, widget):
        """Book の一覧を保持する ListWidget を設定する"""
        self.file_order_widget = widget

    def checkState(self, index):
        if self.filePath(index) in self.check:
            return QtCore.Qt.CheckState.Checked
        else:
            return QtCore.Qt.CheckState.Unchecked

    def relativePath(self, index):
        """index位置の相対パスを取得"""
        return str(Path(self.filePath(index)).relative_to(self.rootPath()))

    #
    # override
    #
    def flags(self, index):
        """チェックボックス付きであるフラグを追加"""
        return QFileSystemModel.flags(self, index) | Qt.ItemFlag.ItemIsUserCheckable

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.CheckStateRole:
            return QFileSystemModel.data(self, index, role)
        else:
            if index.column() == 0:
                ## ファイル一覧に入っているかどうかでチェックの有無を返す
                items = self.file_order_widget.findItems(self.relativePath(index), Qt.MatchFlag.MatchExactly)
                if len(items) > 0:
                    return Qt.CheckState.Checked
                else:
                    return Qt.CheckState.Unchecked

    def setData(self, index, value, role=None):
        if role == QtCore.Qt.ItemDataRole.CheckStateRole and index.column() == 0:
            LOGGER.debug("SELECT:{}".format(self.relativePath(index)))
            self.dataChanged.emit(index, index)
            self.updateCheckState.emit(self.relativePath(index), value)
            return True
        return super().setData(self, index, value, role)


class FileOrderWidget(QtWidgets.QListWidget):
    """選択したファイルの順番を変更するリスト

    signal: fileOrderChanged(list) list: 変更結果のファイルの一覧
    """
    fileOrderChanged = QtCore.Signal()

    def __init__(self, parent, root):
        super(FileOrderWidget, self).__init__(parent)
        self.root = root
        # Drag & Drop での順番変更を有効にする
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDropIndicatorShown(True)
        self.setMovement(QtWidgets.QListView.Movement.Snap)

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
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsDragEnabled)
            self.addItem(item)
            super().addItem(item)
        else:
            super().addItem(filename)

    @Slot(str, int)
    def updateFileList(self, filename, value):
        """ツリービューでファイルのチェック状態が変更されたとき。

        ファイル一覧のファイルを追加または削除する。"""
        LOGGER.debug("ファイル一覧のファイルを更新する(value: %s)", value)
        LOGGER.debug("ファイル一覧のファイルを更新する(value: %s)", QtCore.Qt.CheckState.Checked)
        if Qt.CheckState(value) == QtCore.Qt.CheckState.Checked:
            self.addItem(filename)
            self.setCurrentRow(self.count() - 1)
            LOGGER.debug("%s を追加した", filename)
        elif Qt.CheckState(value) == QtCore.Qt.CheckState.Unchecked:
            items = self.findItems(filename, Qt.MatchFlag.MatchExactly)
            for item in items:
                self.takeItem(self.row(item))
                LOGGER.debug("%s を削除した", filename)

    @Slot()
    def on_rows_changed(self):
        """ファイル一覧への追加・削除・順番変更"""
        LOGGER.debug("ファイル一覧が変更されました。有効なファイルは次の通りです。")
        paths = [self.item(i).text() for i in range(self.count()) if not self.item(i).isHidden()]
        LOGGER.debug("{}".format(paths))
        self.fileOrderChanged.emit()


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

    def __init__(self, parent, root):
        super(LeftPane, self).__init__(parent)
        self.sheet_selected_dict = {}

        #
        # ファイルのツリービュー
        #
        self.model = CheckableFileSystemModel(self)
        self.tv = QTreeView(self)
        self.tv.setModel(self.model)
        self.tv.setRootIndex(self.model.setRootPath(root))
        self.tv.header().setStretchLastSection(False)  # 一番右のカラムをストレッチする→False
        self.tv.setColumnWidth(0, 200)
        self.tv.doubleClicked.connect(self.open_file)

        #
        # ファイル一覧のビュー
        #
        self.book_list = FileOrderWidget(self, root)

        #
        # Excel のシート一覧のビュー
        #
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

    @Slot()
    def on_fileOrderChanged(self):
        LOGGER.debug("ファイル一覧が変更されました。有効なファイルは次の通りです。")
        paths = [self.book_list.item(i).text() for i in range(self.book_list.count()) if
                 not self.book_list.item(i).isHidden()]
        LOGGER.debug("{}".format(paths))
        self.file_selection_changed.emit(paths)


class MainWindow(QMainWindow):
    def load_sheet_selection(self):
        """シート選択の状態を復元する"""
        try:
            try:
                json_data = json.load(open(self.sheet_selection_filename, "r"))
            except UnicodeDecodeError:
                json_data = json.load(open(self.sheet_selection_filename, "r", encoding="utf-8"))
        except IOError:
            return
        try:
            self.left_pane.sheet_list.sheet_selection = json_data["sheets"]
            blocker = QtCore.QSignalBlocker(self.left_pane.book_list)
            for book_name in json_data["files"]:
                self.left_pane.book_list.addItem(book_name)
                self.left_pane.book_list.watcher.addPath(str(Path(self.source_dir) / book_name))
            del blocker
        except KeyError:
            pass

        # book_list の先頭のアイテムを選択する
        if self.left_pane.book_list.count() > 0:
            self.left_pane.book_list.setCurrentRow(0)

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
        # LOGGER.debug("save setting: {}: {}".format(self.sheet_selection_filename, json_data))
        LOGGER.debug("save")
        json.dump(json_data, open(self.sheet_selection_filename, "w", encoding="utf-8"), indent=4, ensure_ascii=False)

    def save(self):
        """PDFを保存する"""
        LOGGER.debug("save to:{}".format(self.saveto_path))
        shutil.copy(self.output_path, self.saveto_path)

    def __init__(self, source_path: str):
        """

        :param source_path: 対象のファイルまたはディレクトリ
        """
        super(MainWindow, self).__init__()

        cache_dir = util.cache_dir()
        if Path(source_path).is_file():
            # 対象がファイルの場合はファイルのあるディレクトリをツリーに表示する
            # 出力先はファイルと同じ場所。ファイルと同名で拡張子を変えたもの。
            self.source_dir = str(Path(source_path).parent)
            self.output_path = cache_dir / Path(source_path).with_suffix(".PDF").name
            self.saveto_path = Path(self.source_dir) / Path(source_path).with_suffix(".PDF").name
        else:
            # 出力先は対象ディレクトリの中。ディレクトリと同名で拡張子を変えたもの。
            self.source_dir = source_path
            self.output_path = cache_dir / Path(self.source_dir).with_suffix(".PDF").name
            self.saveto_path = Path(self.source_dir) / Path(source_path).with_suffix(".PDF").name

        self.sheet_selection_filename = cache_dir / self.output_path.with_suffix(".PDF.json")

        self.setWindowTitle(str(self.output_path))

        # ファイルツリーのモデルを作成
        self.left_pane = LeftPane(self, self.source_dir)
        self.left_pane.model.updateCheckState.connect(self.save_sheet_selection)  # ツリーでチェックされたら保存
        self.left_pane.file_selection_changed.connect(self.convertToPdf)  # ファイル選択の変更
        self.left_pane.sheet_selection_changed.connect(self.on_sheet_selection_changed)  # シート選択の変更

        self.web = QWebEngineView()
        self.web.settings().setAttribute(self.web.settings().WebAttribute.PluginsEnabled, True)
        self.web.settings().setAttribute(self.web.settings().WebAttribute.PdfViewerEnabled, True)
        self.viewer_html = str(Path(__file__).parent / Path('pdfjs-dist/web/viewer.html'))
        LOGGER.debug(self.viewer_html)

        # ログ表示用のテキストエリア
        self.console = QtWidgets.QTextEdit()
        self.console.setReadOnly(True)
        self.console.setLineWrapMode(QtWidgets.QTextEdit.LineWrapMode.NoWrap)

        # 右側の上下分割用の QSplitter を作成
        self.right_pane = QSplitter(QtCore.Qt.Vertical)
        self.right_pane.addWidget(self.web)
        self.right_pane.addWidget(self.console)
        self.right_pane.setStretchFactor(0, 1)  # 上部のウィジェット（webビューア）を優先
        self.right_pane.setStretchFactor(1, 0)  # 下部のウィジェット（ログ表示）を固定

        # 左右分割用の QSplitter を作成
        base = QSplitter()
        base.addWidget(self.left_pane)
        base.addWidget(self.right_pane)
        base.setStretchFactor(0, 0)  # 左はウインドウサイズ変更に追随させない
        base.setStretchFactor(1, 1)

        # 標準出力と標準エラー出力を textedit にリダイレクト   
        class QTextEditLogger(logging.Handler):
            def __init__(self, text_edit):
                super().__init__()
                self.text_edit = text_edit

            def emit(self, record):
                msg = self.format(record)
                QtCore.QMetaObject.invokeMethod(
                    self.text_edit,
                    "append",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, msg)
                )
                QtCore.QMetaObject.invokeMethod(
                    self.text_edit.verticalScrollBar(),
                    "setValue",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(int, self.text_edit.verticalScrollBar().maximum())
                )

        # Create a QTextEditLogger and set it up
        text_edit_logger = QTextEditLogger(self.console)
        log_format = "%(asctime)s:%(levelname)-7s:%(threadName)s:%(filename)s:%(lineno)d:%(funcName)s:%(message)s"
        text_edit_logger.setFormatter(logging.Formatter(log_format))
        logging.getLogger().addHandler(text_edit_logger)
        
        # メニューの追加
        menu = self.menuBar().addMenu(self.tr("File"))
        save_action = menu.addAction(self.tr("Save"), self.save)
        save_action.setShortcut(QKeySequence("Ctrl+S"))
        menu.addAction(self.tr("Exit"), self.close)

        self.load_sheet_selection()
        self.left_pane.book_list.fileOrderChanged.emit()
        self.setCentralWidget(base)
        self.resize(QtWidgets.QApplication.screens()[0].size() * 0.7)

    @Slot()
    def reload(self):
        # url = QUrl.fromLocalFile(str(self.output_path.absolute()))
        # LOGGER.debug("PDF表示を更新します {}".format(url))
        # self.web.load(url)

        param = QUrl.fromLocalFile(str(self.output_path.absolute())).toString()
        url= QUrl.fromUserInput(QUrl.fromLocalFile(self.viewer_html).toString() + "?file={}".format(param))
        LOGGER.debug("PDF表示を更新します {}".format(self.output_path))
        # if self.web.url() != url:
        self.web.load(url)
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
        p = ConvertThread(self.source_dir, self.output_path, book_names, recreate_file,
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
