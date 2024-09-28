# -*- coding: utf-8 -*-
import hashlib
import logging
import os
import tempfile
import shutil
from typing import Optional

from pathlib import Path, PureWindowsPath
from contextlib import contextmanager

import win32com.client
import win32con
import win32ui
from win32com.universal import com_error

LOGGER = logging.getLogger(__name__)


class OfficeBase(object):
    def __init__(self, application):
        self.application = application
        self.office = win32com.client.DispatchEx(self.application)
        self.st_mtime = None


class Word(OfficeBase):
    """PDF変換用MS-Wordクラス"""
    wdExportFormatPDF = 17  # PDF

    def __init__(self):
        super(Word, self).__init__("Word.Application")

    @contextmanager
    def _open(self, filename):
        filename = str(PureWindowsPath(filename))
        logging.debug("Document.Open({})".format(filename))
        application = self.office.Documents.Open(
            filename, 0, True, False, "something")
        yield application
        application.Saved = True
        application.Close()
        del application
        logging.debug("Document.Close()")

    def saveAsPDF(self, filename, tmp_name, select_sheet):
        assert select_sheet is None
        with self._open(filename) as word:
            logging.debug("ExportAsFixedFormat:{}".format(filename))
            word.ExportAsFixedFormat(str(PureWindowsPath(tmp_name)), self.wdExportFormatPDF)
            return None


class Excel(OfficeBase):
    xlTypePDF = 0
    xlQualityStandard = 0
    xlQualityMinimum = 1
    xlQuality = xlQualityStandard

    def __init__(self):
        super().__init__("Excel.Application")
        self.office.DisplayAlerts = False

    @contextmanager
    def _open(self, filename):
        # Excel が同名のファイルを開けないので異なるファイル名にコピーして開く
        (fd, tmp_filename) = tempfile.mkstemp(Path(filename).suffix)
        os.close(fd)
        shutil.copy2(filename, tmp_filename)
        LOGGER.debug("copy to %s", tmp_filename)
        application = self.office.Workbooks.Open(tmp_filename, 0, True)
        yield application
        application.Saved = True
        application.Close()
        del application
        os.unlink(tmp_filename)
        LOGGER.debug("delete %s", tmp_filename)


    def saveAsPDF(self, filename, pdf_filename, selected_sheet: dict = None):
        with self._open(filename) as excel_workbook:
            if selected_sheet is None:
                # シートを選択したことがない場合には全体を変換する
                excel_workbook.ExportAsFixedFormat(self.xlTypePDF, pdf_filename, self.xlQuality)
            else:
                # 元のExcelのシートの選択状態とは無関係に、ツールが指定するシートを選択します。
                # 
                # 選択したかどうかわからないものは、印刷対象として選択していることにします。
                #
                # 非表示のシートは過去に選択した記録があっても、印刷用のSelectメソッドを実行しません。
                #
                LOGGER.debug("sheet selection:{}".format(selected_sheet))
                do_replace = True
                for excel_sheet in excel_workbook.sheets:
                    if selected_sheet.get(excel_sheet.name, True):  # 指定がない場合のデフォルトは「選択あり」
                        if excel_sheet.Visible:
                            try:
                                excel_sheet.Select(do_replace)  # １シート目は新規選択、２シート目以降を追加選択
                            except com_error:
                                LOGGER.warning("シート名「{}」に対する Select メソッドの実行時時にエラーが発生しました。".format(excel_sheet.name))
                            do_replace = False
                excel_workbook.ActiveSheet.ExportAsFixedFormat(self.xlTypePDF, pdf_filename, self.xlQuality)


class Converter(object):
    """Office ドキュメントを PDF に変換する"""

    @staticmethod
    def convert(src_filename: str, selected_sheets: dict = None, force=False, cache_dir=".") -> Optional[str]:
        """Office ファイルを PDF に変換する

        :param src_filename: 処理対象の Office ドキュメントファイル名
        :param selected_sheets: 印刷対象のシート名（Excelの場合にのみ有効）
        :param force: 変換済みファイルとタイムスタンプが同じ場合でも処理する
        :param dest_dir: 変換後のファイルの配置場所
        :return: 変換後のファイル名をフルパス
        """
        LOGGER.debug("ENTER:convert({})".format(src_filename))
        src = Path(src_filename).absolute()
        ext = Path(src_filename).suffix.lower()

        # 変換後のファイル名を作成する
        dst_filename = (Path(cache_dir) / hashlib.md5(str(src).encode()).hexdigest()).with_suffix(".pdf")
        dst_filename.parent.mkdir(exist_ok=True, parents=True)

        # 変換元のファイルの存在を確認する
        if not src.exists():
            LOGGER.info("ファイルがありません:{}".format(src_filename))
            return None

        # Excel/Word を開く前にタイムスタンプを保存する
        src_mtime = src.stat().st_mtime
        # 変換先の PDF のタイムスタンプが同じなら変換しない
        if dst_filename.exists():
            dst_mtime = Path(dst_filename).stat().st_mtime
            if not force:
                if src_mtime == dst_mtime:
                    LOGGER.info("ファイルが作成済みです:{}".format(dst_filename))
                    return dst_filename

        if ext in [".pdf"]:
            shutil.copy2(src, dst_filename)
            return dst_filename

        if ext in [".xlsx", ".xls", ".xlsm"]:
            office = Excel()
        elif ext in [".docx", ".doc"]:
            office = Word()
        else:
            return None

        # 出力先のファイルが OPEN されていると出力できない(Excel)
        # 一時ファイルを作成して後から移動する
        (fd, tmp_filename) = tempfile.mkstemp(".PDF")
        os.close(fd)

        # Officeの機能でPDFを作成する
        office.saveAsPDF(str(src), tmp_filename, selected_sheets)

        while True:
            try:
                shutil.move(tmp_filename, dst_filename)
                os.utime(dst_filename, (src_mtime, src_mtime))  # タイムスタンプをコピー
                return dst_filename
            except OSError:
                import traceback
                traceback.print_exc()
                r = win32ui.MessageBox("出力先 PDF ファイルが使用中です", None, win32con.MB_RETRYCANCEL)
                if r == win32con.IDCANCEL:
                    return None


def main(sources):
    LOGGER.debug("ENTER:main({})".format(sources))
    for source in sources:
        abspath = Path(source).absolute()
        Converter.convert(abspath)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("sources", type=str, nargs="+")
    parser.add_argument("-d", "--debug", action="store_true", default=False)
    args = parser.parse_args()
    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    main(args.source)
