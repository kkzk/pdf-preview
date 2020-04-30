# -*- coding: utf-8 -*-
import hashlib
import logging
import os
import tempfile
import shutil

from pathlib import Path, PureWindowsPath
from contextlib import contextmanager

import win32com.client
import win32con
import win32ui

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
        oFile = self.office.Documents.Open(
            filename, 0, True, False, "something")
        yield oFile
        oFile.Saved = True
        oFile.Close()
        oFile = None
        logging.debug("Document.Close()")

    def saveAsPDF(self, filename, tmp_name, select_sheet):
        with self._open(filename) as word:
            logging.debug("ExportAsFixedFormat:{}".format(filename))
            word.ExportAsFixedFormat(str(PureWindowsPath(tmp_name)), self.wdExportFormatPDF)
            return None           # 版数なし


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
        excel_vba = self.office.Workbooks.Open(filename, 0, True)
        yield excel_vba
        excel_vba.Saved = True
        excel_vba.Close()
        excel_vba = None

    def saveAsPDF(self, filename, tmp_fileame, selected_sheet: dict = None):
        with self._open(filename) as excel_workbook:
            if selected_sheet is None:
                # シートを選択したことがない場合には全体を変換する
                excel_workbook.ExportAsFixedFormat(self.xlTypePDF, tmp_fileame, self.xlQuality)
            else:
                # 最初の選択シートをExcelで選択します。このとき他のExcelのシートは選択を解除します。
                # 他の選択シートもExcelで選択するときに、選択状態を置換せず追加で選択します。
                # 選択したかどうかわからないものは、印刷対象として選択していることにします。
                LOGGER.debug("sheet selection:{}".format(selected_sheet))
                do_replace = True
                for excel_sheet in excel_workbook.sheets:
                    if selected_sheet.get(excel_sheet.name, True):  # デフォルトは「選択あり」
                        excel_sheet.Select(do_replace)
                        do_replace = False
                excel_workbook.ActiveSheet.ExportAsFixedFormat(self.xlTypePDF, tmp_fileame, self.xlQuality)


class Converter(object):
    """Office ドキュメントを PDF に変換する"""

    @staticmethod
    def convert(src_filename: str, selected_sheets: dict = None, force=False, dest_dir=".") -> str:
        """Office ファイルを PDF に変換する

        :param src_filename: 処理対象の Office ドキュメントファイル名
        :param selected_sheets: 印刷対象のシート名（Excelの場合にのみ有効）
        :param force: 変換済みファイルとタイムスタンプが同じ場合でも処理する
        :param dest_dir: 変換後のファイルの配置場所
        :return: 変換後のファイル名をフルパス
        """
        LOGGER.debug("ENTER:convert({})".format(src_filename))
        src = Path(src_filename).absolute()
        ext = Path(src_filename).suffix

        # 変換後のファイル名を作成する
        dst_filename = (Path(dest_dir) / hashlib.md5(str(src).encode()).hexdigest()).with_suffix(".pdf")
        dst_filename.parent.mkdir(exist_ok=True, parents=True)

        # Excel/Word を開く前にタイムスタンプを保存する
        src_mtime = src.stat().st_mtime
        # 変換先の PDF のタイムスタンプが同じなら変換しない
        if dst_filename.exists():
            dst_mtime = Path(dst_filename).stat().st_mtime
            if not force:
                if src_mtime == dst_mtime:
                    LOGGER.info("ファイルが作成済みです:{}".format(dst_filename))
                    return dst_filename

        # 出力先のファイルが OPEN されていると出力できない(Excel)
        # 一時ファイルを作成して後から移動する
        (fd, tmp_filename) = tempfile.mkstemp(".PDF")
        os.close(fd)

        if ext in [".xlsx", ".xls", ".xlsm"]:
            office = Excel()
        elif ext in [".docx", ".doc"]:
            office = Word()
        else:
            return None

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


def main(args):
    LOGGER.debug("ENTER:main({})".format(args))
    for source in args.source:
        abspath = Path(source).absolute()
        Converter.convert(abspath)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=str, nargs="+")
    parser.add_argument("-d", "--debug", action="store_true", default=False)
    args = parser.parse_args()
    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    main(args)
