import argparse
import ctypes
import logging.config
import sys
import winreg
import os

from . import util

import yaml

LOGGER = logging.getLogger(__name__)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except OSError:
        return False


def register():
    import sys
    import pathlib
    path = str(pathlib.Path(sys.argv[0]).resolve())
    HKCR = winreg.HKEY_CLASSES_ROOT
    KEY = "pdf_Preview"

    h = winreg.CreateKeyEx(HKCR, r"Directory\shell\{}.convert".format(KEY))
    winreg.SetValue(h, "", winreg.REG_SZ, "PDF Preview")
    i = winreg.CreateKeyEx(h, "command")
    winreg.SetValue(i, "", winreg.REG_SZ, '{} -m pdf_preview "%1"'.format(sys.executable, path))

    h = winreg.CreateKeyEx(HKCR, r"*\shell\{}.convert".format(KEY))
    winreg.SetValue(h, "", winreg.REG_SZ, "PDF Preview")
    i = winreg.CreateKeyEx(h, "command")
    winreg.SetValue(i, "", winreg.REG_SZ, '{} -m pdf_preview "%1"'.format(sys.executable, path))

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source", nargs="?")
    parser.add_argument("-d", "--debug", action="store_true", default=False)
    parser.add_argument("-i", "--install", action="store_true", default=False)
    args = parser.parse_args()

    # モジュールと同一ディレクトリにある logging.ini をよみこみ、loggerを設定する
    # log_dir をログファイルの出力先として設定する
    log_dir = util.log_dir()
    log_path = log_dir / "pdf_preview.log"
    os.makedirs(log_dir, exist_ok=True)

    config_path = os.path.join(os.path.dirname(__file__), 'logging.yaml')
    config = yaml.safe_load(open(config_path).read())
    config["handlers"]["file"]["filename"] = log_path
    logging.config.dictConfig(config)
    LOGGER.debug("start")

    if args.install:
        if is_admin():
            LOGGER.debug("install")
            register()
        else:
            LOGGER.debug("elevate to install")
            LOGGER.debug("command: {} : {}".format(sys.executable, "{} -i".format(__file__)))
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, "{} -i -d".format(__file__), None, 1)

    if args.source:
        from . import main_window
        main_window.main(args.source)


if __name__ == '__main__':
                                           
    main()
