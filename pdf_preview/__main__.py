import argparse
import ctypes
import logging
import sys
import winreg

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

    if args.debug:
        log_format = "%(asctime)s:%(levelname)-7s:%(threadName)s:%(filename)s:%(lineno)d:%(funcName)s:%(message)s"
        logging.basicConfig(level=logging.DEBUG, format=log_format)
    else:
        logging.basicConfig(level=logging.INFO)

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
