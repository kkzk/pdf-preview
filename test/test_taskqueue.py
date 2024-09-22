import time
from PySide6.QtCore import Signal, SignalInstance

from pdf_preview.taskqueue import Worker, TaskQueue
from pytestqt.plugin import QtBot

def test_queue(qtbot: QtBot, qapp):
    queue = TaskQueue()

    with qtbot.waitSignal(queue.queue_empty, timeout=10000) as blocker:
        queue.add_task("123")
        queue.add_task("456")
        queue.add_task("789")
        queue.add_task("101")

