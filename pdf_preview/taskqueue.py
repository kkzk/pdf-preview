from PySide6.QtCore import QThread, Signal, QObject, QMutex, QMutexLocker, QThreadPool

class Worker(QObject):
    task_finished = Signal()

    def __init__(self):
        super().__init__()
        self._mutex = QMutex()

    def run_task(self, task_name):
        with QMutexLocker(self._mutex):
            print(f"Task {task_name} started")
            self.task_finished.emit()

class TaskQueue(QObject):
    queue_empty = Signal()

    def __init__(self):
        super().__init__()
        self.worker = Worker()
        self.thread = QThread()
        self.worker.moveToThread(self.thread)
        self.thread.start()
        self.task_queue = []
        self.task_running = False

    def __del__(self) -> None:
        self.thread.quit()
        self.thread.wait()

    def add_task(self, task_name):
        self.task_queue.append(task_name)
        self.start_next_task()

    def start_next_task(self):
        if not self.task_queue:
            self.queue_empty.emit()
            return
        if not self.task_running and self.task_queue:
            self.task_running = True
            next_task = self.task_queue.pop(0)
            self.worker.task_finished.connect(self.on_task_finished)
            QThreadPool.globalInstance().start(lambda x = next_task: self.worker.run_task(x))

    def on_task_finished(self):
        self.worker.task_finished.disconnect(self.on_task_finished)
        self.task_running = False
        self.start_next_task()