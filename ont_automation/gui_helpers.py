import logging
import threading
from tkinter import simpledialog

GUI_ROOT = None


def gui_ask(prompt_text: str) -> str:
    """Safely prompts the user for input from a background thread."""
    if not GUI_ROOT:
        return "exit"
    res = {"val": None}
    ev = threading.Event()

    def _ask():
        val = simpledialog.askstring("Input Required", prompt_text, parent=GUI_ROOT)
        res["val"] = val if val is not None else "exit"
        ev.set()

    GUI_ROOT.after(0, _ask)
    ev.wait()
    return res["val"]


class SafeTextRedirector:
    """Redirects stdout and stderr to a tkinter Text widget safely across threads."""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.after(0, self._insert, string)

    def _insert(self, string):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", string)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass


class TextHandler(logging.Handler):
    """Routes python logging to the SafeTextRedirector."""

    def __init__(self, redirector):
        logging.Handler.__init__(self)
        self.redirector = redirector

    def emit(self, record):
        msg = self.format(record)
        self.redirector.write(msg + "\n")
