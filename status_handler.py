# status_handler.py
# Centralized status and message handler for Tkinter GUIs

class StatusHandler:
    def __init__(self, status_var, status_label, app=None):
        self.status_var = status_var
        self.status_label = status_label
        self.app = app

    def set_status(self, msg, color="blue"):
        self.status_var.set(msg)
        self.status_label.config(foreground=color)
        if self.app:
            self.app.update_idletasks()

    def clear_status(self):
        self.status_var.set("")
        self.status_label.config(foreground="blue")
        if self.app:
            self.app.update_idletasks()
