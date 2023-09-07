#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
import time


class NewprojectApp:
    def __init__(self, master=None):
        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        self.frame1 = tk.Frame(self.toplevel1)
        self.button1 = tk.Button(self.frame1)
        self.button1.configure(text="pop progress")
        self.button1.grid(column="0", padx="20", pady="20", row="0")
        self.button1.configure(command=self.pop_progress_bar)
        self.frame1.configure(height="200", width="200")
        self.frame1.grid(column="0", row="0")
        self.toplevel1.configure(height="200", width="200")

        # Main widget
        self.mainwindow = self.toplevel1

    def run(self):
        self.mainwindow.mainloop()

    def pop_progress_bar(self, status="start"):
        popup = tk.Toplevel()
        popup.title("Extraindo Arquivos")
        self.label3 = tk.Label(popup)
        self.label3.configure(font="{Sitka Text} 12 {}", text="Extraindo Arquivos...")
        self.label3.grid(column="0", padx="50", pady="20", row="0")
        self.progressbar3 = ttk.Progressbar(popup)
        self.progressbar3.configure(
            length="200", mode="indeterminate", orient="horizontal"
        )
        self.progressbar3.grid(column="0", padx="50", pady="20", row="1")

        def animate_label(text, n=0):
            if n < len(text) - 1:
                # not complete yet, schedule next run one second later
                self.label3.after(100, animate_label, text, n + 1)
            # update the text of the label
            self.label3["text"] = text[: n + 1]

            if len(self.label3["text"]) == 12:  # loop_sanimation
                time.sleep(0.5)
                animate_label(text="Extraindo Arquivos...")

        def animate_status():
            if status == "start":
                self.progressbar3.start()
            else:
                self.progressbar3.stop()

        animate_label(text="Extraindo Arquivos...")

        animate_status()


if __name__ == "__main__":
    app = NewprojectApp()
    app.run()
