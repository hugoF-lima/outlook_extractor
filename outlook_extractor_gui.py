#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, END
from tkinter import TclError, messagebox
import os
import win32com.client
from pathlib import Path
from threading import Thread
import os

# import pywintypes

# import win32api


class OutlookExtractorUiApp:
    def __init__(self, master=None):
        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        self.frame1 = tk.Frame(self.toplevel1)

        self.str_pdf_var = tk.StringVar()
        self.str_xml_var = tk.StringVar()

        self.select_path = tk.Button(self.frame1)
        self.select_path.configure(text="Selecionar Pasta")
        self.select_path.grid(
            column="0", padx="55", pady="15", row="1", rowspan="1", sticky="e"
        )
        self.select_path.configure(command=self.find_path)
        self.path_of_folder = tk.Entry(self.frame1)
        self.path_of_folder.configure(width="90")
        self.path_of_folder.grid(column="0", columnspan="1", row="2", padx="20")
        self.path_lbl = tk.Label(self.frame1)
        self.path_lbl.configure(text="Caminho da Pasta:")
        self.path_lbl.grid(column="0", padx="40", row="1")
        self.extract_files = tk.Button(self.frame1)
        self.extract_files.configure(text="Extrair Anexos")
        self.extract_files.grid(column="0", padx="55", pady="15", row="3", sticky="e")
        self.extract_files.configure(command=self.call_outlook_extract)
        self.choose_type_lbl = tk.Label(self.frame1)
        self.choose_type_lbl.configure(text="Formato a Extrair:")
        self.choose_type_lbl.grid(column="0", padx="63", row="3", sticky="w")
        self.pdf_check = tk.Checkbutton(self.frame1)
        self.pdf_check.configure(text="PDF", onvalue=".pdf", variable=self.str_pdf_var)
        self.pdf_check.deselect()
        self.pdf_check.grid(padx="180", row="3", sticky="e")
        self.xml_check = tk.Checkbutton(self.frame1)
        self.xml_check.configure(text="XML", onvalue=".xml", variable=self.str_xml_var)
        self.xml_check.deselect()
        self.xml_check.grid(column="0", padx="195", row="3", sticky="w")
        self.frame1.configure(height="200", width="200")
        self.frame1.grid(column="0", row="0")
        self.toplevel1.configure(height="200", width="200")
        self.toplevel1.title("Extrator de Anexos - Outlook")

        # Vars

        self.pick_folder = ""
        self.state_of_check_btns = ""

        # Main widget
        self.mainwindow = self.toplevel1

    def run(self):
        self.mainwindow.mainloop()

    def call_outlook_extract(self):  # bool to kill here?
        done = False
        try:
            paralel_t = Thread(target=self.extract_and_clean, daemon=True)
            paralel_t.start()
            done = True
            if done == True:
                return
        except TclError as tc:
            messagebox.showerror(title="Error", message="Unable to launch")

    def return_folder_path(self, text_widget):
        folder_string = filedialog.askdirectory()

        if folder_string != "":
            text_widget.delete(0, END)
            text_widget.insert(END, folder_string)

            # self.display_path.insert(END, json_text)  #

            return folder_string

    def find_path(self):
        self.pick_folder = self.return_folder_path(self.path_of_folder)

    def extract_and_clean(self):
        extract_first = False
        question = messagebox.askquestion(
            title="Confirmação!",
            message="Após extrair, O processo EXCLUIRÁ os arquivos .msg \nDeseja Continuar?",
        )
        if question:
            if self.pick_folder != "":
                if self.str_pdf_var.get() == "" and self.str_xml_var.get == "":
                    messagebox.showerror(
                        title="Erro!",
                        message=f"Escolha ao menos Um formato para Extrair os arquivos",
                    )
                else:
                    self.pick_folder = Path(self.pick_folder)
                    outlook = win32com.client.Dispatch(
                        "Outlook.Application"
                    ).GetNamespace("MAPI")
                    for file in os.listdir(self.pick_folder):
                        if file.endswith(".msg"):
                            filePath = str(self.pick_folder) + "\\\\" + file
                            msg = outlook.OpenSharedItem(filePath)
                            att = msg.Attachments
                            for i in att:
                                i.SaveAsFile(
                                    os.path.join(str(self.pick_folder), i.FileName)
                                )
                        extract_first = True

                    outlook.Application.Quit()
                    outlook = None
                    del outlook  # None of these lines quit outlook
                    os.system("taskkill /im outlook.exe /f")
                    """ os.system(
                        "reg add HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Security\\ /v ObjectModelGuard /t REG_DWORD /d 2 /f"
                    ) """

                    # remove msg files after (as well as checked)
                    if extract_first == True:
                        remove_msg = os.listdir(self.pick_folder)
                        for item in remove_msg:
                            if not item.endswith(
                                (self.str_pdf_var.get(), self.str_xml_var.get())
                            ):
                                os.remove(os.path.join(str(self.pick_folder), item))

                    messagebox.showinfo(
                        title="Exito!",
                        message="Arquivos extraídos com sucesso\n Foram extraídas",
                    )
            else:
                messagebox.showerror(
                    title="Caminho Não Encontrado!",
                    message=f"Escolha um caminho válido para Extrair os Anexos!",
                )


if __name__ == "__main__":
    app = OutlookExtractorUiApp()
    app.run()
