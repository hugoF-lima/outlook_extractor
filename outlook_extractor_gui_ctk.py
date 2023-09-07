import customtkinter
from tkinter import filedialog, END
from tkinter import TclError, messagebox
import os
import win32com.client
from pathlib import Path
from threading import Thread
import os
import tkinter as tk
from tkinter import PhotoImage
from bg4_strings import icon_16_b64, icon_128_b64
import win_32_custom


class OutlookExtractorUiApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("587x163")
        self.title("Extrator de Anexos - Outlook")
        self.icon_16_img = PhotoImage(data=(icon_16_b64))
        self.iconphoto("-default", self.icon_16_img)
        self.minsize(587, 163)
        customtkinter.set_appearance_mode("dark")

        self.str_pdf_var = tk.StringVar()
        self.str_xml_var = tk.StringVar()

        self.path_lbl = customtkinter.CTkLabel(master=self)
        self.path_lbl.configure(text="Caminho da Pasta:")
        self.path_lbl.grid(row=0, column=0, columnspan=1, padx=55, pady=10, sticky="n")

        self.select_path = customtkinter.CTkButton(master=self)
        self.select_path.configure(text="Selecionar Pasta")
        self.select_path.configure(command=self.find_path)
        self.select_path.grid(row=0, column=1, padx=85, pady=10, sticky="n")

        self.path_of_folder = customtkinter.CTkEntry(master=self, width=490)
        self.path_of_folder.grid(column=0, columnspan=2, row=2, padx=20, pady=15)

        self.extract_files = customtkinter.CTkButton(master=self)
        self.extract_files.configure(text="Extrair Anexos")
        self.extract_files.configure(state="disabled")
        self.extract_files.grid(column=1, padx=55, pady=15, row=3, sticky="e")
        self.extract_files.configure(command=self.call_outlook_extract)

        self.choose_type_lbl = customtkinter.CTkLabel(master=self)
        self.choose_type_lbl.configure(text="Formato a Extrair:")
        self.choose_type_lbl.grid(column=0, padx=18, row=3, sticky="w")
        self.pdf_check = customtkinter.CTkCheckBox(master=self, onvalue=".pdf")
        self.pdf_check.configure(text="PDF", variable=self.str_pdf_var)
        self.pdf_check.configure(state="disabled")
        self.pdf_check.deselect()
        self.pdf_check.grid(padx=170, row=3, columnspan=2, sticky="w")
        self.xml_check = customtkinter.CTkCheckBox(master=self, onvalue=".xml")
        self.xml_check.configure(text="XML", variable=self.str_xml_var)
        self.xml_check.configure(state="disabled")
        self.xml_check.deselect()
        self.xml_check.grid(padx=250, columnspan=2, row=3, sticky="w")

        self.pick_folder = ""
        self.state_of_check_btns = ""

        self.icon_128_img = PhotoImage(data=(icon_128_b64))
        # setting icon through b64string
        self.tk.call("wm", "iconphoto", self._w, self.icon_128_img)

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
        if self.pick_folder != None:
            print(self.pick_folder)
            self.extract_files.configure(state="active")
            self.xml_check.configure(state="normal")
            self.pdf_check.configure(state="normal")

    def popup_progress_bar(self):
        popup = customtkinter.CTkToplevel()
        customtkinter.CTkLabel(popup, text="Files being downloaded").grid(
            row=0, column=0
        )
        progress_bar = customtkinter.CTkProgressBar(
            popup, mode="indeterminate", maximum=100
        )
        progress_bar.grid(row=1, column=0)

    def extract_and_clean(self):
        extract_first = False
        question = messagebox.askquestion(
            title="Confirmação!",
            message="Após extrair, O processo EXCLUIRÁ os arquivos .msg \nDeseja Continuar?",
        )
        if question == "yes":
            if self.pick_folder != "":
                if self.str_pdf_var.get() == "" and self.str_xml_var.get == "":
                    messagebox.showerror(
                        title="Erro!",
                        message=f"Escolha ao menos Um formato para Extrair os arquivos",
                    )
                else:
                    #self.popup_progress_bar()
                    self.pick_folder = Path(self.pick_folder)
                    # outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                    outlook = win_32_custom.custom_dispatch("Outlook.Application")
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
                    # outlook = None
                    # del outlook  # None of these lines quit outlook
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
                    self.progress_check.stop()
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
    app.mainloop()
