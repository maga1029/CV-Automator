from ttkbootstrap import *
from tkinter import messagebox, scrolledtext
from Writing_01 import fun_write_word, fun_read_excel
import os
import re
import tkinter.filedialog

lista_btn = []
lista_directions = ["", "", ""]  # Archivo excel, carpeta destino, nombre final word
font_ttk = ("Comic Sans", 12)

with open("Instructions.txt", "r") as file:
    instructions_text = file.read()

root = Window(themename="sandstone")
root.geometry("800x300")
root.title("CV Automator")


def fun_instructions():
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    win_instructions = Window(themename="sandstone")
    win_instructions.title("Instructions")
    win_instructions.geometry("915x695")  # wxh
    text_widget = scrolledtext.ScrolledText(win_instructions, wrap="word", height=30, font=("Comic Sans", 12))
    text_widget.insert(tkinter.END, instructions_text)
    text_widget.grid(row=0, column=0)

    def win_ins_del():
        root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())
        win_instructions.destroy()

    win_instructions.protocol("WM_DELETE_WINDOW", win_ins_del)


def fun_file():
    file_dialog_file = tkinter.filedialog.askopenfilename(title="Select file", filetypes=[("Excel Files", ".xlsx")])
    if file_dialog_file:
        lista_directions[0] = file_dialog_file
        lbl_selected_file.config(text="")
        lbl_selected_file.config(text=re.search(r'[^/]+\.xlsx$', file_dialog_file).group())
        print(file_dialog_file)


def fun_folder():
    file_dialog_folder = tkinter.filedialog.askdirectory(title="Select folder")
    if file_dialog_folder:
        lista_directions[1] = file_dialog_folder
        lbl_selected_folder.config(text="")
        lbl_selected_folder.config(text=re.search(r'[^/\\]+(?=[/\\]?$)', file_dialog_folder).group())


def fun_create():
    lista_directions[2] = ent_name.get()
    if "" in lista_directions:
        messagebox.showinfo(title="Incomplete fields", message="Please fill all fields")
        return
    for _ in lista_btn:
        _.config(state="disabled")
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    if "." in lista_directions[2]:
        lista_directions[2] = lista_directions[2].split(".")[0]
    file_directory_docx = lista_directions[1] + "/" + lista_directions[2] + ".docx"
    file_directory_pdf = lista_directions[1] + "/" + lista_directions[2] + ".pdf"
    counter_filename_docx = 0
    counter_filename_pdf = 0
    while os.path.exists(file_directory_docx):
        file_directory_docx = lista_directions[1] + "/" + lista_directions[2] + f"({counter_filename_docx+1})" + ".docx"
        counter_filename_docx += 1
    while os.path.exists(file_directory_pdf):
        file_directory_pdf = lista_directions[1] + "/" + lista_directions[2] + f"({counter_filename_pdf+1})" + ".pdf"
        counter_filename_pdf += 1
    var_no_spaces = 0
    var_to_pdf = 0
    if var_btn_space.get():
        var_no_spaces = 1
    if var_btn_pdf.get():
        var_to_pdf = 1
    fun_write_word(file_directory_docx, fun_read_excel(lista_directions[0]), file_directory_pdf, var_no_spaces,
                   var_to_pdf)
    for _ in lista_btn:
        _.config(state="normal")
    root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())
    messagebox.showinfo(title="Processed finished", message="The program has finished generating the files. "
                                                            "You may continue or exit")


style_ttk = Style()
style_ttk.configure('comic_sans.TButton', font=('Arial', 12))
style_ttk.configure('comic_sans.TCheckbutton', font=('Comic Sans', 12))
var_btn_space = BooleanVar()
var_btn_pdf = BooleanVar()

btn_file = Button(root, text="Select Excel file", command=fun_file, style="success, comic_sans.TButton")
btn_folder = Button(root, text="Select folder", command=fun_folder, style="warning, comic_sans.TButton")
lbl_selected_file = Label(root, font=font_ttk)
lbl_selected_folder = Label(root, font=font_ttk)
lbl_name = Label(root, text="File Name", font=font_ttk)
ent_name = Entry(root, font=font_ttk)
btn_create = Button(root, text="Create CV file", command=fun_create, style="comic_sans.TButton")
btn_space = Checkbutton(root, text="No spaces", onvalue=1, offvalue=0, style="round-toggle, comic_sans.TCheckbutton, "
                                                                             "dark", variable=var_btn_space)
btn_pdf = Checkbutton(root, text="To PDF", onvalue=1, offvalue=0, style="round-toggle, comic_sans.TCheckbutton, dark",
                      variable=var_btn_pdf)
btn_instructions = Button(root, text="Instructions", command=fun_instructions, style="info, comic_sans.TButton")
lista_btn.extend([btn_file, btn_folder, btn_create, btn_instructions, btn_space, btn_pdf])

btn_file.grid(row=0, column=0, pady=10, padx=10)
btn_folder.grid(row=1, column=0, pady=10, padx=10, sticky="ew")
lbl_selected_file.grid(row=0, column=1, pady=10, padx=10)
lbl_selected_folder.grid(row=1, column=1, pady=10, padx=10)
lbl_name.grid(row=2, column=0, pady=10, padx=10)
ent_name.grid(row=2, column=1, pady=10, padx=10)
btn_space.grid(row=3, column=0, pady=10, padx=10)
btn_pdf.grid(row=3, column=1, pady=10, padx=10)
btn_instructions.grid(row=3, column=2, pady=10, padx=10)
btn_create.grid(row=4, column=0, columnspan=3, pady=10, padx=10, sticky="ew")

root.mainloop()
