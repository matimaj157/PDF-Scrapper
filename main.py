import os
import customtkinter
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from modules.pdf_reader import read_pdf
from modules.docx_writer import write_docx
from tkinterdnd2 import DND_FILES, TkinterDnD


class MainApp:
    def __init__(self):
        self.file_path = None  
        self.open_first_window()

    def open_first_window(self):
        self.root1 = TkinterDnD.Tk()
        self.root1.title("Drag and Drop PDF File")

        drop_area = tk.Frame(self.root1, width=250, height=200, bg="dark grey" )
        drop_area.pack(padx=10, pady=10)

        drop_area.drop_target_register(DND_FILES)
        drop_area.dnd_bind('<<Drop>>', self.file_dropped)
        self.root1.eval('tk::PlaceWindow . center')

        self.root1.mainloop()

    def file_dropped(self, event):
        self.first_file_path = event.data.strip("{}")
        self.root1.destroy()
        self.open_second_window()
        
    def open_second_window(self):
        self.root2 = tk.Tk()
        self.root2.title("Convert PDF to DOCX")
        self.root2.eval('tk::PlaceWindow . center')
        self.root2.columnconfigure([0, 1], minsize=100)

        self.lines = read_pdf(self.first_file_path)
        self.load_data()

        self.buttons = ttk.Frame(self.root2)
        self.buttons.grid(row=0, column=1, columnspan=1, sticky="n")

        title = tk.Button(self.buttons, text="Title", command=lambda: self.change_style("Title"))
        new_sub = tk.Button(self.buttons, text="Subtitle", command=lambda: self.change_style("New Subtitle"))
        heading1 = tk.Button(self.buttons, text="Heading 1", command=lambda: self.change_style("Heading 1"))
        heading2 = tk.Button(self.buttons, text="Heading 2", command=lambda: self.change_style("Heading 2"))
        normal = tk.Button(self.buttons, text="Normal", command=lambda: self.change_style("Normal"))
        strong = tk.Button(self.buttons, text="Strong", command=lambda: self.change_style("Strong"))
        no_spacing = tk.Button(self.buttons, text="No spacing", command=lambda: self.change_style("No Spacing"))

        title.grid(row=0)
        new_sub.grid(row=1)
        heading1.grid(row=2)
        heading2.grid(row=3)
        normal.grid(row=4)
        strong.grid(row=5)
        no_spacing.grid(row=6)

        self.buttons.grid_rowconfigure(7, minsize=20)

        merge = tk.Button(self.buttons, text="Merge", command=lambda: self.merge())
        merge.grid(row=8)

        self.buttons.grid_rowconfigure(9, minsize=20)

        del_first_word = tk.Button(self.buttons, text="Del word", command=lambda: self.del_first_word())
        remove = tk.Button(self.buttons, text="Remove", command=lambda: self.change_style("Remove"))

        del_first_word.grid(row=10)
        remove.grid(row=11)

        self.buttons.grid_rowconfigure(12, minsize=20)
        
        convert_bt = tk.Button(self.buttons, text="Choose folder", command=self.select_folder)
        convert_bt.grid(row=13)

        convert_bt = tk.Button(self.buttons, text="Convert", command=self.convert_to_docx)
        convert_bt.grid(row=14)
    
    def load_data(self):
        self.content = customtkinter.CTkScrollableFrame(self.root2, width=600, height=500, orientation="vertical")
        self.content.grid(row=0, column=0, columnspan=1)

        style_to_bg = {
            "Title": "purple",
            "New Subtitle": "light blue",
            "Heading 1": "green",
            "Heading 2": "orange",
            "Normal": "grey",
            "Strong": "pink",
            "No Spacing": "light green"
        }

        for index, item in enumerate(self.lines):
            item['var'] = tk.IntVar()
            item['bg'] = style_to_bg.get(item['style'])

            checkbox = tk.Checkbutton(self.content, text=item['text'], variable=item['var'], onvalue=1, offvalue=0, bg=item['bg'])
            checkbox.grid(row=index, column=0, columnspan=1, sticky="w")
        
        self.root2.eval('tk::PlaceWindow . center')
        
    def merge(self):
        temp = ""
        merge_point = None

        for index, item in enumerate(self.lines):
            merge_point = index if item['var'].get() == 1 and merge_point is None else merge_point
            if item['var'].get() == 1:
                temp += item['text']

        self.lines[merge_point]['text'] = temp
        self.lines = [item for item in self.lines if item['var'].get() != 1 or item == self.lines[merge_point]]

        self.clear_columns()
        self.load_data()

    def change_style(self, style):
        if style == "Remove":
            self.lines = [item for item in self.lines if item['var'].get() != 1]
        else:
            for item in self.lines:
                if item['var'].get() == 1:
                    item['style'] = style

        self.clear_columns()
        self.load_data()

    def del_first_word(self):
        for item in self.lines:
            if item['var'].get() == 1:
                item['text'] = " ".join(item['text'].split()[1:])

        self.clear_columns()
        self.load_data()

    def select_folder(self):
        self.folder_selected = filedialog.askdirectory()

    def clear_columns(self):
        for widget in self.root2.grid_slaves():
            if int(widget.grid_info()["column"]) in [0]:
                widget.grid_forget()

    def convert_to_docx(self):
        current_path = os.path.dirname(os.path.abspath(__file__))
        template = f"{current_path}/templates/template.docx"
        output_path = f"{self.folder_selected}/converted.docx"
        write_docx(self.lines, template, output_path)
        self.root2.destroy()


if __name__ == "__main__":
    app = MainApp()
