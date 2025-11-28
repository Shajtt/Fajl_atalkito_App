import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document as BÁ_Docx
from docx2pdf import convert as BÁ_DocxToPdf
import fitz

class BÁ_FileConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Fájl átalakító App")
        master.geometry("750x450")
        master.resizable(False, False)
        
        master.configure(bg='#f0f0f0')

        self.full_paths = []

        style = ttk.Style(master)
        style.theme_use('clam')
        
        bg_color = '#f0f0f0'
        button_blue = '#4a86e8'
        button_green = '#4CAF50'
        listbox_bg = '#ffffff'
        listbox_fg = '#333333'
        
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, font=('Segoe UI', 12), foreground='#333333')
        style.configure('TButton', font=('Segoe UI', 11, 'bold'), padding=5)
        
        style.configure('Blue.TButton', 
                       background=button_blue, 
                       foreground='white',
                       borderwidth=0)
        style.map('Blue.TButton', 
                 background=[('active', '#3a76d8'), ('pressed', '#2a66c8')])
        
        style.configure('Green.TButton', 
                       background=button_green, 
                       foreground='white',
                       borderwidth=0)
        style.map('Green.TButton', 
                 background=[('active', '#3c9f40'), ('pressed', '#2c8f30')])
        btn_frame = ttk.Frame(master)
        btn_frame.pack(pady=20)
        self.refresh_button = ttk.Button(
            btn_frame, 
            text="Mappa kiválasztása a számítógépről", 
            command=self.BÁ_refresh_file_list,
            style='Blue.TButton'
        )
        self.refresh_button.grid(row=0, column=0, padx=10, ipadx=10, ipady=8)
        self.label = ttk.Label(master, text="Válasszon ki egy fájlt az átalakításhoz:")
        self.label.pack(pady=20)

        frame = ttk.Frame(master)
        frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=10)

        self.file_listbox = tk.Listbox(
            frame, 
            selectmode=tk.SINGLE, 
            font=('Segoe UI', 11),
            bg=listbox_bg,
            fg=listbox_fg,
            selectbackground='#e0e0e0',
            selectforeground='#333333',
            relief='flat',
            bd=1,
            highlightthickness=1,
            highlightcolor=button_blue,
            highlightbackground='#cccccc'
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        


        
        
        self.convert_button = ttk.Button(
            btn_frame, 
            text="Kiválasztott fájl átalakítása", 
            command=self.BÁ_on_convert_click,
            style='Green.TButton'
        )
        self.convert_button.grid(row=0, column=1, padx=10, ipadx=10, ipady=8)

    def BÁ_refresh_file_list(self):
        drive_path = filedialog.askdirectory(title="Válasszon egy mappát a számítógépről.")
        if not drive_path:
            return
        self.file_listbox.delete(0, tk.END)
        self.full_paths.clear()

        for root, _, files in os.walk(drive_path):
            for f in files:
                if f.lower().endswith(('.pdf', '.docx')):
                    full_path = os.path.join(root, f)
                    self.full_paths.append(full_path)
                    self.file_listbox.insert(tk.END, f)

    def BÁ_on_convert_click(self):
        selection = self.file_listbox.curselection()
        if not selection:
            messagebox.showwarning("Figyelem", "Nincs kiválasztott fájl!")
            return

        file_path = self.full_paths[selection[0]]

        if file_path.lower().endswith('.pdf'):
            self.BÁ_pdf_to_word(file_path)
        elif file_path.lower().endswith('.docx'):
            self.BÁ_word_to_pdf(file_path)
        else:
            messagebox.showwarning("Hiba", "Nem támogatott fájltípus!")

    def BÁ_pdf_to_word(self, pdf_path):
        output_path = pdf_path.replace('.pdf', '_atalakitott.docx')
        try:
            doc = BÁ_Docx()
            with fitz.open(pdf_path) as pdf:
                for page in pdf:
                    text = page.get_text("text")
                    doc.add_paragraph(text)
            doc.save(output_path)
            messagebox.showinfo("Siker", f"A PDF átalakítva Word dokumentummá:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Hiba", f"Hiba történt a PDF átalakítás során:\n{e}")

    def BÁ_word_to_pdf(self, word_path):
        try:
            output_path = word_path.replace('.docx', '_atalakitott.pdf')
            BÁ_DocxToPdf(word_path, output_path)
            messagebox.showinfo("Siker", f"A Word átalakítva PDF-fé:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Hiba", f"Hiba történt a Word átalakítás során:\n{e}\n\n"
                                         "Tanács: Zárja be a megnyitott Word dokumentumot, "
                                         "vagy futtassa a programot rendszergazdaként!")


if __name__ == "__main__":
    root = tk.Tk()
    app = BÁ_FileConverterApp(root)
    root.mainloop()