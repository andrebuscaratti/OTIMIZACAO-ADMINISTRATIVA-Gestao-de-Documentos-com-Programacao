import os
import io
import fitz
from tkinter import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from PyPDF2 import PdfWriter, PdfReader

class PDFPagePreview:
    def __init__(self, parent, pdf_path, page_index, move_up_callback, move_down_callback, remove_callback, fullscreen_callback, width=140):
        self.parent = parent
        self.pdf_path = pdf_path
        self.page_index = page_index
        self.width = width

        self.frame = Frame(parent, bd=2, relief="raised", padx=4, pady=4)
        self.frame.pack(pady=4, padx=4, fill='x')

        self.label_img = Label(self.frame)
        self.label_img.pack(side='left')

        nome_curto = os.path.basename(pdf_path)
        if len(nome_curto) > 30:
            nome_curto = nome_curto[:27] + "..."

        self.label_text = Label(
            self.frame,
            text=f"{nome_curto}\nPágina {page_index + 1}",
            justify='left',
            wraplength=self.width,
            anchor='w'
        )
        self.label_text.pack(side='left', padx=5, fill='x', expand=True)

        self.btn_frame = Frame(self.frame)
        self.btn_frame.pack(side='right', padx=5, anchor='ne')

        btn_width = 4
        btn_height = 2
        btn_font = ("Arial", 12, "bold")

        self.btn_remove = Button(
            self.btn_frame,
            text="✖",
            fg="red",
            font=btn_font,
            width=btn_width,
            height=btn_height,
            command=self.remove
        )
        self.btn_remove.pack(pady=2)
        self.add_hover_effect(self.btn_remove, hover_bg="#f99", hover_fg="darkred", original_fg='red')

        self.btn_up = Button(
            self.btn_frame,
            text="▲",
            width=btn_width,
            height=btn_height,
            font=btn_font,
            command=self.move_up
        )
        self.btn_up.pack(pady=2)
        self.add_hover_effect(self.btn_up)

        self.btn_down = Button(
            self.btn_frame,
            text="▼",
            width=btn_width,
            height=btn_height,
            font=btn_font,
            command=self.move_down
        )
        self.btn_down.pack(pady=2)
        self.add_hover_effect(self.btn_down)

        self.btn_fullscreen = Button(
            self.btn_frame,
            text="🔍",
            width=btn_width,
            height=btn_height,
            font=btn_font,
            command=self.fullscreen
        )
        self.btn_fullscreen.pack(pady=2)
        self.add_hover_effect(self.btn_fullscreen)

        self.move_up_callback = move_up_callback
        self.move_down_callback = move_down_callback
        self.remove_callback = remove_callback
        self.fullscreen_callback = fullscreen_callback

        self.image = None
        self.photoimage = None
        self.render_preview()

    def add_hover_effect(self, button, hover_bg="#4A90E2", hover_fg="white", original_bg='SystemButtonFace', original_fg='black'):
        def on_enter(e):
            e.widget['background'] = hover_bg
            e.widget['foreground'] = hover_fg
        def on_leave(e):
            e.widget['background'] = original_bg
            e.widget['foreground'] = original_fg
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def render_preview(self):
        try:
            doc = fitz.open(self.pdf_path)
            page = doc.load_page(self.page_index)
            zoom = self.width / page.rect.width
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            self.image = img
            self.photoimage = ImageTk.PhotoImage(img)
            self.label_img.config(image=self.photoimage)
        except Exception as e:
            print(f"Erro ao renderizar preview: {e}")

    def move_up(self):
        self.move_up_callback(self)

    def move_down(self):
        self.move_down_callback(self)

    def remove(self):
        self.remove_callback(self)

    def fullscreen(self):
        self.fullscreen_callback(self)

    def destroy(self):
        self.frame.destroy()

class PDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OTIMIZAÇÃO ADMINISTRATIVA: Gestão de Documentos com Programação")

        self.root.update_idletasks()
        width = 800
        height = 600
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.resizable(False, False)

        self.preview_list = []

        self.container_vertical = Frame(self.root)
        self.container_vertical.pack(fill=BOTH, expand=True)

        self.content_frame = Frame(self.container_vertical)
        self.content_frame.pack(fill=BOTH, expand=True)

        self.frame_left = Frame(self.content_frame, width=700)
        self.frame_left.pack(side=LEFT, fill=BOTH, expand=True)

        self.canvas = Canvas(self.frame_left)
        self.scroll_y = Scrollbar(self.frame_left, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = Frame(self.canvas)

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor='nw')
        self.canvas.configure(yscrollcommand=self.scroll_y.set)

        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.scroll_y.pack(side=RIGHT, fill=Y)

        self.btn_frame = Frame(self.content_frame, width=300)
        self.btn_frame.pack(side=RIGHT, fill=Y)

        btn_config = {
            "width": 30,
            "height": 2,
            "font": ("Arial", 14),
            "padx": 0,
            "pady": 0,
        }

        for (txt, cmd) in [
            ("📂 Importar Arquivos", self.importar_arquivos),
            ("✂ Dividir PDF", self.dividir_pdfs),
            ("📤 Exportar PDF Final", self.exportar_pdf),
            ("🗑  Limpar Tudo", self.limpar_tudo),
        ]:
            btn = Button(self.btn_frame, text=txt, command=cmd, **btn_config, anchor='w', justify='left')
            btn.pack(fill='x', expand=True)
            self.add_hover_effect(btn)

        self.footer = Label(self.container_vertical, text="Desenvolvido por André Buscaratti", font=("Arial", 10), bd=1, relief="sunken", anchor='center')
        self.footer.pack(side=BOTTOM, fill=X)

    def add_hover_effect(self, button, hover_bg="#4A90E2", hover_fg="white", original_bg='SystemButtonFace', original_fg='black'):
        def on_enter(e):
            e.widget['background'] = hover_bg
            e.widget['foreground'] = hover_fg
        def on_leave(e):
            e.widget['background'] = original_bg
            e.widget['foreground'] = original_fg
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def importar_arquivos(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDFs e Imagens", "*.pdf;*.png;*.jpg;*.jpeg")])
        for path in paths:
            if path.lower().endswith(".pdf"):
                reader = PdfReader(path)
                for i in range(len(reader.pages)):
                    self.add_preview(path, i)
            elif path.lower().endswith((".jpg", ".jpeg", ".png")):
                pdf_temp = path + ".pdf"
                img = Image.open(path).convert("RGB")
                img.save(pdf_temp)
                self.add_preview(pdf_temp, 0)

    def dividir_pdfs(self):
        pasta_destino = filedialog.askdirectory(title="Escolha a pasta para salvar as páginas")
        if not pasta_destino:
            return

        for idx, preview in enumerate(self.preview_list):
            try:
                reader = PdfReader(preview.pdf_path)
                writer = PdfWriter()
                writer.add_page(reader.pages[preview.page_index])

                nome_arquivo = f"pagina_{idx+1}.pdf"
                caminho_completo = os.path.join(pasta_destino, nome_arquivo)

                with open(caminho_completo, "wb") as f:
                    writer.write(f)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao dividir PDF: {e}")
                return

        messagebox.showinfo("Sucesso", "Páginas salvas com sucesso!")

    def add_preview(self, pdf_path, page_index):
        preview = PDFPagePreview(
            self.scroll_frame,
            pdf_path,
            page_index,
            self.move_up,
            self.move_down,
            self.remove_preview,
            self.open_fullscreen_preview,
            width=140
        )
        self.preview_list.append(preview)

    def move_up(self, preview):
        index = self.preview_list.index(preview)
        if index > 0:
            self.preview_list[index], self.preview_list[index-1] = self.preview_list[index-1], self.preview_list[index]
            self.refresh_previews()

    def move_down(self, preview):
        index = self.preview_list.index(preview)
        if index < len(self.preview_list) - 1:
            self.preview_list[index], self.preview_list[index+1] = self.preview_list[index+1], self.preview_list[index]
            self.refresh_previews()

    def remove_preview(self, preview):
        if messagebox.askyesno("Confirmar remoção", "Deseja remover essa página?"):
            self.preview_list.remove(preview)
            preview.destroy()
            self.refresh_previews()

    def refresh_previews(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        for preview in self.preview_list:
            preview.destroy()
        for preview in self.preview_list:
            preview.__init__(self.scroll_frame, preview.pdf_path, preview.page_index, self.move_up, self.move_down, self.remove_preview, self.open_fullscreen_preview, width=140)

    def exportar_pdf(self):
        destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not destino:
            return

        writer = PdfWriter()
        try:
            for preview in self.preview_list:
                reader = PdfReader(preview.pdf_path)
                writer.add_page(reader.pages[preview.page_index])

            with open(destino, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Sucesso", "PDF exportado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar PDF: {e}")

    def limpar_tudo(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.preview_list.clear()

    def open_fullscreen_preview(self, preview):
        win = Toplevel(self.root)
        win.title(f"Preview Página {preview.page_index + 1} - {os.path.basename(preview.pdf_path)}")
        win.attributes("-fullscreen", True)

        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        try:
            doc = fitz.open(preview.pdf_path)
            page = doc.load_page(preview.page_index)
            rect = page.rect
            scale_x = screen_width / rect.width
            scale_y = screen_height / rect.height
            zoom = min(scale_x, scale_y)
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            photo = ImageTk.PhotoImage(img)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar preview fullscreen: {e}")
            win.destroy()
            return

        container = Frame(win, bg="black")
        container.pack(fill=BOTH, expand=True)

        label_img = Label(container, image=photo, bg="black")
        label_img.image = photo
        label_img.pack(expand=True)

        btn_close = Button(container, text="✖ Fechar (Esc)", font=("Arial", 14), command=win.destroy, bg="#e74c3c", fg="white")
        btn_close.place(relx=1.0, rely=0.0, anchor="ne", x=-20, y=20)

        self.add_hover_effect(btn_close, hover_bg="#c0392b", hover_fg="white", original_bg="#e74c3c", original_fg="white")

        win.bind("<Escape>", lambda e: win.destroy())

if __name__ == "__main__":
    root = Tk()
    app = PDFApp(root)
    root.mainloop()
