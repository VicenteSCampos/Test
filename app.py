import logging
import os
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox

import anthropic
import customtkinter as ctk
from dotenv import load_dotenv

_base = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_base, ".env"))

logging.basicConfig(
    filename=os.path.join(_base, "error.log"),
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s: %(message)s",
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


@dataclass
class QueueItem:
    audio_paths: list = field(default_factory=list)
    asignatura: str = ""
    fecha: str = ""
    pptx_path: str = ""
    nombre: str = ""
    estado: str = "pendiente"  # pendiente / procesando / listo / error


def extract_pptx(filepath: str) -> str:
    from pptx import Presentation
    prs = Presentation(filepath)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
        if texts:
            slides.append(f"[Slide {i}]: {' | '.join(texts)}")
    return "\n".join(slides)


def scan_vault_folders(vault: str) -> list:
    if not os.path.isdir(vault):
        return []
    return sorted([
        f for f in os.listdir(vault)
        if os.path.isdir(os.path.join(vault, f)) and not f.startswith(".")
    ])


def _strip_emoji(text: str) -> str:
    return re.sub(r"[^\x00-\x7F\u00C0-\u024F\u0400-\u04FF]", "", text).strip()


def _pdf_safe(text: str) -> str:
    return text.replace("\u00a0", " ")


def save_pdf(content: str, filepath: str):
    from fpdf import FPDF

    class PDF(FPDF):
        def footer(self):
            self.set_y(-12)
            self.set_font("Arial", "I", 8)
            self.set_text_color(150, 150, 150)
            self.cell(0, 10, f"Pagina {self.page_no()}", align="C")

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    pdf.add_font("Arial", fname="C:\\Windows\\Fonts\\arial.ttf")
    pdf.add_font("Arial", style="B", fname="C:\\Windows\\Fonts\\arialbd.ttf")
    pdf.add_font("Arial", style="I", fname="C:\\Windows\\Fonts\\ariali.ttf")

    lines = content.split("\n")
    start, in_front = 0, False
    for i, line in enumerate(lines):
        if i == 0 and line.strip() == "---":
            in_front = True
            continue
        if in_front and line.strip() == "---":
            start = i + 1
            break

    lm = pdf.l_margin
    pw = pdf.w - pdf.l_margin - pdf.r_margin

    in_mermaid = False
    for line in lines[start:]:
        s = _pdf_safe(_strip_emoji(line.rstrip())).strip()

        if s.startswith("```mermaid"):
            in_mermaid = True
            continue
        if in_mermaid:
            if s == "```":
                in_mermaid = False
            continue

        pdf.set_x(lm)

        if not s:
            pdf.ln(2)
        elif s.startswith("## "):
            pdf.set_font("Arial", "B", 14)
            pdf.set_text_color(30, 120, 200)
            pdf.ln(4)
            pdf.multi_cell(pw, 8, s[3:])
            pdf.set_text_color(0, 0, 0)
            pdf.ln(1)
        elif s.startswith("### "):
            pdf.set_font("Arial", "B", 11)
            pdf.ln(3)
            pdf.multi_cell(pw, 7, s[4:])
            pdf.ln(1)
        elif s.startswith("> "):
            pdf.set_font("Arial", "I", 10)
            pdf.set_fill_color(240, 240, 240)
            indent = 8
            pdf.set_x(lm + indent)
            pdf.multi_cell(pw - indent, 6, s[2:], fill=True)
        elif re.match(r"^[-*] ", s):
            pdf.set_font("Arial", "", 10)
            indent = 6
            pdf.set_x(lm + indent)
            pdf.multi_cell(pw - indent, 6, f"• {s[2:]}")
        elif re.match(r"^\d+\.", s):
            pdf.set_font("Arial", "", 10)
            indent = 6
            pdf.set_x(lm + indent)
            pdf.multi_cell(pw - indent, 6, s)
        elif s == "---":
            pdf.ln(2)
            pdf.set_draw_color(180, 180, 180)
            pdf.line(lm, pdf.get_y(), lm + pw, pdf.get_y())
            pdf.ln(2)
        else:
            clean = re.sub(r"\*\*(.*?)\*\*", r"\1", s)
            pdf.set_font("Arial", "", 10)
            pdf.multi_cell(pw, 6, clean)

    pdf.output(filepath)


class EditQueueItemDialog(ctk.CTkToplevel):
    def __init__(self, parent, item: QueueItem = None, vault_folders: list = None):
        super().__init__(parent)
        self.title("Configurar elemento")
        self.geometry("520x720")
        self.resizable(False, False)
        self.grab_set()
        self.result: QueueItem | None = None
        self._audio_paths: list[str] = list(item.audio_paths) if item else []
        self._pptx_path = item.pptx_path if item else ""
        self._vault_folders = vault_folders or []
        self._setup_ui(item)

    def _setup_ui(self, item: QueueItem = None):
        pad = {"padx": 20, "pady": (8, 2)}

        # Audio section
        audio_header = ctk.CTkFrame(self, fg_color="transparent")
        audio_header.pack(fill="x", **pad)
        ctk.CTkLabel(audio_header, text="Archivos de Audio",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left")
        ctk.CTkButton(audio_header, text="+ Agregar audio", width=120, height=28,
                      command=self._add_audio).pack(side="right")

        # Scrollable list of audio files (fixed height — scrolls internally)
        self._audio_list_frame = ctk.CTkScrollableFrame(self, height=120, fg_color="#1e1e1e")
        self._audio_list_frame.pack(fill="x", padx=20, pady=(4, 10))
        self._audio_list_frame.grid_columnconfigure(0, weight=1)
        self._render_audio_list()

        # PPT
        ctk.CTkLabel(self, text="PowerPoint (opcional)",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", **pad)
        ppt_row = ctk.CTkFrame(self, fg_color="transparent")
        ppt_row.pack(fill="x", padx=20, pady=(0, 10))
        ppt_row.grid_columnconfigure(0, weight=1)
        self._pptx_label = ctk.CTkLabel(
            ppt_row,
            text=os.path.basename(self._pptx_path) if self._pptx_path else "Ningún archivo seleccionado",
            text_color="white" if self._pptx_path else "gray",
            anchor="w",
        )
        self._pptx_label.grid(row=0, column=0, sticky="ew")
        btn_frame = ctk.CTkFrame(ppt_row, fg_color="transparent")
        btn_frame.grid(row=0, column=1, padx=(8, 0))
        ctk.CTkButton(btn_frame, text="Seleccionar", width=100,
                      command=self._select_pptx).grid(row=0, column=0, padx=(0, 4))
        ctk.CTkButton(btn_frame, text="✕", width=30, fg_color="gray30",
                      command=self._clear_pptx).grid(row=0, column=1)

        # Asignatura
        ctk.CTkLabel(self, text="Asignatura",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", **pad)
        self._asig_combo = ctk.CTkComboBox(self, values=self._vault_folders, height=36)
        self._asig_combo.set(item.asignatura if item else "")
        self._asig_combo.pack(fill="x", padx=20, pady=(0, 10))

        # Fecha
        ctk.CTkLabel(self, text="Fecha",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", **pad)
        self._fecha_entry = ctk.CTkEntry(self, height=36)
        self._fecha_entry.insert(0, item.fecha if item else datetime.now().strftime("%Y-%m-%d"))
        self._fecha_entry.pack(fill="x", padx=20, pady=(0, 10))

        # Nombre del archivo
        ctk.CTkLabel(self, text="Nombre del archivo (opcional)",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", **pad)
        ctk.CTkLabel(self, text="Si se deja vacío se usará: fecha - asignatura",
                     text_color="gray", anchor="w", font=ctk.CTkFont(size=11)).pack(fill="x", padx=20)
        self._nombre_entry = ctk.CTkEntry(self, height=36, placeholder_text="ej: Clase 1 - Introducción")
        self._nombre_entry.insert(0, item.nombre if item else "")
        self._nombre_entry.pack(fill="x", padx=20, pady=(4, 14))

        # Buttons
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 16))
        ctk.CTkButton(btn_row, text="Cancelar", fg_color="gray30", hover_color="gray40",
                      command=self.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_row, text="Aceptar", command=self._accept).pack(side="right")

    def _render_audio_list(self):
        for w in self._audio_list_frame.winfo_children():
            w.destroy()

        if not self._audio_paths:
            ctk.CTkLabel(self._audio_list_frame, text="Sin archivos — pulsa '+ Agregar audio'",
                         text_color="gray").grid(row=0, column=0, pady=10)
            return

        for i, path in enumerate(self._audio_paths):
            row_frame = ctk.CTkFrame(self._audio_list_frame, fg_color="transparent")
            row_frame.grid(row=i, column=0, sticky="ew", pady=2)
            row_frame.grid_columnconfigure(2, weight=1)

            ctk.CTkButton(row_frame, text="↑", width=26, height=26,
                          fg_color="gray30", hover_color="gray40",
                          command=lambda idx=i: self._move_audio(idx, -1)).grid(row=0, column=0, padx=(0, 2))
            ctk.CTkButton(row_frame, text="↓", width=26, height=26,
                          fg_color="gray30", hover_color="gray40",
                          command=lambda idx=i: self._move_audio(idx, 1)).grid(row=0, column=1, padx=(0, 6))

            name = os.path.basename(path)
            if len(name) > 36:
                name = name[:33] + "..."
            ctk.CTkLabel(row_frame, text=f"{i + 1}. {name}", anchor="w",
                         text_color="white").grid(row=0, column=2, sticky="ew", padx=(0, 6))

            ctk.CTkButton(row_frame, text="✕", width=26, height=26,
                          fg_color="#7a1a1a", hover_color="#a02020",
                          command=lambda idx=i: self._remove_audio(idx)).grid(row=0, column=3)

        self._audio_list_frame.grid_columnconfigure(0, weight=1)

    def _add_audio(self):
        paths = filedialog.askopenfilenames(
            title="Seleccionar archivos de audio",
            filetypes=[("Audio", "*.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mkv")],
        )
        for p in paths:
            if p not in self._audio_paths:
                self._audio_paths.append(p)
        self._render_audio_list()

    def _remove_audio(self, idx: int):
        self._audio_paths.pop(idx)
        self._render_audio_list()

    def _move_audio(self, idx: int, direction: int):
        new_idx = idx + direction
        if 0 <= new_idx < len(self._audio_paths):
            self._audio_paths[idx], self._audio_paths[new_idx] = \
                self._audio_paths[new_idx], self._audio_paths[idx]
            self._render_audio_list()

    def _select_pptx(self):
        path = filedialog.askopenfilename(
            title="Seleccionar PowerPoint",
            filetypes=[("PowerPoint", "*.pptx *.ppt")],
        )
        if path:
            self._pptx_path = path
            self._pptx_label.configure(text=os.path.basename(path), text_color="white")

    def _clear_pptx(self):
        self._pptx_path = ""
        self._pptx_label.configure(text="Ningún archivo seleccionado", text_color="gray")

    def _accept(self):
        if not self._audio_paths:
            messagebox.showerror("Error", "Agrega al menos un archivo de audio.", parent=self)
            return
        if not self._asig_combo.get().strip():
            messagebox.showerror("Error", "Ingresa o selecciona la asignatura.", parent=self)
            return
        self.result = QueueItem(
            audio_paths=list(self._audio_paths),
            asignatura=self._asig_combo.get().strip(),
            fecha=self._fecha_entry.get().strip(),
            pptx_path=self._pptx_path,
            nombre=self._nombre_entry.get().strip(),
        )
        self.destroy()


class ClaseObsidianApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Clase → Obsidian")
        self.geometry("660x760")
        self.resizable(False, False)
        self.queue: list[QueueItem] = []
        self._selected_row: int | None = None
        self._save_folder: str | None = None
        self._whisper_model = None
        self._setup_ui()

    def _setup_ui(self):
        ctk.CTkLabel(self, text="🎙️ Clase → Obsidian",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(20, 3))
        ctk.CTkLabel(self, text="Transcripción automática de clases universitarias",
                     text_color="gray").pack(pady=(0, 10))

        self.tabs = ctk.CTkTabview(self, height=300)
        self.tabs.pack(fill="x", padx=20)
        self.tabs.add("  Principal  ")
        self.tabs.add("  Opciones  ")

        self._setup_principal_tab()
        self._setup_opciones_tab()

        self.process_btn = ctk.CTkButton(
            self, text="🚀  Procesar Cola",
            command=self._start_processing,
            height=46, font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.process_btn.pack(pady=(12, 6), padx=20, fill="x")

        prog = ctk.CTkFrame(self, fg_color="transparent")
        prog.pack(fill="x", padx=20)
        self.progress_label = ctk.CTkLabel(prog, text="", text_color="gray")
        self.progress_label.pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(prog)
        self.progress_bar.pack(fill="x", pady=3)
        self.progress_bar.set(0)

        ctk.CTkLabel(self, text="📋  Estado:", anchor="w",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        self.log_box = ctk.CTkTextbox(self, height=145, font=ctk.CTkFont(size=11))
        self.log_box.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    def _setup_principal_tab(self):
        tab = self.tabs.tab("  Principal  ")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=5, pady=(10, 6))
        ctk.CTkButton(toolbar, text="+ Agregar", width=100, height=32,
                      command=self._add_item).pack(side="left", padx=(0, 4))
        ctk.CTkButton(toolbar, text="✏ Editar", width=90, height=32,
                      command=self._edit_item).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="✕ Quitar", width=90, height=32,
                      fg_color="gray30", hover_color="gray40",
                      command=self._remove_item).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="↑ Aplicar a todos", width=140, height=32,
                      fg_color="gray30", hover_color="gray40",
                      command=self._apply_to_all).pack(side="left", padx=4)

        tree_frame = ctk.CTkFrame(tab, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 5))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        self._build_treeview(tree_frame)

    def _setup_opciones_tab(self):
        tab = self.tabs.tab("  Opciones  ")
        tab.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(tab, text="🤖  Modelo Whisper",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=5, pady=(10, 4))
        model_row = ctk.CTkFrame(tab, fg_color="transparent")
        model_row.grid(row=1, column=0, sticky="w", padx=5, pady=(0, 12))
        self.model_var = ctk.StringVar(value="large-v3")
        for i, (lbl, val) in enumerate([("small  (~5 min)", "small"),
                                         ("medium  (~12 min)", "medium"),
                                         ("large-v3  (~20 min)", "large-v3")]):
            ctk.CTkRadioButton(model_row, text=lbl, variable=self.model_var,
                               value=val).grid(row=0, column=i, padx=10)

        ctk.CTkLabel(tab, text="📄  Formato del documento",
                     font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, sticky="w", padx=5, pady=(4, 4))
        fmt_row = ctk.CTkFrame(tab, fg_color="transparent")
        fmt_row.grid(row=3, column=0, sticky="w", padx=5, pady=(0, 12))
        self.format_var = ctk.StringVar(value="md")
        ctk.CTkRadioButton(fmt_row, text="Markdown (.md)", variable=self.format_var,
                           value="md").grid(row=0, column=0, padx=(0, 20))
        ctk.CTkRadioButton(fmt_row, text="PDF (.pdf)", variable=self.format_var,
                           value="pdf").grid(row=0, column=1)

        ctk.CTkLabel(tab, text="💾  Destino del documento",
                     font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, sticky="w", padx=5, pady=(4, 4))
        dest_row = ctk.CTkFrame(tab, fg_color="transparent")
        dest_row.grid(row=5, column=0, sticky="w", padx=5, pady=(0, 8))
        self.dest_var = ctk.StringVar(value="obsidian")
        ctk.CTkRadioButton(dest_row, text="Vault de Obsidian", variable=self.dest_var,
                           value="obsidian", command=self._toggle_dest).grid(row=0, column=0, padx=(0, 20))
        ctk.CTkRadioButton(dest_row, text="Guardar como archivo...", variable=self.dest_var,
                           value="file", command=self._toggle_dest).grid(row=0, column=1)

        self.vault_frame = ctk.CTkFrame(tab, fg_color="transparent")
        self.vault_frame.grid(row=6, column=0, sticky="ew", padx=5, pady=(0, 10))
        self.vault_frame.grid_columnconfigure(0, weight=1)
        self.vault_entry = ctk.CTkEntry(self.vault_frame, height=36)
        self.vault_entry.insert(0, os.getenv("OBSIDIAN_VAULT", r"C:\Users\vicen\Desktop\claude\Claude"))
        self.vault_entry.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(self.vault_frame, text="Cambiar", width=90,
                      command=self._select_vault).grid(row=0, column=1, padx=(8, 0))

    # ── helpers ──────────────────────────────────────────────────────────────

    def _toggle_dest(self):
        if self.dest_var.get() == "obsidian":
            self.vault_frame.grid()
        else:
            self.vault_frame.grid_remove()

    def _select_vault(self):
        path = filedialog.askdirectory(title="Seleccionar vault de Obsidian")
        if path:
            self.vault_entry.delete(0, "end")
            self.vault_entry.insert(0, path)

    def _log(self, msg):
        self.after(0, lambda: (self.log_box.insert("end", f"{msg}\n"), self.log_box.see("end")))

    def _set_progress(self, value, label=""):
        self.after(0, lambda: (self.progress_bar.set(value),
                               self.progress_label.configure(text=label)))

    # ── queue table ──────────────────────────────────────────────────────────

    def _audio_display(self, item: QueueItem) -> str:
        if not item.audio_paths:
            return "—"
        name = os.path.basename(item.audio_paths[0])
        if len(name) > 20:
            name = name[:17] + "..."
        if len(item.audio_paths) > 1:
            name += f" (+{len(item.audio_paths) - 1})"
        return name

    def _build_treeview(self, parent):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Queue.Treeview",
            background="#2b2b2b", foreground="white",
            fieldbackground="#2b2b2b", rowheight=28, borderwidth=0,
            font=("", 10),
        )
        style.configure("Queue.Treeview.Heading",
            background="#1a4a7a", foreground="white",
            font=("", 10, "bold"), relief="flat",
        )
        style.map("Queue.Treeview",
            background=[("selected", "#1a5276")],
            foreground=[("selected", "white")],
        )

        cols = ("#", "Archivos", "Asignatura", "Fecha", "PPT", "Estado")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings",
                                   style="Queue.Treeview", selectmode="browse")

        col_widths = {"#": 30, "Archivos": 155, "Asignatura": 135, "Fecha": 88, "PPT": 38, "Estado": 82}
        for col in cols:
            self._tree.heading(col, text=col)
            self._tree.column(col, width=col_widths[col], minwidth=col_widths[col], anchor="w")

        sb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self._tree.tag_configure("pendiente", foreground="white")
        self._tree.tag_configure("procesando", foreground="#f6e05e")
        self._tree.tag_configure("listo", foreground="#68d391")
        self._tree.tag_configure("error", foreground="#fc8181")

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)

    def _refresh_table(self):
        self._tree.delete(*self._tree.get_children())
        for i, item in enumerate(self.queue):
            asig = item.asignatura
            if len(asig) > 18:
                asig = asig[:15] + "..."
            self._tree.insert("", "end", iid=str(i), tags=(item.estado,), values=(
                str(i + 1),
                self._audio_display(item),
                asig or "—",
                item.fecha or "—",
                "✓" if item.pptx_path else "—",
                item.estado,
            ))
        if self._selected_row is not None and self._selected_row < len(self.queue):
            self._tree.selection_set(str(self._selected_row))
            self._tree.focus(str(self._selected_row))

    def _on_row_select(self, _event):
        sel = self._tree.selection()
        if sel:
            self._selected_row = int(sel[0])

    def _add_item(self):
        vault = self.vault_entry.get().strip() if hasattr(self, "vault_entry") else \
                os.getenv("OBSIDIAN_VAULT", "")
        dlg = EditQueueItemDialog(self, vault_folders=scan_vault_folders(vault))
        self.wait_window(dlg)
        if dlg.result:
            self.queue.append(dlg.result)
            self._selected_row = len(self.queue) - 1
            self._refresh_table()

    def _edit_item(self):
        if self._selected_row is None or self._selected_row >= len(self.queue):
            messagebox.showwarning("Advertencia", "Selecciona un elemento de la cola.")
            return
        item = self.queue[self._selected_row]
        if item.estado == "procesando":
            messagebox.showwarning("Advertencia", "No se puede editar un elemento en proceso.")
            return
        vault = self.vault_entry.get().strip() if hasattr(self, "vault_entry") else \
                os.getenv("OBSIDIAN_VAULT", "")
        dlg = EditQueueItemDialog(self, item=item, vault_folders=scan_vault_folders(vault))
        self.wait_window(dlg)
        if dlg.result:
            dlg.result.estado = item.estado
            self.queue[self._selected_row] = dlg.result
            self._refresh_table()

    def _remove_item(self):
        if self._selected_row is None or self._selected_row >= len(self.queue):
            messagebox.showwarning("Advertencia", "Selecciona un elemento de la cola.")
            return
        if self.queue[self._selected_row].estado == "procesando":
            messagebox.showwarning("Advertencia", "No se puede quitar un elemento en proceso.")
            return
        self.queue.pop(self._selected_row)
        if self.queue:
            self._selected_row = min(self._selected_row, len(self.queue) - 1)
        else:
            self._selected_row = None
        self._refresh_table()

    def _apply_to_all(self):
        if self._selected_row is None or self._selected_row >= len(self.queue):
            messagebox.showwarning("Advertencia", "Selecciona el elemento fuente.")
            return
        src = self.queue[self._selected_row]
        for i, item in enumerate(self.queue):
            if i != self._selected_row:
                item.asignatura = src.asignatura
                item.fecha = src.fecha
        self._refresh_table()

    # ── processing ───────────────────────────────────────────────────────────

    def _start_processing(self):
        pending = [i for i, item in enumerate(self.queue) if item.estado == "pendiente"]
        if not pending:
            messagebox.showerror("Error", "No hay elementos pendientes en la cola.")
            return

        if self.dest_var.get() == "file":
            folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
            if not folder:
                return
            self._save_folder = folder
        else:
            self._save_folder = None

        self.process_btn.configure(state="disabled", text="⏳  Procesando...")
        self.log_box.delete("1.0", "end")
        self.progress_bar.set(0)
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        pending = [i for i, item in enumerate(self.queue) if item.estado == "pendiente"]
        total = len(pending)
        success = 0
        try:
            for step, idx in enumerate(pending):
                item = self.queue[idx]
                item.estado = "procesando"
                self.after(0, self._refresh_table)

                try:
                    n = len(item.audio_paths)
                    label = os.path.basename(item.audio_paths[0]) if n == 1 \
                            else f"{n} archivos"
                    self._log(f"\n[{step + 1}/{total}] 🎵 {label}")

                    transcription = self._transcribe(item.audio_paths)

                    pptx_text = ""
                    if item.pptx_path:
                        self._log("📊 Extrayendo texto del PowerPoint...")
                        pptx_text = extract_pptx(item.pptx_path)
                        self._log(f"✅ PPT procesado — {len(pptx_text.split())} palabras")

                    is_obsidian_md = (self.dest_var.get() == "obsidian" and self.format_var.get() == "md")
                    content = self._call_claude(
                        transcription,
                        item.asignatura,
                        item.fecha,
                        pptx_text=pptx_text,
                        include_mermaid=is_obsidian_md,
                    )
                    full = content + f"\n\n---\n\n## 📄 Transcripción\n\n{transcription}"
                    self._save(full, item.asignatura, item.fecha, item.nombre)

                    item.estado = "listo"
                    success += 1
                    self._set_progress((step + 1) / total, f"[{step + 1}/{total}] Completado")

                except Exception as e:
                    logging.exception(f"Error procesando item {idx}")
                    self._log(f"❌ Error: {type(e).__name__}: {e}")
                    item.estado = "error"

                self.after(0, self._refresh_table)

            self._set_progress(1.0, f"¡Completado! {success}/{total} exitosos")
        finally:
            self.after(0, lambda: self.process_btn.configure(
                state="normal", text="🚀  Procesar Cola"))

    def _transcribe(self, audio_paths: list[str]) -> str:
        from faster_whisper import WhisperModel

        if self._whisper_model is None:
            self._log("⏳ Cargando modelo Whisper...")
            self._set_progress(0.05, "Cargando modelo...")
            self._whisper_model = WhisperModel(self.model_var.get(), device="cuda", compute_type="int8")

        model = self._whisper_model
        all_parts = []

        for file_idx, audio_path in enumerate(audio_paths):
            prefix = f"[{file_idx + 1}/{len(audio_paths)}] " if len(audio_paths) > 1 else ""
            self._log(f"🎙️ {prefix}Transcribiendo {os.path.basename(audio_path)}...")
            self._set_progress(0.1, f"{prefix}Transcribiendo audio...")

            segments, info = model.transcribe(audio_path, language="es", beam_size=5)
            parts = []
            for seg in segments:
                parts.append(seg.text.strip())
                if info.duration > 0:
                    file_progress = file_idx / len(audio_paths)
                    seg_progress = (seg.end / info.duration) / len(audio_paths)
                    pct = 0.1 + (file_progress + seg_progress) * 0.5
                    self._set_progress(
                        min(pct, 0.6),
                        f"{prefix}Transcribiendo... {int(seg.end/60)}:{int(seg.end%60):02d} / "
                        f"{int(info.duration/60)}:{int(info.duration%60):02d}"
                    )
            all_parts.extend(parts)
            self._log(f"✅ {prefix}Listo — {len(parts)} segmentos")

        transcription = " ".join(all_parts)
        self._log(f"✅ Transcripción total — {len(transcription.split())} palabras")
        self._set_progress(0.65, "Procesando con Claude AI...")
        return transcription

    def _call_claude(self, transcription: str, asignatura: str, fecha: str,
                     pptx_text: str = "", include_mermaid: bool = False) -> str:
        self._log("🤖 Enviando a Claude AI para estructurar el documento...")

        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("No se encontró ANTHROPIC_API_KEY en el archivo .env")

        client = anthropic.Anthropic(api_key=api_key)

        ppt_block = f"\n\nCONTENIDO DEL POWERPOINT:\n{pptx_text}" if pptx_text else ""

        mermaid_block = (
            "\n## 🗺️ Mapa Conceptual\n\n"
            "```mermaid\n"
            "graph TD\n"
            "    [Genera el diagrama aquí con máximo 10 nodos representando los conceptos principales y sus relaciones]\n"
            "```\n"
        ) if include_mermaid else ""

        tag = asignatura.lower().replace(" ", "-")

        prompt = (
            f"Eres un asistente académico experto. Analiza la siguiente información de una clase universitaria "
            f"y genera un documento en Markdown con EXACTAMENTE esta estructura:\n\n"
            f"---\n"
            f"asignatura: {asignatura}\n"
            f"fecha: {fecha}\n"
            f"tema: [extrae el tema principal en máximo 5 palabras]\n"
            f"tags: [{tag}]\n"
            f"---\n\n"
            f"## 📋 Ficha Técnica\n"
            f"- **Fecha:** {fecha}\n"
            f"- **Asignatura:** {asignatura}\n"
            f"- **Tema Principal:** [extrae el tema central]\n"
            f"- **Objetivo de la Clase:** [extrae o infiere el objetivo principal]\n\n"
            f"## 📝 Resumen Ejecutivo\n"
            f"[Máximo 5 párrafos. Explica el contexto y el núcleo de la clase de forma comprensible "
            f"para alguien que no asistió. Si hay PPT, intégralo con la transcripción.]\n\n"
            f"## 🔑 Conceptos Clave y Definiciones\n"
            f"[Para cada concepto importante usa este formato exacto:]\n\n"
            f"**Término:** Definición clara y concisa.\n"
            f"> Ejemplo: [si el profesor o el PPT dieron un ejemplo]\n\n"
            f"## 📊 Marco Lógico\n"
            f"[Si la clase presenta un modelo de negocio, estrategia empresarial o política pública:]\n"
            f"- **Objetivo:**\n"
            f"- **Medios:**\n"
            f"- **Resultados:**\n"
            f"- **Supuestos:**\n"
            f"[Si NO aplica escribe únicamente: \"No aplica para esta clase.\"]\n"
            f"{mermaid_block}\n"
            f"## ❓ Preguntas de Examen\n"
            f"[Entre 5 y 8 preguntas desafiantes. Usa EXACTAMENTE este formato para cada una:]\n\n"
            f"> [!question]- [escribe aquí la pregunta]\n"
            f"> [escribe aquí la respuesta detallada]\n\n"
            f"---\n"
            f"TRANSCRIPCIÓN:\n{transcription}"
            f"{ppt_block}"
        )

        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        self._set_progress(0.9, "Guardando documento...")
        return msg.content[0].text

    def _save(self, content: str, asignatura: str, fecha: str, nombre: str = ""):
        fmt = self.format_var.get()
        ext = ".pdf" if fmt == "pdf" else ".md"
        base = nombre if nombre else f"{fecha} - {asignatura}"
        filename = f"{base}{ext}"

        if self.dest_var.get() == "obsidian":
            vault = self.vault_entry.get().strip()
            subfolder = os.path.join(vault, asignatura)
            folder = subfolder if os.path.isdir(subfolder) else vault
            filepath = os.path.join(folder, filename)
        else:
            filepath = os.path.join(self._save_folder, filename)

        if fmt == "pdf":
            save_pdf(content, filepath)
        else:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(content)

        self._log(f"✅ Guardado: {filename}")


class SetupDialog(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.title("Configuración inicial")
        self.geometry("460x280")
        self.resizable(False, False)
        self.grab_set()
        self.api_key = None
        self._setup_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_ui(self):
        ctk.CTkLabel(self, text="🎙️ Clase → Obsidian",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(25, 4))
        ctk.CTkLabel(self, text="Configuración inicial — solo se hace una vez",
                     text_color="gray").pack(pady=(0, 20))

        ctk.CTkLabel(self, text="API Key de Anthropic:",
                     font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", padx=30)
        self.key_entry = ctk.CTkEntry(self, placeholder_text="sk-ant-api03-...",
                                      height=40, show="•")
        self.key_entry.pack(fill="x", padx=30, pady=(5, 6))

        ctk.CTkButton(self, text="👁 Mostrar / Ocultar", width=160, height=28,
                      fg_color="gray30", hover_color="gray40",
                      command=self._toggle_show).pack(pady=(0, 15))

        ctk.CTkButton(self, text="Guardar y continuar",
                      height=42, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._save).pack(fill="x", padx=30, pady=(0, 20))

    def _toggle_show(self):
        self.key_entry.configure(show="" if self.key_entry.cget("show") == "•" else "•")

    def _save(self):
        key = self.key_entry.get().strip()
        if not key.startswith("sk-ant-"):
            messagebox.showerror("Error", "La key debe comenzar con 'sk-ant-'", parent=self)
            return
        env_path = os.path.join(_base, ".env")
        vault = os.getenv("OBSIDIAN_VAULT", r"C:\Users\vicen\Desktop\claude\Claude")
        with open(env_path, "w", encoding="utf-8") as f:
            f.write(f"ANTHROPIC_API_KEY={key}\n")
            f.write(f"OBSIDIAN_VAULT={vault}\n")
        load_dotenv(dotenv_path=env_path, override=True)
        self.api_key = key
        self.destroy()

    def _on_close(self):
        if not self.api_key:
            if messagebox.askokcancel("Salir", "Sin API key la app no puede funcionar. ¿Salir?",
                                      parent=self):
                self.master.destroy()


if __name__ == "__main__":
    root = ctk.CTk()
    root.withdraw()

    if not os.getenv("ANTHROPIC_API_KEY"):
        dialog = SetupDialog()
        dialog.master = root
        root.wait_window(dialog)
        if not os.getenv("ANTHROPIC_API_KEY"):
            root.destroy()
            exit()

    root.destroy()
    app = ClaseObsidianApp()
    app.mainloop()
