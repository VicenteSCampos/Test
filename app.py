import os
import threading
from datetime import datetime
from tkinter import filedialog, messagebox

import anthropic
import customtkinter as ctk
from dotenv import load_dotenv

load_dotenv()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ClaseObsidianApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Clase → Obsidian")
        self.geometry("640x780")
        self.resizable(False, False)
        self.audio_path = None
        self._setup_ui()

    def _setup_ui(self):
        ctk.CTkLabel(self, text="🎙️ Clase → Obsidian",
                     font=ctk.CTkFont(size=26, weight="bold")).pack(pady=(25, 4))
        ctk.CTkLabel(self, text="Transcripción automática de clases universitarias",
                     text_color="gray").pack(pady=(0, 18))

        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=20)
        frame.grid_columnconfigure(0, weight=1)

        # Audio
        ctk.CTkLabel(frame, text="📁  Archivo de Audio",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 5))

        audio_row = ctk.CTkFrame(frame, fg_color="transparent")
        audio_row.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 10))
        audio_row.grid_columnconfigure(0, weight=1)

        self.audio_label = ctk.CTkLabel(audio_row, text="Ningún archivo seleccionado",
                                        text_color="gray", anchor="w")
        self.audio_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(audio_row, text="Seleccionar", width=120,
                      command=self._select_audio).grid(row=0, column=1, padx=(10, 0))

        # Asignatura
        ctk.CTkLabel(frame, text="📋  Asignatura",
                     font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, sticky="w", padx=15, pady=(10, 5))
        self.asignatura_entry = ctk.CTkEntry(frame, placeholder_text="Ej: Marketing Estratégico", height=38)
        self.asignatura_entry.grid(row=3, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Fecha
        ctk.CTkLabel(frame, text="📅  Fecha",
                     font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, sticky="w", padx=15, pady=(10, 5))
        self.fecha_entry = ctk.CTkEntry(frame, height=38)
        self.fecha_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.fecha_entry.grid(row=5, column=0, sticky="ew", padx=15, pady=(0, 10))

        # Modelo
        ctk.CTkLabel(frame, text="🤖  Modelo Whisper",
                     font=ctk.CTkFont(weight="bold")).grid(row=6, column=0, sticky="w", padx=15, pady=(10, 5))

        model_row = ctk.CTkFrame(frame, fg_color="transparent")
        model_row.grid(row=7, column=0, sticky="w", padx=15, pady=(0, 15))

        self.model_var = ctk.StringVar(value="large-v3")
        models = [
            ("small  (~5 min)", "small"),
            ("medium  (~12 min)", "medium"),
            ("large-v3  (~20 min)", "large-v3"),
        ]
        for i, (label, value) in enumerate(models):
            ctk.CTkRadioButton(model_row, text=label, variable=self.model_var,
                               value=value).grid(row=0, column=i, padx=12)

        # Botón
        self.process_btn = ctk.CTkButton(
            self, text="🚀  Procesar Clase",
            command=self._start_processing,
            height=50, font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.process_btn.pack(pady=18, padx=20, fill="x")

        # Progreso
        prog_frame = ctk.CTkFrame(self, fg_color="transparent")
        prog_frame.pack(fill="x", padx=20)

        self.progress_label = ctk.CTkLabel(prog_frame, text="", text_color="gray")
        self.progress_label.pack(anchor="w")
        self.progress_bar = ctk.CTkProgressBar(prog_frame)
        self.progress_bar.pack(fill="x", pady=4)
        self.progress_bar.set(0)

        # Log
        ctk.CTkLabel(self, text="📋  Estado:", anchor="w",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(12, 3))
        self.log_box = ctk.CTkTextbox(self, height=165, font=ctk.CTkFont(size=12))
        self.log_box.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    # ── helpers ──────────────────────────────────────────────────────────────

    def _select_audio(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de audio",
            filetypes=[("Audio", "*.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mkv")]
        )
        if path:
            self.audio_path = path
            self.audio_label.configure(text=os.path.basename(path), text_color="white")

    def _log(self, msg):
        self.after(0, lambda: (self.log_box.insert("end", f"{msg}\n"), self.log_box.see("end")))

    def _set_progress(self, value, label=""):
        self.after(0, lambda: (self.progress_bar.set(value),
                               self.progress_label.configure(text=label)))

    def _start_processing(self):
        if not self.audio_path:
            messagebox.showerror("Error", "Selecciona un archivo de audio primero.")
            return
        if not self.asignatura_entry.get().strip():
            messagebox.showerror("Error", "Ingresa la asignatura.")
            return

        self.process_btn.configure(state="disabled", text="⏳  Procesando...")
        self.log_box.delete("1.0", "end")
        self.progress_bar.set(0)
        threading.Thread(target=self._process, daemon=True).start()

    # ── pipeline ─────────────────────────────────────────────────────────────

    def _process(self):
        try:
            transcription = self._transcribe()
            content = self._call_claude(
                transcription,
                self.asignatura_entry.get().strip(),
                self.fecha_entry.get().strip(),
            )
            self._save(content, self.asignatura_entry.get().strip(), self.fecha_entry.get().strip())
            self._set_progress(1.0, "¡Completado!")
            self._log("✅ ¡Documento creado exitosamente en Obsidian!")

        except Exception as e:
            self._log(f"❌ Error: {e}")
            self._set_progress(0, "Error al procesar")
        finally:
            self.after(0, lambda: self.process_btn.configure(
                state="normal", text="🚀  Procesar Clase"))

    def _transcribe(self) -> str:
        from faster_whisper import WhisperModel

        self._log("⏳ Cargando modelo Whisper (la primera vez descarga el modelo)...")
        self._set_progress(0.05, "Cargando modelo...")

        model = WhisperModel(self.model_var.get(), device="cuda", compute_type="int8")

        self._log(f"🎙️ Transcribiendo con {self.model_var.get()}...")
        self._set_progress(0.1, "Transcribiendo audio...")

        segments, info = model.transcribe(self.audio_path, language="es", beam_size=5)

        parts = []
        for seg in segments:
            parts.append(seg.text.strip())
            if info.duration > 0:
                pct = 0.1 + (seg.end / info.duration) * 0.5
                mins_done = int(seg.end / 60)
                secs_done = int(seg.end % 60)
                total_mins = int(info.duration / 60)
                total_secs = int(info.duration % 60)
                self._set_progress(
                    min(pct, 0.6),
                    f"Transcribiendo... {mins_done}:{secs_done:02d} / {total_mins}:{total_secs:02d}"
                )

        transcription = " ".join(parts)
        self._log(f"✅ Transcripción completa — {len(transcription.split())} palabras")
        self._set_progress(0.65, "Procesando con Claude AI...")
        return transcription

    def _call_claude(self, transcription: str, asignatura: str, fecha: str) -> str:
        self._log("🤖 Enviando a Claude AI para estructurar el documento...")

        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

        prompt = f"""Eres un asistente académico experto. Analiza la transcripción de una clase universitaria y genera un documento en Markdown con EXACTAMENTE estas 5 secciones:

## 📋 Ficha Técnica
- **Fecha:** {fecha}
- **Asignatura:** {asignatura}
- **Tema Principal:** [extrae el tema central]
- **Objetivo de la Clase:** [extrae o infiere el objetivo principal]

## 📝 Resumen Ejecutivo
[Máximo 5 párrafos. Explica el contexto y el núcleo de la clase de forma comprensible para alguien que no asistió. Incluye el hilo conductor.]

## 🔑 Conceptos Clave y Definiciones
[Para cada concepto importante usa este formato:]

**Término:** Definición clara y concisa.
> Ejemplo: [si el profesor dio un ejemplo, inclúyelo]

## 📊 Marco Lógico
[Si la clase presenta un modelo de negocio, estrategia empresarial o política pública, desarrolla:]
- **Objetivo:** [qué se quiere lograr]
- **Medios:** [cómo se logrará]
- **Resultados:** [qué se espera obtener]
- **Supuestos:** [condiciones asumidas como ciertas]
[Si NO aplica, escribe únicamente: "No aplica para esta clase."]

## ❓ Preguntas de Examen
[Entre 5 y 8 preguntas desafiantes y específicas. Incluye preguntas de análisis, no solo memorización.]

1.
2.
...

---
TRANSCRIPCIÓN:
{transcription}"""

        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        self._set_progress(0.9, "Guardando en Obsidian...")
        return msg.content[0].text

    def _save(self, content: str, asignatura: str, fecha: str):
        vault = os.getenv("OBSIDIAN_VAULT", r"C:\Users\vicen\Desktop\claude\Claude")
        os.makedirs(vault, exist_ok=True)
        filename = f"{fecha} - {asignatura}.md"
        filepath = os.path.join(vault, filename)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        self._log(f"💾 Guardado: {filename}")


if __name__ == "__main__":
    app = ClaseObsidianApp()
    app.mainloop()
