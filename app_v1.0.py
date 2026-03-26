import os
import re
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style
from PyPDF2 import PdfReader
from openpyxl import Workbook
import webbrowser

# ── Paleta de cores ──────────────────────────────────────────────
BG        = "#0e0f13"
PANEL     = "#16181f"
BORDER    = "#2a2d38"
ACCENT    = "#4d9fff"       # azul vibrante
ACCENT2   = "#ff6b35"       # laranja para alertas / cancelar
TEXT_PRI  = "#e8eaf0"
TEXT_SEC  = "#6b7080"
TEXT_OK   = "#4d9fff"
TEXT_ERR  = "#ff5252"
FONT_MONO = ("Courier New", 9)
FONT_UI   = ("Segoe UI", 10)
FONT_HEAD = ("Segoe UI Semibold", 11)
FONT_BIG  = ("Segoe UI Bold", 13)


def rounded_rect(canvas, x1, y1, x2, y2, r, **kwargs):
    canvas.create_arc(x1, y1, x1+2*r, y1+2*r, start=90,  extent=90,  style="pieslice", **kwargs)
    canvas.create_arc(x2-2*r, y1, x2, y1+2*r, start=0,   extent=90,  style="pieslice", **kwargs)
    canvas.create_arc(x1, y2-2*r, x1+2*r, y2, start=180, extent=90,  style="pieslice", **kwargs)
    canvas.create_arc(x2-2*r, y2-2*r, x2, y2, start=270, extent=90,  style="pieslice", **kwargs)
    canvas.create_rectangle(x1+r, y1, x2-r, y2, **kwargs)
    canvas.create_rectangle(x1, y1+r, x2, y2-r, **kwargs)


class FlatButton(tk.Canvas):
    """Botão flat customizado com hover e acento colorido."""
    def __init__(self, master, text="", command=None, color=ACCENT,
                 fg=BG, width=120, height=34, **kw):
        super().__init__(master, width=width, height=height,
                         bg=PANEL, bd=0, highlightthickness=0, cursor="hand2", **kw)
        self.cmd     = command
        self.color   = color
        self.fg      = fg
        self.text    = text
        self.w       = width
        self.h       = height
        self._draw(self.color)
        self.bind("<Enter>",    self._on_enter)
        self.bind("<Leave>",    self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _draw(self, fill):
        self.delete("all")
        self.create_rectangle(0, 0, self.w, self.h, fill=fill, outline=fill)
        self.create_text(self.w//2, self.h//2, text=self.text,
                         fill=self.fg, font=("Segoe UI Semibold", 9))

    def _on_enter(self, _):
        # clareia ligeiramente
        self._draw(self._lighten(self.color))

    def _on_leave(self, _):
        self._draw(self.color)

    def _on_click(self, _):
        if self.cmd:
            self.cmd()

    @staticmethod
    def _lighten(hex_color, factor=0.2):
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r = min(255, int(r + (255 - r) * factor))
        g = min(255, int(g + (255 - g) * factor))
        b = min(255, int(b + (255 - b) * factor))
        return f"#{r:02x}{g:02x}{b:02x}"


class PathSelector(tk.Frame):
    """Widget de seleção de pasta com ícone e path truncado."""
    def __init__(self, master, label_text, command, **kw):
        super().__init__(master, bg=PANEL, **kw)
        self.command = command

        # label título
        tk.Label(self, text=label_text, bg=PANEL, fg=TEXT_SEC,
                 font=("Segoe UI", 8)).pack(anchor="w")

        row = tk.Frame(self, bg=PANEL)
        row.pack(fill="x", pady=(2, 0))

        # campo de exibição
        self.display = tk.Label(row, text="Nenhuma pasta selecionada",
                                bg="#1e2029", fg=TEXT_SEC,
                                font=FONT_MONO, anchor="w",
                                padx=8, pady=6, width=44, relief="flat")
        self.display.pack(side="left", fill="x", expand=True)

        btn = tk.Button(row, text="ABRIR", command=self.command,
                        bg=BORDER, fg=TEXT_PRI, activebackground="#3a3d48",
                        activeforeground=TEXT_PRI, relief="flat", bd=0,
                        font=("Segoe UI Semibold", 9), cursor="hand2",
                        padx=10, pady=6)
        btn.pack(side="left", padx=(4, 0))

    def set_path(self, path):
        if path:
            short = path if len(path) <= 48 else "…" + path[-45:]
            self.display.config(text=short, fg=ACCENT)
        else:
            self.display.config(text="Nenhuma pasta selecionada", fg=TEXT_SEC)


class PDFSearcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Texto em PDF")
        self.root.geometry("660x680")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.root.minsize(660, 680)

        self.pasta_pdfs  = ""
        self.pasta_saida = ""
        self.regex       = None
        self.running     = False
        self.paused      = False
        self.cancelled   = False
        self.pdfs        = []
        self.palavras    = []
        self.wb          = None
        self.inicio_total = 0

        self._build_ui()

    # ── Construção da interface ──────────────────────────────────

    def _build_ui(self):
        # ── Cabeçalho ────────────────────────────────────────────
        header = tk.Frame(self.root, bg=PANEL, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="⬡  BUSCADOR DE TEXTO EM PDF", bg=PANEL,
                 fg=ACCENT, font=("Courier New", 14, "bold")).pack(side="left", padx=20)
        tk.Label(header, text="v1.0  ·  by Eng. Allan Tafuri", bg=PANEL,
                 fg=TEXT_SEC, font=("Segoe UI", 8)).pack(side="right", padx=20)
        

        sep = tk.Frame(self.root, bg=ACCENT, height=2)
        sep.pack(fill="x")

        # ── Botões fixos na base (devem ser empacotados ANTES do body) ───
        tk.Frame(self.root, bg=BORDER, height=1).pack(side="bottom", fill="x")
        btn_frame = tk.Frame(self.root, bg=PANEL, pady=10, padx=16)
        btn_frame.pack(side="bottom", fill="x")
        for i in range(4):
            btn_frame.columnconfigure(i, weight=1)

        btn_cfg_primary = dict(bg=ACCENT, fg=BG, activebackground="#7ab8ff",
                               activeforeground=BG, relief="flat", bd=0,
                               font=("Segoe UI Semibold", 10), cursor="hand2", pady=10)
        btn_cfg_neutral = dict(bg="#2a3540", fg=TEXT_PRI, activebackground="#3a4550",
                               activeforeground=TEXT_PRI, relief="flat", bd=0,
                               font=("Segoe UI Semibold", 10), cursor="hand2", pady=10)
        btn_cfg_cancel  = dict(bg=ACCENT2, fg="white", activebackground="#ff8c5a",
                               activeforeground="white", relief="flat", bd=0,
                               font=("Segoe UI Semibold", 10), cursor="hand2", pady=10)

        tk.Button(btn_frame, text="▶  Iniciar",    command=self.iniciar,   **btn_cfg_primary).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        tk.Button(btn_frame, text="⏸  Pausar",     command=self.pausar,    **btn_cfg_neutral).grid(row=0, column=1, sticky="ew", padx=4)
        tk.Button(btn_frame, text="▶▶  Continuar", command=self.continuar, **btn_cfg_neutral).grid(row=0, column=2, sticky="ew", padx=4)
        tk.Button(btn_frame, text="✕  Cancelar",   command=self.cancelar,  **btn_cfg_cancel ).grid(row=0, column=3, sticky="ew", padx=(4, 0))

        # ── Corpo principal ───────────────────────────────────────
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=18)

        # ── Seção: palavras-chave ─────────────────────────────────
        self._section_label(body, "01  PALAVRAS-CHAVE")

        kw_frame = tk.Frame(body, bg=PANEL, pady=10, padx=12)
        kw_frame.pack(fill="x", pady=(6, 14))

        tk.Label(kw_frame, text="Termos (separados por  ;)",
                 bg=PANEL, fg=TEXT_SEC, font=("Segoe UI", 8)).pack(anchor="w")

        self.entry = tk.Entry(kw_frame, bg="#1e2029", fg=TEXT_PRI,
                              insertbackground=ACCENT, relief="flat",
                              font=("Courier New", 11), bd=0)
        self.entry.pack(fill="x", pady=(4, 2), ipady=7)

        tk.Label(kw_frame, text="ex:  umi ; topo ; geofísica",
                 bg=PANEL, fg=TEXT_SEC, font=("Segoe UI", 8)).pack(anchor="w")

        # borda inferior sutil no entry
        tk.Frame(kw_frame, bg=BORDER, height=1).pack(fill="x")

        # ── Seção: pastas ─────────────────────────────────────────
        self._section_label(body, "02  PASTAS")

        folders = tk.Frame(body, bg=PANEL, pady=12, padx=12)
        folders.pack(fill="x", pady=(6, 14))

        self.ps_pdfs = PathSelector(folders, "BUSCA — pasta com os PDFs",
                                    self.selecionar_pasta_pdfs)
        self.ps_pdfs.pack(fill="x", pady=(0, 10))

        self.ps_saida = PathSelector(folders, "SAÍDA — onde salvar o .xlsx",
                                     self.selecionar_pasta_saida)
        self.ps_saida.pack(fill="x")

        # ── Seção: progresso ──────────────────────────────────────
        self._section_label(body, "03  PROGRESSO")

        prog_frame = tk.Frame(body, bg=PANEL, pady=10, padx=12)
        prog_frame.pack(fill="x", pady=(6, 14))

        # Barra de progresso estilizada via ttk.Style
        style = Style()
        style.theme_use("default")
        style.configure("custom.Horizontal.TProgressbar",
                         troughcolor="#1e2029",
                         background=ACCENT,
                         bordercolor=PANEL,
                         lightcolor=ACCENT,
                         darkcolor=ACCENT,
                         thickness=10)

        self.progress_var = tk.DoubleVar()
        self.progress = Progressbar(prog_frame, variable=self.progress_var,
                                    orient="horizontal", length=598,
                                    mode="determinate",
                                    style="custom.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 8))

        # Info linha
        info_row = tk.Frame(prog_frame, bg=PANEL)
        info_row.pack(fill="x")
        self.lbl_arquivo = tk.Label(info_row, text="—", bg=PANEL,
                                    fg=TEXT_SEC, font=FONT_MONO, anchor="w")
        self.lbl_arquivo.pack(side="left")
        self.lbl_pct = tk.Label(info_row, text="0 %", bg=PANEL,
                                fg=ACCENT, font=("Courier New", 9, "bold"), anchor="e")
        self.lbl_pct.pack(side="right")

        # ── Seção: log ────────────────────────────────────────────
        self._section_label(body, "04  LOG")

        log_frame = tk.Frame(body, bg=PANEL, pady=8, padx=8)
        log_frame.pack(fill="both", expand=True, pady=(6, 10))

        self.status_text = tk.Text(log_frame, height=8,
                                   bg="#0b0c10", fg=TEXT_PRI,
                                   insertbackground=ACCENT,
                                   font=FONT_MONO, relief="flat",
                                   bd=0, wrap="word",
                                   selectbackground=ACCENT, selectforeground=BG)
        sb = tk.Scrollbar(log_frame, command=self.status_text.yview, bg=PANEL)
        self.status_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.status_text.pack(fill="both", expand=True)

        # Tags de cor no log
        self.status_text.tag_config("ok",    foreground=TEXT_OK)
        self.status_text.tag_config("err",   foreground=TEXT_ERR)
        self.status_text.tag_config("warn",  foreground=ACCENT2)
        self.status_text.tag_config("info",  foreground=TEXT_PRI)
        self.status_text.tag_config("muted", foreground=TEXT_SEC)

        # (botões ficam fora do body — ver abaixo)

    def _section_label(self, parent, text):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(0, 0))
        tk.Label(f, text=text, bg=BG, fg=TEXT_SEC,
                 font=("Courier New", 8, "bold")).pack(side="left")
        tk.Frame(f, bg=BORDER, height=1).pack(side="left", fill="x",
                                               expand=True, padx=(8, 0), pady=1)

    # ── Seleção de pastas ────────────────────────────────────────

    def selecionar_pasta_pdfs(self):
        pasta = filedialog.askdirectory(title="Pasta com os PDFs")
        self.pasta_pdfs = pasta or ""
        self.ps_pdfs.set_path(self.pasta_pdfs)

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Pasta de saída")
        self.pasta_saida = pasta or ""
        self.ps_saida.set_path(self.pasta_saida)

    # ── Controles ────────────────────────────────────────────────

    def pausar(self):
        if not self.running:
            self._log("⚠ Busca não está em execução.", "warn"); return
        if self.paused:
            self._log("⚠ Já está pausado.", "warn"); return
        self.paused = True
        self._log("⏸ Pausado.", "warn")

    def continuar(self):
        if not self.running:
            self._log("⚠ Busca não está em execução.", "warn"); return
        if not self.paused:
            self._log("⚠ Não está pausado.", "warn"); return
        self.paused = False
        self._log("▶ Continuando…", "ok")

    def cancelar(self):
        if not self.running:
            self._log("⚠ Nenhuma busca em andamento.", "warn"); return
        self.cancelled = True
        self._log("✕ Cancelado pelo usuário.", "err")

    def iniciar(self):
        if self.running:
            self._log("⚠ Busca já em andamento.", "warn"); return

        entrada = self.entry.get().strip()
        if not entrada:
            messagebox.showwarning("Aviso", "Digite ao menos uma palavra-chave."); return
        if not self.pasta_pdfs or not self.pasta_saida:
            messagebox.showwarning("Aviso", "Selecione as pastas de entrada e saída."); return

        self.palavras = [p.strip() for p in entrada.split(";") if p.strip()]
        if not self.palavras:
            messagebox.showwarning("Aviso", "Nenhuma palavra válida inserida."); return

        variacoes = set()
        for p in self.palavras:
            variacoes.update([p, p.lower(), p.upper(), p.capitalize(),
                               rf"{p}[a-z]*", rf"{p.lower()}[a-z]*",
                               rf"{p.upper()}[A-Z]*", rf"{p.capitalize()}[a-z]*"])
        try:
            self.regex = re.compile("|".join(variacoes))
        except re.error as e:
            messagebox.showerror("Erro regex", str(e)); return

        self.pdfs = []
        for dirpath, _, arquivos in os.walk(self.pasta_pdfs):
            for nome in arquivos:
                if nome.lower().endswith(".pdf"):
                    self.pdfs.append(os.path.join(dirpath, nome))

        if not self.pdfs:
            messagebox.showinfo("Aviso", "Nenhum PDF encontrado."); return

        self.running   = True
        self.paused    = False
        self.cancelled = False
        self.wb        = Workbook()
        self.inicio_total = time.time()

        self.status_text.config(state="normal")
        self.status_text.delete("1.0", tk.END)
        self._log(f"🔎 Palavras-chave: {', '.join(self.palavras)}", "ok")
        self._log(f"   Variações:      {', '.join(sorted(variacoes))}\n", "muted")

        threading.Thread(target=self.executar_busca, daemon=True).start()

    # ── Execução da busca ────────────────────────────────────────

    def executar_busca(self):
        ws = self.wb.active
        ws.title = "Ocorrências"
        ws.append(["Arquivo PDF", "Caminho", "Página", "Palavra encontrada"])

        total = len(self.pdfs)
        self.progress.config(maximum=total)
        self.progress_var.set(0)

        for idx, caminho_pdf in enumerate(self.pdfs, 1):
            if self.cancelled:
                break
            while self.paused:
                time.sleep(0.5)
                if self.cancelled:
                    break

            nome_arquivo = os.path.basename(caminho_pdf)
            self._log(f"[{idx:>3}/{total}]  {nome_arquivo}", "info")
            self.lbl_arquivo.config(text=nome_arquivo[:70])
            start_time = time.time()

            try:
                leitor = PdfReader(caminho_pdf)
                for i, pagina in enumerate(leitor.pages):
                    texto = pagina.extract_text() or ""
                    for termo in self.regex.findall(texto):
                        ws.append([nome_arquivo, caminho_pdf, i + 1, termo])
                tempo = time.time() - start_time
                self._log(f"         ✓ {tempo:.2f}s\n", "ok")
            except Exception as e:
                self._log(f"         ✗ Erro: {e}\n", "err")

            pct = idx / total * 100
            self.progress_var.set(idx)
            self.lbl_pct.config(text=f"{pct:.0f} %")
            self.root.update_idletasks()

        if not self.cancelled:
            nome_saida = os.path.join(self.pasta_saida, "resultados_palavra.xlsx")
            try:
                self.wb.save(nome_saida)
            except Exception as e:
                self._log(f"✗ Erro ao salvar: {e}", "err")

            tempo_total = time.time() - self.inicio_total
            self._log(f"\n✓ Arquivo salvo: {nome_saida}", "ok")
            self._log(f"⏱ Tempo total:  {tempo_total:.2f}s", "ok")
            self.lbl_arquivo.config(text="Concluído")
            self._show_conclusao(nome_saida, tempo_total)
        else:
            tempo_total = time.time() - self.inicio_total
            self._log(f"\n⚠ Cancelado após {tempo_total:.2f}s. Arquivo não salvo.", "warn")
            self.lbl_arquivo.config(text="Cancelado")
            self._show_conclusao(None, tempo_total, cancelado=True)

        self.running = self.paused = self.cancelled = False

    # ── Modal de conclusão ───────────────────────────────────────

    def _show_conclusao(self, caminho, tempo, cancelado=False):
        win = tk.Toplevel(self.root)
        win.title("Resultado")
        win.geometry("440x260")
        win.configure(bg=PANEL)
        win.resizable(False, False)
        win.grab_set()

        # Linha de destaque no topo
        cor = ACCENT if not cancelado else ACCENT2
        tk.Frame(win, bg=cor, height=3).pack(fill="x")

        icon = "✓" if not cancelado else "✕"
        msg  = "BUSCA CONCLUÍDA" if not cancelado else "BUSCA CANCELADA"
        tk.Label(win, text=f"{icon}  {msg}", bg=PANEL, fg=cor,
                 font=("Courier New", 13, "bold")).pack(pady=(18, 4))

        tk.Label(win, text=f"Tempo:  {tempo:.2f} s", bg=PANEL,
                 fg=TEXT_SEC, font=FONT_MONO).pack()

        if not cancelado and caminho:
            short = caminho if len(caminho) <= 55 else "…" + caminho[-52:]
            tk.Label(win, text=short, bg=PANEL, fg=TEXT_PRI,
                     font=FONT_MONO, wraplength=400).pack(pady=8)

        tk.Label(win, text="by Eng. Allan Tafuri", bg=PANEL,
                 fg=TEXT_SEC, font=("Segoe UI", 8, "italic")).pack(pady=(4, 14))

        tk.Label(win, text="Realizar outra busca?", bg=PANEL,
                 fg=TEXT_PRI, font=FONT_UI).pack()

        row = tk.Frame(win, bg=PANEL)
        row.pack(pady=10)
        tk.Button(row, text="SIM", command=lambda: [self._nova_busca(), win.destroy()],
                  bg=ACCENT, fg=BG, activebackground="#7ab8ff", activeforeground=BG,
                  relief="flat", bd=0, font=("Segoe UI Semibold", 9),
                  cursor="hand2", padx=20, pady=8).pack(side="left", padx=8)
        tk.Button(row, text="NÃO", command=self.root.quit,
                  bg="#2a2d38", fg=TEXT_PRI, activebackground="#3a3d48", activeforeground=TEXT_PRI,
                  relief="flat", bd=0, font=("Segoe UI Semibold", 9),
                  cursor="hand2", padx=20, pady=8).pack(side="left", padx=8)

    def _nova_busca(self):
        self.entry.delete(0, tk.END)
        self.status_text.config(state="normal")
        self.status_text.delete("1.0", tk.END)
        self.progress_var.set(0)
        self.lbl_arquivo.config(text="—")
        self.lbl_pct.config(text="0 %")
        self.pasta_pdfs = self.pasta_saida = ""
        self.pdfs = []
        self.cancelled = self.paused = False
        self.ps_pdfs.set_path("")
        self.ps_saida.set_path("")

    # ── Log utilitário ───────────────────────────────────────────

    def _log(self, texto, tag="info"):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, texto + "\n", tag)
        self.status_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSearcherApp(root)
    root.mainloop()