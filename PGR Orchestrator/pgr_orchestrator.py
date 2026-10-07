"""
PGR Orchestrator v8.3.0
========================
Preenche automaticamente os indicadores de Waste e Consumo no portal PGR.

Novidades v8.3:
- ETA dinâmico calculado a partir da velocidade real dos últimos N fills
- Dual progress bars: Waste e Consumo com barras independentes
- Modo verbose (toggle na GUI): log detalhado de cada XPath/tentativa
- Explicit waits em switch_tab: aguarda celula-base aparecer antes de proceder
- Retry com backoff exponencial em fill_field: 1s → 2s → 4s (3 tentativas)
- Sem terminal (pythonw / pyw): sem janela preta ao iniciar
- GUI redesenhada: dark-theme consistente, ícones, cards com separadores
- Headless removido: inviável com SSO da rede corporativa
- Email de login persistido e auto-preenchido no SSO
"""

import os
import sys
import time
import json
import threading
import queue
import subprocess
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import deque

import xlwings as xw
from tkcalendar import DateEntry
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

VERSAO      = "8.3.0"
CONFIG_FILE = Path.home() / ".pgr_orchestrator_config.json"
SYSTEM_URL  = "https://portal.example.com/cr-digital?crDigitalId=1362"

PROFILE_DIR_WASTE   = os.path.join(os.getcwd(), "edge_profile_waste")
PROFILE_DIR_CONSUMO = os.path.join(os.getcwd(), "edge_profile_consumo")

COLORS = {
    "bg":        "#0d1117",
    "bg2":       "#161b22",
    "border":    "#30363d",
    "text":      "#8b949e",
    "subtext":   "#8b949e",
    "blue":      "#58a6ff",
    "green":     "#3fb950",
    "yellow":    "#d29922",
    "red":       "#f85149",
    "bar_waste":   "#1f6feb",
    "bar_consumo": "#388bfd",
    "bar_trough":  "#21262d",
}

INDICATORS_MAP = {
    "Higiênicos": {
        "items": {
            "Waste | Higiênicos - Consolidado":        (16, "waste"),
            "C.E de Papel | Higiênicos - Consolidado": (16, "consumo"),
            "Waste | PH.04": (8,  "waste"),  "Waste | PH.06": (9,  "waste"),
            "Waste | PH.07": (10, "waste"),  "Waste | PH.08": (11, "waste"),
            "Waste | PH.12": (12, "waste"),  "Waste | PH.13": (13, "waste"),
            "Waste | PH.14": (14, "waste"),  "Waste | PH.15": (15, "waste"),
            "C.E de Papel | PH.04": (8,  "consumo"), "C.E de Papel | PH.06": (9,  "consumo"),
            "C.E de Papel | PH.07": (10, "consumo"), "C.E de Papel | PH.08": (11, "consumo"),
            "C.E de Papel | PH.12": (12, "consumo"), "C.E de Papel | PH.13": (13, "consumo"),
            "C.E de Papel | PH.14": (14, "consumo"), "C.E de Papel | PH.15": (15, "consumo"),
        },
    },
    "Conversão": {
        "items": {
            "Waste | Conversão - Sem Duramax":        (32, "waste"),
            "C.E de Papel | Conversão - Sem Duramax": (32, "consumo"),
        },
    },
    "Duramax": {
        "items": {
            "Waste | DMX.05":        (7, "waste"),
            "C.E de Papel | DMX.05": (7, "consumo"),
        },
    },
    "Professional": {
        "items": {
            "Waste | Professional - Consolidado":        (29, "waste"),
            "C.E de Papel | Professional - Consolidado": (29, "consumo"),
            "Waste | PROF.01": (27, "waste"),   "Waste | PROF.02": (28, "waste"),
            "Waste | PROF.03": (26, "waste"),   "Waste | PROF.04": (25, "waste"),
            "C.E de Papel | PROF.01": (27, "consumo"),  "C.E de Papel | PROF.02": (28, "consumo"),
            "C.E de Papel | PROF.03": (26, "consumo"),  "C.E de Papel | PROF.04": (25, "consumo"),
        },
    },
    "Lenços": {
        "items": {
            "Waste | Lenços - Consolidado":        (22, "waste"),
            "C.E de Papel | Lenços - Consolidado": (22, "consumo"),
            "Waste | FOR.01": (20, "waste"),  "Waste | FOR.02": (21, "waste"),
            "C.E de Papel | FOR.01": (20, "consumo"), "C.E de Papel | FOR.02": (21, "consumo"),
        },
    },
    "Guardanapos": {
        "items": {
            "Waste | Guardanapos - Consolidado":        (19, "waste"),
            "C.E de Papel | Guardanapos - Consolidado": (19, "consumo"),
        },
    },
}


@dataclass
class Estado:
    progresso_atual: int = 0
    progresso_total: int = 100
    status: str = "Aguardando..."
    concluido: bool = False
    erro: str | None = None

    waste_atual:   int = 0
    waste_total:   int = 0
    consumo_atual: int = 0
    consumo_total: int = 0

    _tempos_fill: deque = field(default_factory=lambda: deque(maxlen=20))
    _t_inicio:    float = field(default_factory=time.time)

    log_queue:   queue.Queue = field(default_factory=queue.Queue)
    login_queue: queue.Queue = field(default_factory=queue.Queue)

    verbose: bool = False

    def log(self, msg: str):
        self.log_queue.put(msg)

    def log_verbose(self, msg: str):
        """Só enfileira se verbose estiver ativo — zero overhead em modo normal."""
        if self.verbose:
            self.log_queue.put(f"[DBG] {msg}")

    def avancar(self, tipo: str = "", incremento: int = 1):
        """Avança o progresso global e o da barra específica (waste/consumo)."""
        now = time.time()
        self._tempos_fill.append(now)

        self.progresso_atual = min(self.progresso_atual + incremento, self.progresso_total)

        if tipo == "waste":
            self.waste_atual = min(self.waste_atual + incremento, self.waste_total)
        elif tipo == "consumo":
            self.consumo_atual = min(self.consumo_atual + incremento, self.consumo_total)

    def eta_str(self) -> str:
        """
        ETA baseado na velocidade média dos últimos N fills reais.
        Mais preciso que uma estimativa linear simples porque o PGR tem
        variações de latência por aba/indicador.
        """
        if len(self._tempos_fill) < 2 or self.progresso_atual == 0:
            return "--:--"

        fills_janela = list(self._tempos_fill)
        # Tempo médio entre fills na janela deslizante
        if len(fills_janela) >= 2:
            tempo_por_fill = (fills_janela[-1] - fills_janela[0]) / (len(fills_janela) - 1)
        else:
            tempo_por_fill = (time.time() - self._t_inicio) / max(self.progresso_atual, 1)

        restantes = max(self.progresso_total - self.progresso_atual, 0)
        segundos  = int(tempo_por_fill * restantes)

        if segundos > 3600:
            return f"{segundos // 3600}h{(segundos % 3600) // 60:02d}m"
        m, s = divmod(segundos, 60)
        return f"{m:02d}:{s:02d}"

    def solicitar_login(self, perfil: str):
        ev = threading.Event()
        self.login_queue.put((perfil, ev))
        ev.wait()


class JanelaProgresso:
    POLL_MS = 200  # Mais responsiva que 250ms

    def __init__(self, titulo: str, subtitulo: str, estado: Estado, dual: bool = False):
        self._estado = estado
        self._dual   = dual   # True quando há 2 drivers em paralelo
        self._t0     = time.time()

        C = COLORS

        self.root = tk.Tk()
        self.root.title(f"PGR Orchestrator v{VERSAO}")
        self.root.geometry("700x520" if dual else "700x460")
        self.root.resizable(False, False)
        self.root.configure(bg=C["bg"])
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        hdr = tk.Frame(self.root, bg=C["bg"], pady=12)
        hdr.pack(fill="x", padx=20)

        tk.Label(hdr, text="⚙  PGR Orchestrator", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI", 14, "bold"), anchor="w").pack(side="left")
        tk.Label(hdr, text=f"v{VERSAO}", bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 9), anchor="e").pack(side="right", padx=(0, 2))

        tk.Label(self.root, text=subtitulo, bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 8)).pack(anchor="w", padx=20)
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", pady=(8, 0))

        body = tk.Frame(self.root, bg=C["bg"], padx=20, pady=10)
        body.pack(fill="both", expand=True)

        self._lbl_status = tk.Label(body, text="Iniciando...", bg=C["bg"], fg=C["text"],
                                    font=("Segoe UI", 10, "bold"), anchor="w")
        self._lbl_status.pack(fill="x", pady=(0, 8))

        style = ttk.Style()
        style.theme_use("clam")

        if dual:
            tk.Label(body, text="WASTE", bg=C["bg"], fg=C["subtext"],
                     font=("Segoe UI", 7, "bold"), anchor="w").pack(fill="x")
            style.configure("Waste.Horizontal.TProgressbar",
                            troughcolor=C["bar_trough"], background=C["bar_waste"],
                            bordercolor=C["border"], lightcolor=C["bar_waste"],
                            darkcolor=C["bar_waste"], thickness=12)
            self._bar_waste = ttk.Progressbar(body, style="Waste.Horizontal.TProgressbar",
                                               length=660, mode="determinate")
            self._bar_waste.pack(fill="x", pady=(0, 2))
            self._lbl_waste = tk.Label(body, text="0 / 0", bg=C["bg"], fg=C["subtext"],
                                       font=("Segoe UI", 7), anchor="e")
            self._lbl_waste.pack(fill="x", pady=(0, 6))

            tk.Label(body, text="CONSUMO ESPECÍFICO", bg=C["bg"], fg=C["subtext"],
                     font=("Segoe UI", 7, "bold"), anchor="w").pack(fill="x")
            style.configure("Consumo.Horizontal.TProgressbar",
                            troughcolor=C["bar_trough"], background=C["bar_consumo"],
                            bordercolor=C["border"], lightcolor=C["bar_consumo"],
                            darkcolor=C["bar_consumo"], thickness=12)
            self._bar_consumo = ttk.Progressbar(body, style="Consumo.Horizontal.TProgressbar",
                                                 length=660, mode="determinate")
            self._bar_consumo.pack(fill="x", pady=(0, 2))
            self._lbl_consumo = tk.Label(body, text="0 / 0", bg=C["bg"], fg=C["subtext"],
                                          font=("Segoe UI", 7), anchor="e")
            self._lbl_consumo.pack(fill="x", pady=(0, 6))
        else:
            style.configure("PGR.Horizontal.TProgressbar",
                            troughcolor=C["bar_trough"], background=C["bar_waste"],
                            bordercolor=C["border"], lightcolor=C["bar_waste"],
                            darkcolor=C["bar_waste"], thickness=14)
            self._bar = ttk.Progressbar(body, style="PGR.Horizontal.TProgressbar",
                                        length=660, mode="determinate")
            self._bar.pack(fill="x", pady=(0, 3))

        metrics = tk.Frame(body, bg=C["bg"])
        metrics.pack(fill="x", pady=(4, 8))

        self._lbl_cnt  = tk.Label(metrics, text="0 / 0",  bg=C["bg"], fg=C["subtext"], font=("Segoe UI", 8))
        self._lbl_eta  = tk.Label(metrics, text="ETA --:--", bg=C["bg"], fg=C["blue"],   font=("Segoe UI", 8, "bold"))
        self._lbl_time = tk.Label(metrics, text="00:00:00", bg=C["bg"], fg=C["subtext"], font=("Segoe UI", 8))

        self._lbl_cnt.pack(side="left")
        self._lbl_eta.pack(side="left", padx=16)
        self._lbl_time.pack(side="right")

        tk.Frame(body, bg=C["border"], height=1).pack(fill="x", pady=(0, 6))

        self._log = tk.Text(body, height=11, state="disabled",
                            bg=C["bg2"], fg="#c9d1d9",
                            font=("Consolas", 8), relief="flat",
                            insertbackground=C["text"], padx=6, pady=4,
                            selectbackground="#264f78")
        self._log.pack(fill="both", expand=True)

        self._log.tag_configure("ok",    foreground=C["green"])
        self._log.tag_configure("warn",  foreground=C["yellow"])
        self._log.tag_configure("error", foreground=C["red"])
        self._log.tag_configure("info",  foreground=C["blue"])
        self._log.tag_configure("debug", foreground="#6e7681")
        self._log.tag_configure("dim",   foreground=C["subtext"])

    def _append_log(self, msg: str):
        self._log.config(state="normal")
        ts  = datetime.now().strftime("%H:%M:%S")
        tag = ("ok"    if msg.startswith("[OK]")
               else "error" if any(msg.startswith(p) for p in ("[FALHA]", "[ERRO]"))
               else "warn"  if msg.startswith("[!]")
               else "info"  if msg.startswith("[→]")
               else "debug" if msg.startswith("[DBG]")
               else "dim")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _poll(self):
        e = self._estado

        if self._dual and e.progresso_total > 0:
            if e.waste_total > 0:
                self._bar_waste["value"] = (e.waste_atual / e.waste_total) * 100
                self._lbl_waste.config(text=f"{e.waste_atual} / {e.waste_total}")
            if e.consumo_total > 0:
                self._bar_consumo["value"] = (e.consumo_atual / e.consumo_total) * 100
                self._lbl_consumo.config(text=f"{e.consumo_atual} / {e.consumo_total}")
        elif not self._dual and e.progresso_total > 0:
            self._bar["value"] = (e.progresso_atual / e.progresso_total) * 100

        self._lbl_status.config(text=e.status)
        self._lbl_cnt.config(text=f"{e.progresso_atual} / {e.progresso_total}")
        self._lbl_eta.config(text=f"ETA {e.eta_str()}")

        elapsed = int(time.time() - self._t0)
        h, rem = divmod(elapsed, 3600)
        m, s   = divmod(rem, 60)
        self._lbl_time.config(text=f"{h:02d}:{m:02d}:{s:02d}")

        try:
            while True:
                self._append_log(e.log_queue.get_nowait())
        except queue.Empty:
            pass

        # Login request da worker thread (nunca chama messagebox de dentro da thread)
        try:
            perfil, evento = e.login_queue.get_nowait()
            messagebox.showinfo(
                "Login necessário",
                f"Faça login no Edge ({perfil})\ne clique OK quando estiver na tela principal do PGR.",
            )
            evento.set()
        except queue.Empty:
            pass

        if e.concluido:
            if e.erro:
                self._lbl_status.config(fg=COLORS["red"])
            else:
                self._lbl_status.config(fg=COLORS["green"])
            self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
            self.root.after(2500, self.root.destroy)
        else:
            self.root.after(self.POLL_MS, self._poll)

    def iniciar(self, worker_fn, *args):
        threading.Thread(target=worker_fn, args=args, daemon=True).start()
        self.root.after(self.POLL_MS, self._poll)
        self.root.mainloop()


def notificar_windows(titulo: str, mensagem: str, icone: str = "Info"):
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$n=New-Object System.Windows.Forms.NotifyIcon;"
        "$n.Icon=[System.Drawing.SystemIcons]::Information;"
        "$n.Visible=$true;"
        f"$n.ShowBalloonTip(6000,'{titulo}','{mensagem}',"
        f"[System.Windows.Forms.ToolTipIcon]::{icone});"
        "Start-Sleep -Milliseconds 6500;$n.Dispose()"
    )
    subprocess.Popen(["powershell", "-WindowStyle", "Hidden", "-Command", script],
                     creationflags=subprocess.CREATE_NO_WINDOW)

def carregar_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

class ConfiguracaoGUI:
    """
    GUI dark-theme da tela de configuração.
    Usa tk puro com estilos manuais (sem ttk.LabelFrame) para controle
    total das cores — ttk no Windows não respeita bg em todos os widgets.
    """

    def __init__(self, config_salva: dict):
        self.resultado: dict | None = None
        self._cfg = config_salva
        self._build()

    def _card(self, parent, titulo: str) -> tk.Frame:
        C = COLORS
        outer = tk.Frame(parent, bg=C["bg2"], bd=0, relief="flat",
                         highlightbackground=C["border"], highlightthickness=1)
        outer.pack(fill="x", padx=16, pady=4)
        tk.Label(outer, text=titulo, bg=C["bg2"], fg=C["blue"],
                 font=("Segoe UI", 8, "bold"), anchor="w",
                 padx=10, pady=4).pack(fill="x")
        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x")
        inner = tk.Frame(outer, bg=C["bg2"], padx=10, pady=8)
        inner.pack(fill="x")
        return inner

    def _entry_dark(self, parent, textvariable, width=50, readonly=False) -> ttk.Entry:
        style = ttk.Style()
        style.configure("Dark.TEntry",
                        fieldbackground=COLORS["bg"],
                        foreground=COLORS["text"],
                        insertcolor=COLORS["text"],
                        bordercolor=COLORS["border"],
                        lightcolor=COLORS["border"],
                        darkcolor=COLORS["border"])
        state = "readonly" if readonly else "normal"
        e = ttk.Entry(parent, textvariable=textvariable,
                      width=width, style="Dark.TEntry", state=state)
        return e

    def _btn(self, parent, text, command, primary=False, width=12) -> tk.Button:
        C = COLORS
        bg  = C["blue"] if primary else C["bg2"]
        fg  = C["bg"]   if primary else C["text"]
        abg = "#1158c7"  if primary else C["border"]
        b = tk.Button(parent, text=text, command=command,
                      bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                      font=("Segoe UI", 9, "bold" if primary else "normal"),
                      relief="flat", bd=0, padx=10, pady=5,
                      cursor="hand2", width=width)
        return b

    def _build(self):
        C = COLORS
        self.root = tk.Tk()
        self.root.title(f"PGR Orchestrator v{VERSAO}")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        hoje = datetime.now()

        hdr = tk.Frame(self.root, bg=C["bg"], pady=14)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="⚙  PGR Orchestrator", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI", 14, "bold"), anchor="w").pack(side="left")
        tk.Label(hdr, text=f"v{VERSAO}", bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="right", padx=4)
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")

        c1 = self._card(self.root, "📊  Planilha Base")
        self._excel = tk.StringVar(value=self._cfg.get("excel_path", ""))
        row1 = tk.Frame(c1, bg=C["bg2"])
        row1.pack(fill="x")
        self._entry_dark(row1, self._excel, width=52, readonly=True).pack(side="left", padx=(0, 8))
        self._btn(row1, "Buscar…", self._sel_excel).pack(side="left")

        c2 = self._card(self.root, "🔐  Login PGR  —  usuário de rede ou e-mail")
        row2 = tk.Frame(c2, bg=C["bg2"])
        row2.pack(fill="x")
        self._email = tk.StringVar(value=self._cfg.get("email", ""))
        self._entry_dark(row2, self._email, width=38).pack(side="left", padx=(0, 10))
        tk.Label(row2, text="Preenchido automaticamente na tela de login.",
                 bg=C["bg2"], fg=C["subtext"], font=("Segoe UI", 8)).pack(side="left")

        c3 = self._card(self.root, "📅  Período de Preenchimento")
        row3 = tk.Frame(c3, bg=C["bg2"])
        row3.pack(fill="x")

        tk.Label(row3, text="De:", bg=C["bg2"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        de_opts = dict(width=12, background="#1f6feb", foreground="white",
                       borderwidth=0, date_pattern="dd/MM/yyyy", locale="pt_BR",
                       font=("Segoe UI", 9))
        self._de_ini = DateEntry(row3, year=hoje.year, month=hoje.month, day=1, **de_opts)
        self._de_ini.pack(side="left", padx=(0, 20))

        tk.Label(row3, text="Até:", bg=C["bg2"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self._de_fim = DateEntry(row3, year=hoje.year, month=hoje.month, day=hoje.day, **de_opts)
        self._de_fim.pack(side="left", padx=(0, 20))

        tk.Label(row3, text="Lançar em:", bg=C["bg2"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(8, 4))
        self._periodo = tk.StringVar(value=self._cfg.get("periodo", "atual"))
        for val, lbl in [("atual", "Mês Atual"), ("anterior", "Mês Anterior")]:
            rb = tk.Radiobutton(row3, text=lbl, variable=self._periodo, value=val,
                                bg=C["bg2"], fg=C["text"], selectcolor=C["bg"],
                                activebackground=C["bg2"], activeforeground=C["blue"],
                                font=("Segoe UI", 9), bd=0)
            rb.pack(side="left", padx=4)

        c4 = self._card(self.root, "🏭  Categorias Operacionais")
        self._cats: dict[str, tk.BooleanVar] = {}
        cats = list(INDICATORS_MAP.keys())
        grid = tk.Frame(c4, bg=C["bg2"])
        grid.pack(fill="x")
        cols = 3
        for i, cat in enumerate(cats):
            v = tk.BooleanVar(value=True)
            self._cats[cat] = v
            cb = tk.Checkbutton(grid, text=cat, variable=v,
                                 bg=C["bg2"], fg=C["text"],
                                 selectcolor=C["bg"],
                                 activebackground=C["bg2"],
                                 activeforeground=C["blue"],
                                 font=("Segoe UI", 9), bd=0)
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=12, pady=2)

        btn_row = tk.Frame(c4, bg=C["bg2"])
        btn_row.pack(fill="x", pady=(6, 0))
        self._btn(btn_row, "+ Todas", lambda: self._toggle(True),  width=10).pack(side="left", padx=(0, 6))
        self._btn(btn_row, "− Nenhuma", lambda: self._toggle(False), width=10).pack(side="left")

        c5 = self._card(self.root, "⚡  Configuração de Execução")
        row5 = tk.Frame(c5, bg=C["bg2"])
        row5.pack(fill="x")

        col_dados = tk.Frame(row5, bg=C["bg2"])
        col_dados.pack(side="left", padx=(0, 30))
        tk.Label(col_dados, text="DADOS", bg=C["bg2"], fg=C["blue"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w")
        self._tipo = tk.StringVar(value=self._cfg.get("tipo", "all"))
        for val, lbl in [("all", "Waste + Consumo"), ("waste", "Só Waste"), ("consumo", "Só Consumo")]:
            tk.Radiobutton(col_dados, text=lbl, variable=self._tipo, value=val,
                           bg=C["bg2"], fg=C["text"], selectcolor=C["bg"],
                           activebackground=C["bg2"], activeforeground=C["blue"],
                           font=("Segoe UI", 9), bd=0).pack(anchor="w")

        col_drivers = tk.Frame(row5, bg=C["bg2"])
        col_drivers.pack(side="left", padx=(0, 30))
        tk.Label(col_drivers, text="DRIVERS EDGE", bg=C["bg2"], fg=C["blue"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w")
        self._n_drivers = tk.IntVar(value=self._cfg.get("n_drivers", 2))
        for val, lbl in [(1, "1 driver  (sequencial)"), (2, "2 drivers (paralelo)")]:
            tk.Radiobutton(col_drivers, text=lbl, variable=self._n_drivers, value=val,
                           bg=C["bg2"], fg=C["text"], selectcolor=C["bg"],
                           activebackground=C["bg2"], activeforeground=C["blue"],
                           font=("Segoe UI", 9), bd=0).pack(anchor="w")

        col_extra = tk.Frame(row5, bg=C["bg2"])
        col_extra.pack(side="left")
        tk.Label(col_extra, text="OPÇÕES", bg=C["bg2"], fg=C["blue"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w")
        self._verbose = tk.BooleanVar(value=self._cfg.get("verbose", False))
        tk.Checkbutton(col_extra, text="Modo verbose (log detalhado)",
                       variable=self._verbose,
                       bg=C["bg2"], fg=C["text"], selectcolor=C["bg"],
                       activebackground=C["bg2"], activeforeground=C["blue"],
                       font=("Segoe UI", 9), bd=0).pack(anchor="w")

        # ── Botões ───────────────────────────────────────────
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", pady=(8, 0))
        footer = tk.Frame(self.root, bg=C["bg"], pady=12)
        footer.pack(fill="x", padx=16)
        self._btn(footer, "Cancelar", self._cancelar, width=12).pack(side="right", padx=(8, 0))
        self._btn(footer, "▶  Executar", self._confirmar, primary=True, width=14).pack(side="right")

        self.root.mainloop()

    def _sel_excel(self):
        p = filedialog.askopenfilename(
            title="Planilha de Análise",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb")],
            parent=self.root,
        )
        if p:
            self._excel.set(p)

    def _toggle(self, v: bool):
        for var in self._cats.values():
            var.set(v)

    def _cancelar(self):
        self.root.destroy()

    def _confirmar(self):
        if not self._excel.get():
            messagebox.showerror("Erro", "Selecione um arquivo Excel.", parent=self.root)
            return
        cats = [c for c, v in self._cats.items() if v.get()]
        if not cats:
            messagebox.showerror("Erro", "Selecione ao menos uma categoria.", parent=self.root)
            return
        dt_ini = self._de_ini.get_date()
        dt_fim = self._de_fim.get_date()
        if dt_ini > dt_fim:
            messagebox.showerror("Erro", "Data inicial não pode ser posterior à data final.", parent=self.root)
            return

        n_drivers = self._n_drivers.get()
        tipo      = self._tipo.get()
        if tipo != "all":
            n_drivers = 1

        self.resultado = {
            "excel_path":    self._excel.get(),
            "email":         self._email.get().strip(),
            "mes":           dt_ini.month,
            "ano":           dt_ini.year,
            "is_prev_month": self._periodo.get() == "anterior",
            "dia_ini":       dt_ini.day,
            "dia_fim":       dt_fim.day,
            "categorias":    cats,
            "tipo":          tipo,
            "n_drivers":     n_drivers,
            "verbose":       self._verbose.get(),
            "periodo_label": f"{dt_ini.strftime('%d/%m/%Y')} → {dt_fim.strftime('%d/%m/%Y')}",
        }
        self.root.destroy()




class ExcelEngine:
    def __init__(self, file_path: str):
        self.file_path    = Path(file_path).resolve()
        self.app          = None
        self.wb           = None
        self.sheet        = None
        self._nos_abrimos = False

    def abrir(self):
        for app_c in xw.apps:
            for book in app_c.books:
                try:
                    if Path(book.fullname).resolve() == self.file_path:
                        self.app = app_c
                        self.wb  = book
                        break
                except Exception:
                    continue
            if self.wb:
                break
        if not self.wb:
            self.app = xw.App(visible=False, add_book=False)
            self.app.screen_updating = False
            self.app.display_alerts  = False
            self.wb  = self.app.books.open(str(self.file_path))
            self._nos_abrimos = True
        self.sheet = self.wb.sheets["Indicador Conversão"]

    def extrair_dados(self, dia_ini: int, dia_fim: int,
                      linhas: set[int], mes: int, ano: int) -> dict:
        cache: dict = {}
        for dia in range(dia_ini, dia_fim + 1):
            self.sheet.range("I3").value = datetime(ano, mes, dia)
            self.app.api.Calculate()
            timeout = 0
            while self.app.api.CalculationState != 0 and timeout < 50:
                time.sleep(0.1)
                timeout += 1
            raw = self.sheet.range("A1:P40").value
            cache[dia] = {}
            for row_num in linhas:
                row   = raw[row_num - 1]
                w_raw = row[6]
                w_fmt = "0,00"
                if isinstance(w_raw, (int, float)):
                    v     = w_raw * 100 if w_raw < 1.0 else w_raw
                    w_fmt = f"{v:.2f}".replace(".", ",")
                c_raw = row[15]
                c_fmt = "0,000"
                if isinstance(c_raw, (int, float)):
                    c_fmt = f"{c_raw:.3f}".replace(".", ",")
                cache[dia][row_num] = {"waste": w_fmt, "consumo": c_fmt}
        return cache

    def fechar(self):
        if not self._nos_abrimos:
            return
        if self.wb:
            self.wb.close()
        if self.app:
            self.app.quit()




class WebAutomator:
    def __init__(self, profile_dir: str):
        self.profile_dir = profile_dir
        self.driver = None
        self.wait   = None

    def start_browser(self):
        if self.driver:
            return
        opts = Options()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        opts.add_argument(f"user-data-dir={self.profile_dir}")
        self.driver = webdriver.Edge(options=opts)
        self.wait   = WebDriverWait(self.driver, 15)

    def prepare_environment(self, url: str, estado: Estado, email: str = ""):
        """
        Navega para a URL e autentica se necessário.
        Pós-login: navega explicitamente para url (SSO redireciona para home).
        """
        self.start_browser()
        self.driver.get(url)
        time.sleep(3)

        url_atual = self.driver.current_url.lower()
        if "login" not in url_atual:
            # Sessão válida — garante que está na URL correta
            if not self._na_url_correta(url):
                estado.log("[→] Sessão válida — redirecionando para URL do PGR...")
                self.driver.get(url)
                time.sleep(3)
            else:
                estado.log("[OK] Sessão ativa — nenhum login necessário.")
            return

        estado.log("[→] Sessão expirada — autenticando...")
        login_ok = False

        if email:
            login_ok = self._tentar_login_auto(email, estado)

        if login_ok:
            estado.log("[→] Navegando para URL do PGR após SSO...")
            self.driver.get(url)
            time.sleep(3)
            estado.log("[OK] Pronto na URL alvo.")
            return

        # Pede login manual via fila (main thread exibe messagebox)
        estado.solicitar_login(Path(self.profile_dir).name)
        self.driver.get(url)
        time.sleep(3)
        estado.log("[OK] Pronto na URL alvo após login manual.")

    def _na_url_correta(self, url_alvo: str) -> bool:
        try:
            atual = self.driver.current_url.split("?")[0].rstrip("/")
            alvo  = url_alvo.split("?")[0].rstrip("/")
            return atual == alvo
        except Exception:
            return False

    def _tentar_login_auto(self, email: str, estado: Estado) -> bool:
        """
        Preenche o campo de usuário e submete o formulário SSO.
        O laço 'for' executa até 2 vezes para lidar com o recarregamento 
        da página que exige a credencial novamente.
        """
        XPATH_CAMPO = "/html/body/div/div/div[3]/div/form/div[1]/input"

        for etapa in range(2):
            estado.log(f"[→] Preenchendo credencial (Passo {etapa + 1}/2)...")
            try:
                # Aguarda o campo estar clicável
                campo = WebDriverWait(self.driver, 12).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO))
                )

                # Garante foco antes de digitar
                self.driver.execute_script("arguments[0].focus();", campo)
                campo.clear()
                campo.send_keys(email)
                time.sleep(0.3)

                # Verifica se o valor foi aceito pelo campo
                valor_atual = campo.get_attribute("value") or ""
                if valor_atual.strip() != email.strip():
                    estado.log("[→] Injetando credencial via JS...")
                    self.driver.execute_script(
                        """
                        var el = arguments[0];
                        var valor = arguments[1];
                        var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                            window.HTMLInputElement.prototype, 'value').set;
                        nativeInputValueSetter.call(el, valor);
                        el.dispatchEvent(new Event('input',  { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        """,
                        campo, email
                    )
                    time.sleep(0.3)

                estado.log("[→] Campo preenchido — submetendo...")

                # Tenta clicar no botão de submit visível no form
                try:
                    btn = self.driver.find_element(
                        By.XPATH,
                        "/html/body/div/div/div[3]/div/form//button"
                    )
                    btn.click()
                except Exception:
                    # Fallback: submit direto no form via JS
                    self.driver.execute_script(
                        "document.querySelector('form').submit();"
                    )

                estado.log("[→] Formulário enviado — aguardando tela...")

                # O pulo do gato: espera a página recarregar verificando se o campo antigo sumiu (Stale)
                try:
                    WebDriverWait(self.driver, 10).until(EC.staleness_of(campo))
                except Exception:
                    pass # Se não detectar a mudança pelo elemento, a verificação da URL logo abaixo resolve

                time.sleep(1.5) # Pausa breve para a nova página/DOM estabilizar

                # Verifica se saiu da tela de login
                if "login" not in self.driver.current_url.lower():
                    estado.log("[OK] SSO autenticado com sucesso.")
                    return True
                
                # Se chegou aqui, a URL ainda tem "login" (a página recarregou pedindo o email de novo)
                if etapa == 0:
                    estado.log("[→] A página recarregou. Reiniciando inserção...")

            except Exception as e:
                estado.log(f"[!] Falha na etapa {etapa + 1} de auto-login ({type(e).__name__}).")
                # Apenas dá refresh forçado se houver um erro técnico/timeout (não se for o fluxo normal)
                if etapa == 0:
                    self.driver.refresh()
                    time.sleep(3)

        estado.log("[!] Auto-login não completou após as tentativas — tentando login manual.")
        return False

    def switch_tab(self, tab_name: str, estado: Estado) -> bool:
        """
        Navega para a aba e aguarda explicitamente as células carregarem
        antes de retornar — elimina os sleeps fixos e reduz tempo morto.
        """
        try:
            estado.log_verbose(f"switch_tab: buscando '{tab_name}'")
            xpath = f"//*[contains(normalize-space(text()), '{tab_name}')]"
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            self.driver.execute_script("arguments[0].click();", el)

            # Aguarda as células da aba aparecerem (explicit wait — sem sleep fixo)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "celula-base"))
            )
            estado.log_verbose(f"switch_tab: aba '{tab_name}' carregada.")
            return True
        except TimeoutException:
            estado.log(f"[!] Aba '{tab_name}' não carregou em 10s.")
            return False
        except Exception as e:
            estado.log(f"[!] switch_tab erro: {type(e).__name__}")
            return False

    def go_back_one_month(self, indicator_name: str) -> bool:
        try:
            xpath = f"//span[normalize-space(.)='{indicator_name}']/following::button[1]"
            btn = self.driver.find_element(By.XPATH, xpath)
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            self.driver.execute_script("arguments[0].click();", btn)
            return True
        except Exception:
            return False

    def fill_field(self, site_name: str, day: str, value: str,
                   estado: Estado) -> bool:
        """
        Preenche uma célula do PGR com retry e backoff exponencial.
        3 tentativas: 0s → 1s → 2s de espera entre falhas.
        send_keys é obrigatório (JS injection não dispara eventos ASP.NET).
        """
        for tentativa in range(3):
            try:
                estado.log_verbose(f"fill: tentativa {tentativa+1} — {site_name} dia {day}")

                xpath_ind  = (f"//span[contains(@class,'status-indicador-text')"
                              f" and normalize-space(.)='{site_name}']")
                xpath_cell = (f"{xpath_ind}/following::div[contains(@class,'celula-base')]"
                              f"[.//div[text()='{day}'] or .//span[text()='{day}']][1]")

                cell = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_cell)))
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cell)
                self.driver.execute_script("arguments[0].click();", cell)

                # Aguarda o input aparecer (explicit wait — sem sleep fixo)
                inp = WebDriverWait(self.driver, 8).until(
                    EC.element_to_be_clickable((By.ID, "valorReal"))
                )
                inp.send_keys(Keys.CONTROL + "a")
                inp.send_keys(Keys.BACKSPACE)
                inp.send_keys(str(value))
                inp.send_keys(Keys.ENTER)

                # Aguarda o input sumir (confirmação de que o valor foi aceito)
                WebDriverWait(self.driver, 8).until(
                    EC.invisibility_of_element_located((By.ID, "valorReal"))
                )
                return True

            except Exception as e:
                estado.log_verbose(f"fill falhou ({type(e).__name__}) t={tentativa+1}")
                if tentativa < 2:
                    time.sleep(2 ** tentativa)  # 1s, 2s

        return False

    def close(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass




def _worker_aba(
    tab_name: str,
    items: dict[str, int],
    value_type: str,
    dados: dict,
    dia_ini: int,
    dia_fim: int,
    is_prev_month: bool,
    profile_dir: str,
    email: str,
    estado: Estado,
    lock: threading.Lock,
) -> list[str]:
    bot    = WebAutomator(profile_dir=profile_dir)
    falhas = []

    try:
        bot.start_browser()
        bot.prepare_environment(SYSTEM_URL, estado, email=email)

        with lock:
            estado.log(f"[→] {value_type.upper()} — navegando para aba...")

        if not bot.switch_tab(tab_name, estado):
            return [f"Aba não encontrada: {tab_name}"]

        if is_prev_month:
            with lock:
                estado.log(f"[→] {value_type.upper()} — ajustando para mês anterior...")
            for nome in items:
                bot.go_back_one_month(nome)
            # Aguarda as células recarregarem após trocar o mês
            try:
                WebDriverWait(bot.driver, 8).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "celula-base"))
                )
            except Exception:
                pass

        zero_ref = "0,00" if value_type == "waste" else "0,000"

        for dia in range(dia_ini, dia_fim + 1):
            for nome, linha in items.items():
                val = dados.get(dia, {}).get(linha, {}).get(value_type)
                if val and val != zero_ref:
                    ok = bot.fill_field(nome, str(dia), val, estado)
                    with lock:
                        if ok:
                            estado.log(f"[OK] {value_type.upper()} dia {dia} | {nome}: {val}")
                        else:
                            falhas.append(f"{value_type.upper()} dia {dia} | {nome}")
                            estado.log(f"[FALHA] {value_type.upper()} dia {dia} | {nome}")
                        estado.avancar(tipo=value_type)

    finally:
        bot.close()

    return falhas




def worker_principal(estado: Estado, cfg: dict, dados: dict):
    try:
        waste_items:   dict[str, int] = {}
        consumo_items: dict[str, int] = {}

        for cat in cfg["categorias"]:
            for nome, (linha, tipo_item) in INDICATORS_MAP[cat]["items"].items():
                if tipo_item == "waste":
                    waste_items[nome]   = linha
                elif tipo_item == "consumo":
                    consumo_items[nome] = linha

        abas: list[tuple] = []
        if cfg["tipo"] in ("waste", "all"):
            abas.append(("Lançamentos de Controle | Waste",
                         waste_items, "waste", PROFILE_DIR_WASTE))
        if cfg["tipo"] in ("consumo", "all"):
            abas.append(("Lançamentos de Controle | Consumo Específico",
                         consumo_items, "consumo", PROFILE_DIR_CONSUMO))

        dias = cfg["dia_fim"] - cfg["dia_ini"] + 1

        estado.waste_total   = len(waste_items) * dias   if cfg["tipo"] in ("waste", "all")   else 0
        estado.consumo_total = len(consumo_items) * dias if cfg["tipo"] in ("consumo", "all") else 0
        estado.progresso_total = sum(len(items) * dias for _, items, _, _ in abas)

        lock         = threading.Lock()
        max_workers  = min(cfg["n_drivers"], len(abas))
        todas_falhas = []

        estado.verbose = cfg.get("verbose", False)
        estado.log(f"[→] Iniciando preenchimento — {max_workers} driver(s)...")
        estado.status = f"Preenchendo PGR — {max_workers} driver(s)..."

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(
                    _worker_aba,
                    tab_name, items, value_type, dados,
                    cfg["dia_ini"], cfg["dia_fim"],
                    cfg["is_prev_month"],
                    profile_dir,
                    cfg.get("email", ""),
                    estado, lock,
                ): tab_name
                for tab_name, items, value_type, profile_dir in abas
            }
            for fut in as_completed(futures):
                try:
                    todas_falhas.extend(fut.result())
                except Exception as e:
                    todas_falhas.append(str(e))

        if todas_falhas:
            estado.log(f"[!] {len(todas_falhas)} falha(s):")
            for f in todas_falhas:
                estado.log(f"  [FALHA] {f}")
            estado.status = f"Concluído com {len(todas_falhas)} falha(s)"
            notificar_windows("PGR Orchestrator",
                              f"{len(todas_falhas)} falha(s). Verifique o log.", icone="Warning")
        else:
            estado.log("[OK] Todos os lançamentos concluídos sem falhas.")
            estado.status = f"✔  Concluído — {cfg['periodo_label']}"
            notificar_windows("PGR Orchestrator", f"Período {cfg['periodo_label']} lançado com sucesso!")

    except Exception as e:
        estado.log(f"[ERRO] {type(e).__name__}: {e}")
        estado.erro   = str(e)
        estado.status = "✖  Erro — verifique o log"
        notificar_windows("PGR Orchestrator", f"Erro: {e}", icone="Error")
    finally:
        estado.concluido = True




def main():
    # Suprime warnings de inicialização do Edge/Selenium que vazam para stderr
    if sys.stderr:
        try:
            sys.stderr = open(os.devnull, "w")
        except Exception:
            pass

    config = carregar_config()

    while True:
        gui = ConfiguracaoGUI(config)
        if not gui.resultado:
            break

        cfg = gui.resultado

        config.update({
            "excel_path": cfg["excel_path"],
            "email":      cfg.get("email", ""),
            "periodo":    "anterior" if cfg["is_prev_month"] else "atual",
            "tipo":       cfg["tipo"],
            "n_drivers":  cfg["n_drivers"],
            "verbose":    cfg.get("verbose", False),
        })
        salvar_config(config)

        # Lê Excel na main thread (antes do worker — evita COM em thread errada)
        required_rows = {
            row
            for cat in cfg["categorias"]
            for _, (row, _) in INDICATORS_MAP[cat]["items"].items()
        }
        engine = ExcelEngine(cfg["excel_path"])
        try:
            engine.abrir()
            dados = engine.extrair_dados(
                cfg["dia_ini"], cfg["dia_fim"],
                required_rows, cfg["mes"], cfg["ano"],
            )
        finally:
            engine.fechar()

        zeros = {"waste": "0,00", "consumo": "0,000"}
        tipos_busca = ["waste", "consumo"] if cfg["tipo"] == "all" else [cfg["tipo"]]
        validos = sum(
            1
            for dia_data in dados.values()
            for vals in dia_data.values()
            for t in tipos_busca
            if vals.get(t) and vals[t] != zeros[t]
        )
        if validos == 0:
            messagebox.showwarning(
                "Sem dados",
                "Nenhum valor diferente de zero encontrado na planilha.\n"
                "Verifique o intervalo de datas e o arquivo selecionado."
            )
            continue

        dual = cfg["n_drivers"] == 2 and cfg["tipo"] == "all"
        estado = Estado()

        modo_txt = "paralelo" if dual else "sequencial"
        janela = JanelaProgresso(
            titulo=f"PGR Orchestrator v{VERSAO}",
            subtitulo=f"{cfg['periodo_label']}  •  {Path(cfg['excel_path']).name}  •  {modo_txt}",
            estado=estado,
            dual=dual,
        )
        janela.iniciar(worker_principal, estado, cfg, dados)

        if estado.erro:
            if not messagebox.askyesno("PGR Orchestrator",
                                       f"Erro: {estado.erro}\n\nDeseja configurar novamente?"):
                break
        else:
            if not messagebox.askyesno("PGR Orchestrator", "Concluído!\nDeseja configurar nova execução?"):
                break


if __name__ == "__main__":
    main()