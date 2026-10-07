"""
SAP Bridge v4.0.0
=================
Extrai ZCO1, MB51 e PRODUÇÃO do SAP e consolida em planilha mestre.

Novidades v4:
- Multithreading configurável (1 a 4 janelas SAP em paralelo) para ZCO1
- sap_lock garante que apenas uma thread por vez interage com os menus SAP
- Cada thread usa seu próprio lista_ordens_N.txt (evita file lock do Windows)
- sbar check: dias sem dados são pulados automaticamente sem quebrar a extração
"""

import os
import glob
import json
import shutil
import time
import sys
import threading
import math
import queue
import subprocess
import pythoncom
import win32com.client
import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from tkcalendar import DateEntry

VERSAO        = "4.0.0"
CONFIG_FILE   = Path.home() / ".sap_bridge_config.json"
SAP_EXE_PATH  = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
AMBIENTE_SAP  = "SAP_CONNECTION_NAME"
EMPRESA_SAP   = "COMP"
PLANTA_SAP    = "1000"

DEPOSITOS_MB51  = ["LI05","LI04","LI06","LI07","LI08","LI12","LI13","LI14","LI15",
                   "LP01","LP02","LF01","LF02","LF03","LF04","LF05","PR01","HD01","MP05","MP06","REB7", "MP07", "MP09", "SA01"]
MOVIMENTOS_MB51 = ["531","532","261","262"]
MOVIMENTOS_PROD = ["101","102"]

COLORS = {
    "bg":      "#0d1117",
    "bg2":     "#161b22",
    "border":  "#30363d",
    "text":    "#8b949e",
    "subtext": "#8b949e",
    "blue":    "#58a6ff",
    "green":   "#3fb950",
    "yellow":  "#d29922",
    "red":     "#f85149",
    "bar":     "#1f6feb",
    "trough":  "#21262d",
}


@dataclass
class Estado:
    progresso_atual: int = 0
    progresso_total: int = 100
    status: str = "Aguardando..."
    concluido: bool = False
    erro: str | None = None
    log_queue: queue.Queue = field(default_factory=queue.Queue)

    def log(self, msg: str):
        self.log_queue.put(msg)

    def avancar(self, status: str = "", incremento: int = 1):
        if status:
            self.status = status
        self.progresso_atual = min(self.progresso_atual + incremento, self.progresso_total)




class JanelaProgresso:
    POLL_MS = 200

    def __init__(self, titulo: str, subtitulo: str, estado: Estado):
        self._estado = estado
        self._t0     = time.time()
        C = COLORS

        self.root = tk.Tk()
        self.root.title(f"SAP Bridge v{VERSAO}")
        self.root.geometry("700x480")
        self.root.resizable(False, False)
        self.root.configure(bg=C["bg"])
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        hdr = tk.Frame(self.root, bg=C["bg"], pady=12)
        hdr.pack(fill="x", padx=20)
        tk.Label(hdr, text="🗄  SAP Bridge", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI", 14, "bold"), anchor="w").pack(side="left")
        tk.Label(hdr, text=f"v{VERSAO}", bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="right", padx=2)
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
        style.configure("SAP.Horizontal.TProgressbar",
                        troughcolor=C["trough"], background=C["bar"],
                        bordercolor=C["border"], lightcolor=C["bar"],
                        darkcolor=C["bar"], thickness=14)
        self._bar = ttk.Progressbar(body, style="SAP.Horizontal.TProgressbar",
                                    length=660, mode="determinate")
        self._bar.pack(fill="x", pady=(0, 4))

        metrics = tk.Frame(body, bg=C["bg"])
        metrics.pack(fill="x", pady=(0, 8))
        self._lbl_cnt  = tk.Label(metrics, text="0 / 0",   bg=C["bg"], fg=C["subtext"], font=("Segoe UI", 8))
        self._lbl_time = tk.Label(metrics, text="00:00:00", bg=C["bg"], fg=C["subtext"], font=("Segoe UI", 8))
        self._lbl_cnt.pack(side="left")
        self._lbl_time.pack(side="right")

        tk.Frame(body, bg=C["border"], height=1).pack(fill="x", pady=(0, 6))

        self._log = tk.Text(body, height=13, state="disabled",
                            bg=C["bg2"], fg="#c9d1d9",
                            font=("Consolas", 8), relief="flat",
                            insertbackground=C["text"], padx=6, pady=4,
                            selectbackground="#264f78")
        self._log.pack(fill="both", expand=True)
        self._log.tag_configure("ok",    foreground=C["green"])
        self._log.tag_configure("warn",  foreground=C["yellow"])
        self._log.tag_configure("error", foreground=C["red"])
        self._log.tag_configure("info",  foreground=C["blue"])
        self._log.tag_configure("dim",   foreground=C["subtext"])

    def _append_log(self, msg: str):
        self._log.config(state="normal")
        ts  = datetime.now().strftime("%H:%M:%S")
        tag = ("ok"    if msg.startswith("[OK]")
               else "error" if any(msg.startswith(p) for p in ("[FALHA]", "[ERRO]"))
               else "warn"  if msg.startswith("[!]")
               else "info"  if msg.startswith("[→]")
               else "dim")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _poll(self):
        e = self._estado
        if e.progresso_total > 0:
            self._bar["value"] = (e.progresso_atual / e.progresso_total) * 100
        self._lbl_status.config(text=e.status)
        self._lbl_cnt.config(text=f"{e.progresso_atual} / {e.progresso_total}")
        elapsed = int(time.time() - self._t0)
        h, rem = divmod(elapsed, 3600)
        m, s   = divmod(rem, 60)
        self._lbl_time.config(text=f"{h:02d}:{m:02d}:{s:02d}")
        try:
            while True:
                self._append_log(e.log_queue.get_nowait())
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
        e = ttk.Entry(parent, textvariable=textvariable, width=width,
                      style="Dark.TEntry", state="readonly" if readonly else "normal")
        return e

    def _btn(self, parent, text, command, primary=False, width=12) -> tk.Button:
        C = COLORS
        bg  = C["blue"] if primary else C["bg2"]
        fg  = C["bg"]   if primary else C["text"]
        abg = "#1158c7"  if primary else C["border"]
        return tk.Button(parent, text=text, command=command,
                         bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                         font=("Segoe UI", 9, "bold" if primary else "normal"),
                         relief="flat", bd=0, padx=10, pady=5,
                         cursor="hand2", width=width)

    def _build(self):
        C = COLORS
        self.root = tk.Tk()
        self.root.title(f"SAP Bridge v{VERSAO}")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        hoje = datetime.now()

        hdr = tk.Frame(self.root, bg=C["bg"], pady=14)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="🗄  SAP Bridge", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI", 14, "bold"), anchor="w").pack(side="left")
        tk.Label(hdr, text=f"v{VERSAO}", bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="right", padx=4)
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")

        c1 = self._card(self.root, "📁  Pasta de Destino SAP")
        row1 = tk.Frame(c1, bg=C["bg2"])
        row1.pack(fill="x")
        self._pasta = tk.StringVar(value=self._cfg.get("pasta_sap", ""))
        self._entry_dark(row1, self._pasta, width=52, readonly=True).pack(side="left", padx=(0, 8))
        self._btn(row1, "Buscar…", self._sel_pasta).pack(side="left")

        c2 = self._card(self.root, "📊  Planilha Mestre ZC")
        row2 = tk.Frame(c2, bg=C["bg2"])
        row2.pack(fill="x")
        self._mestre = tk.StringVar(value=self._cfg.get("mestre_path", ""))
        self._entry_dark(row2, self._mestre, width=52, readonly=True).pack(side="left", padx=(0, 8))
        self._btn(row2, "Buscar…", self._sel_mestre).pack(side="left")

        c3 = self._card(self.root, "📅  Período de Extração")
        row3 = tk.Frame(c3, bg=C["bg2"])
        row3.pack(fill="x")
        de_opts = dict(width=12, background="#1f6feb", foreground="white",
                       borderwidth=0, date_pattern="dd/MM/yyyy", locale="pt_BR",
                       font=("Segoe UI", 9))
        tk.Label(row3, text="De:", bg=C["bg2"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self._de_ini = DateEntry(row3, year=hoje.year, month=hoje.month, day=1, **de_opts)
        self._de_ini.pack(side="left", padx=(0, 20))
        tk.Label(row3, text="Até:", bg=C["bg2"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self._de_fim = DateEntry(row3, year=hoje.year, month=hoje.month, day=hoje.day, **de_opts)
        self._de_fim.pack(side="left")

        hint = tk.Frame(c3, bg=C["bg2"])
        hint.pack(fill="x", pady=(4, 0))
        tk.Label(hint, text="ℹ  Convertido automaticamente para dd.mm.aaaa (formato SAP).",
                 bg=C["bg2"], fg=C["subtext"], font=("Segoe UI", 8)).pack(anchor="w")

        c4 = self._card(self.root, "⚙  Extrações Adicionais e Threads")
        
        row_perf = tk.Frame(c4, bg=C["bg2"])
        row_perf.pack(fill="x", pady=(0, 6))
        tk.Label(row_perf, text="Janelas SAP (Threads ZCO1):", bg=C["bg2"], fg=C["text"],
                 font=("Segoe UI", 9)).pack(side="left")
        self._threads = tk.IntVar(value=self._cfg.get("num_threads", 1))
        cb_threads = ttk.Combobox(row_perf, textvariable=self._threads, values=[1, 2, 3, 4],
                                  state="readonly", width=3)
        cb_threads.pack(side="left", padx=8)
        
        self._mb51 = tk.BooleanVar(value=True)
        self._prod = tk.BooleanVar(value=True)
        for var, lbl in [(self._mb51, "Extrair MB51  (movimentos de estoque)"),
                          (self._prod, "Extrair PRODUÇÃO  (movimentos 101/102)")]:
            tk.Checkbutton(c4, text=lbl, variable=var,
                           bg=C["bg2"], fg=C["text"], selectcolor=C["bg"],
                           activebackground=C["bg2"], activeforeground=C["blue"],
                           font=("Segoe UI", 9), bd=0).pack(anchor="w", pady=1)

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", pady=(8, 0))
        footer = tk.Frame(self.root, bg=C["bg"], pady=12)
        footer.pack(fill="x", padx=16)
        self._btn(footer, "Cancelar", self._cancelar, width=12).pack(side="right", padx=(8, 0))
        self._btn(footer, "▶  Executar", self._confirmar, primary=True, width=14).pack(side="right")

        self.root.mainloop()

    def _sel_pasta(self):
        p = filedialog.askdirectory(title="Pasta raiz SAP", parent=self.root)
        if p:
            self._pasta.set(p)

    def _sel_mestre(self):
        p = filedialog.askopenfilename(title="Planilha mestre ZC",
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
                                       parent=self.root)
        if p:
            self._mestre.set(p)

    def _cancelar(self):
        self.root.destroy()

    def _confirmar(self):
        if not self._pasta.get():
            messagebox.showerror("Erro", "Selecione a pasta de destino SAP.", parent=self.root)
            return
        if not self._mestre.get():
            messagebox.showerror("Erro", "Selecione a planilha mestre.", parent=self.root)
            return
        dt_ini = self._de_ini.get_date()
        dt_fim = self._de_fim.get_date()
        if dt_ini > dt_fim:
            messagebox.showerror("Erro", "Data inicial não pode ser posterior à data final.", parent=self.root)
            return
        self.resultado = {
            "pasta_sap":    self._pasta.get(),
            "mestre_path":  self._mestre.get(),
            "data_ini_sap": dt_ini.strftime("%d.%m.%Y"),
            "data_fim_sap": dt_fim.strftime("%d.%m.%Y"),
            "puxar_mb51":   self._mb51.get(),
            "puxar_prod":   self._prod.get(),
            "dt_ini":       dt_ini,
            "dt_fim":       dt_fim,
            "periodo_label": f"{dt_ini.strftime('%d/%m/%Y')} → {dt_fim.strftime('%d/%m/%Y')}",
            "num_threads": self._threads.get(),
        }
        self.root.destroy()




def preparar_pasta_temp(pasta_sap: str) -> str:
    dest = os.path.join(pasta_sap, "Temp ZC")
    Path(dest).mkdir(parents=True, exist_ok=True)
    for arq in glob.glob(os.path.join(dest, "*")):
        try:
            os.remove(arq)
        except Exception:
            pass
    return dest




def _aguardar_workbook_aberto(nome_arquivo: str, pasta_destino: str,
                              estado: Estado, timeout: int = 20) -> object | None:
    """
    Polling ativo até o Workbook estar acessível via COM.
    Corrige o bug de timing: session.press() não bloqueia, Excel ainda está
    carregando quando o código tenta fechar.
    """
    nome_base    = Path(nome_arquivo).name.lower()
    caminho_full = Path(pasta_destino) / nome_arquivo

    for _ in range(timeout * 2):
        if caminho_full.exists():
            break
        time.sleep(0.5)
    else:
        estado.log(f"[!] {nome_arquivo} nunca apareceu no disco.")
        return None

    for _ in range(timeout * 2):
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            for i in range(1, xl.Workbooks.Count + 1):
                try:
                    wb = xl.Workbooks(i)
                    if wb.Name.lower() == nome_base:
                        _ = wb.Sheets.Count  # Confirma que o COM está totalmente inicializado
                        return wb
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.5)

    estado.log(f"[!] {nome_arquivo} não ficou acessível via COM em {timeout}s.")
    return None


def fechar_arquivo_sap(pasta_destino: str, nome_arquivo: str, estado: Estado):
    wb = _aguardar_workbook_aberto(nome_arquivo, pasta_destino, estado)
    if wb is None:
        estado.log(f"[!] {nome_arquivo} não pôde ser fechado.")
        return
    try:
        app = wb.Application
        restantes = app.Workbooks.Count
        wb.Close(SaveChanges=False)
        estado.log(f"[OK] Fechado: {nome_arquivo}")
        if restantes <= 1:
            try:
                app.Quit()
                estado.log(f"[OK] Instância Excel exclusiva finalizada.")
            except Exception:
                pass
    except Exception as e:
        estado.log(f"[!] Erro ao fechar {nome_arquivo}: {type(e).__name__}")


def fechar_remanescentes_sap(pasta_destino: str, padrao: str, estado: Estado):
    """Rede de segurança: varre ROT e fecha o que escapou do fechamento individual."""
    import fnmatch
    fechados = 0
    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        vistos = set()
        for moniker in rot:
            try:
                display  = moniker.GetDisplayName(ctx, None)
                nome_arq = Path(display).name
                if not fnmatch.fnmatch(nome_arq.lower(), padrao.lower()):
                    continue
                if nome_arq in vistos:
                    continue
                vistos.add(nome_arq)
                obj = rot.GetObject(moniker)
                wb  = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                app = wb.Application
                restantes = app.Workbooks.Count
                wb.Close(SaveChanges=False)
                fechados += 1
                estado.log(f"[OK] Remanescente fechado: {nome_arq}")
                if restantes <= 1:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            except Exception:
                continue
    except Exception as e:
        estado.log(f"[!] Varredura remanescentes: {type(e).__name__}")
        return
    if fechados == 0:
        estado.log("[OK] Nenhum remanescente — todos fechados corretamente.")
    else:
        estado.log(f"[OK] {fechados} remanescente(s) fechado(s).")




def conectar_sap(estado: Estado) -> object:
    """
    Tradução fiel da macro VBA fornecida:

    1. Tenta GetObject("SAPGUI") em loop com 1s de espera
       — se o SAP já estiver aberto, não reabre
    2. Se já há conexão ativa (Children.Count > 0), reutiliza
       — elimina a abertura de aba duplicada
    3. Caso contrário, abre nova conexão com OpenConnection
    4. Retorna a sessão (Children(0)) da conexão

    O sleep fixo de 7s foi eliminado — o polling para assim que o
    scripting engine responder, acelerando o startup em máquinas rápidas
    e sendo mais robusto em máquinas lentas.
    """
    estado.log("[→] Verificando SAP GUI...")

    # Tenta conectar ao SAPGUI já aberto
    sap_gui = None
    for _ in range(30):  # Até 30s aguardando o SAP aparecer
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            pass
        time.sleep(1)

    if sap_gui is None:
        # SAP não estava aberto — inicia e aguarda
        estado.log("[→] Iniciando SAP GUI...")
        subprocess.Popen(SAP_EXE_PATH)
        for _ in range(30):
            try:
                sap_gui = win32com.client.GetObject("SAPGUI")
                break
            except Exception:
                pass
            time.sleep(1)

    if sap_gui is None:
        raise RuntimeError("SAP GUI não ficou disponível em 30s.")

    app = sap_gui.GetScriptingEngine

    # Reutiliza conexão existente se o SAP já tiver uma aberta
    # (lógica equivalente ao "If SAPApp.Children.Count = 0" do VBA)
    if app.Children.Count > 0:
        conn = app.Children(0)
        estado.log("[OK] Conexão SAP existente reutilizada — sem abrir aba dupla.")
    else:
        estado.log("[→] Abrindo conexão SAP...")
        conn = app.OpenConnection(AMBIENTE_SAP, True)
        time.sleep(2)

    return conn.Children(0)




def fechar_sap(session, estado: Estado):
    """
    Fecha a sessão SAP via scripting — tradução do trecho VBA fornecido:
      session.findById("wnd[0]").Close
      session.findById("wnd[1]/usr/btnSPOP-OPTION1").Press

    O botão SPOP-OPTION1 confirma "Deseja encerrar a sessão?" sem interação manual.
    Erros são silenciados (On Error Resume Next do VBA) — se o SAP já tiver
    fechado por outro motivo, não quebra o fluxo.
    """
    try:
        estado.log("[→] Fechando sessão SAP...")
        session.findById("wnd[0]").Close()
        time.sleep(1)
        try:
            # Popup de confirmação de saída
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").Press()
        except Exception:
            pass  # Popup pode não aparecer se a sessão fechar diretamente
        estado.log("[OK] SAP fechado.")
    except Exception as e:
        estado.log(f"[!] Aviso ao fechar SAP: {type(e).__name__} — {e}")




def extrair_ordens(caminho_mestre: str, pasta_destino: str) -> str:
    df = pd.read_excel(caminho_mestre, sheet_name="BASE ORDENS",
                       usecols=[0], skiprows=1, header=None, dtype=str)
    ordens = []
    for v in df[0].tolist():
        s = str(v).strip()
        if s in ("", "nan", "none", "None"):
            break
        ordens.append(s)
    if not ordens:
        raise ValueError("Nenhuma ordem encontrada em 'BASE ORDENS'.")
    arq = os.path.join(pasta_destino, "lista_ordens.txt")
    Path(arq).write_text("\n".join(ordens), encoding="utf-8")
    return arq



sap_lock = threading.Lock()



def extrair_zco1(session, datas: list[str], pasta_destino: str, estado: Estado, arq_ordens: str = "lista_ordens.txt") -> list[str]:
    gerados: list[str] = []
    
    for idx, data_str in enumerate(datas, 1):
        nome = f"ZCO1_{data_str.replace('.', '_')}.xlsx"
        estado.avancar(status=f"ZCO1 — {data_str}  ({idx}/{len(datas)})")
        estado.log(f"[→] Extraindo ZCO1 {data_str}...")

        try:
            with sap_lock:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nzco1"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/ctxtP_CODEMP").text    = EMPRESA_SAP
                session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = PLANTA_SAP
                session.findById("wnd[0]/usr/ctxtS_MES-LOW").text   = data_str
    
                _importar_txt_sap(session, pasta_destino, arq_ordens,
                                 "wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH")

            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)

            try:
                sbar = session.findById("wnd[0]/sbar")
                if sbar and sbar.text:
                    txt = sbar.text
                    if (sbar.messageType in ("E", "A") or
                            "Não há registros" in txt or
                            "No records" in txt or
                            "sem dados" in txt.lower()):
                        estado.log(f"[!] ZCO1 {data_str}: Sem dados — {txt}")
                        continue
            except Exception:
                pass

            with sap_lock:
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/usr/cntlC_ALV/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton("&MB_VARIANT")
                session.findById("wnd[0]/usr/cntlC_ALV/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&LOAD")
                
                # Janela de seleção de layout
                shell_layout = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
                shell_layout.contextMenu()
                shell_layout.selectContextMenuItem("&FIND")
                
                # Janela de busca (/USER_LAYOUT)
                session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
                session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/USER_LAYOUT"
                session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
                session.findById("wnd[2]/tbar[0]/btn[0]").press()
                
                # Fecha a janela de busca e aplica o layout selecionado na grade do S/4HANA
                session.findById("wnd[2]").close()
                shell_layout.doubleClickCurrentCell()  # Garante o carregamento definitivo do layout na grade
                time.sleep(2.0)
    
                # Exportação para Excel
                session.findById("wnd[0]/usr/cntlC_ALV/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton("&MB_EXPORT")
                session.findById("wnd[0]/usr/cntlC_ALV/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
                
                session.findById("wnd[1]/tbar[0]/btn[20]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta_destino
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
                session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Confirmação final do salvamento

            gerados.append(nome)
            estado.log(f"[OK] {nome}")
            fechar_arquivo_sap(pasta_destino, nome, estado)

        except Exception as e:
            estado.log(f"[FALHA] ZCO1 {data_str}: {e}")

    return gerados




def extrair_mb51(session, d_ini: str, pasta_destino: str, estado: Estado) -> bool:
    estado.avancar(status="Extraindo MB51...")
    estado.log("[→] MB51...")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = PLANTA_SAP
        
        _importar_txt_sap(session, pasta_destino, "depositos_mb51.txt",
                          "wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH")
        _importar_txt_sap(session, pasta_destino, "movimentos_mb51.txt",
                          "wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH")
        
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text  = d_ini
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = datetime.now().strftime("%d.%m.%Y")
        
        # Seleção de Lista Plana (HANA)
        try:
            session.findById("wnd[0]/usr/radRFLAT_L").select()
        except Exception:
            pass

        session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/USER_LAYOUT"
        
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(16)
        
        # Fluxo de Exportação HANA
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text     = pasta_destino
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        fechar_arquivo_sap(pasta_destino, "MB51.xlsx", estado)
        estado.log("[OK] MB51.xlsx gerado.")
        return True
    except Exception as e:
        estado.log(f"[FALHA] MB51: {e}")
        return False


def extrair_producao(session, d_ini: str, pasta_destino: str, estado: Estado) -> bool:
    estado.avancar(status="Extraindo PRODUÇÃO...")
    estado.log("[→] PRODUÇÃO...")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = PLANTA_SAP
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "TS00"
        
        _importar_txt_sap(session, pasta_destino, "movimentos_prod.txt",
                          "wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH")
        
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text  = d_ini
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = datetime.now().strftime("%d.%m.%Y")
        
        # Seleção de Lista Plana (HANA)
        try:
            session.findById("wnd[0]/usr/radRFLAT_L").select()
        except Exception:
            pass

        session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/PRODMDC"
        
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(16)
        
        # Fluxo de Exportação HANA
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text     = pasta_destino
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PROD.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        fechar_arquivo_sap(pasta_destino, "PROD.xlsx", estado)
        estado.log("[OK] PROD.xlsx gerado.")
        return True
    except Exception as e:
        estado.log(f"[FALHA] PRODUÇÃO: {e}")
        return False


def _importar_txt_sap(session, pasta: str, nome_txt: str, btn_id: str):
    session.findById(btn_id).press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[23]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text     = pasta
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = nome_txt
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()




def tratar_valor(valor):
    if pd.isna(valor) or valor is None:
        return ""
    s = str(valor).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    try:
        n = float(s)
        return int(n) if n.is_integer() else n
    except ValueError:
        return s


def ler_excel_blindado(caminho: str) -> str:
    try:
        os.rename(caminho, caminho)
        return caminho
    except OSError:
        copia = os.path.join(os.path.dirname(caminho), f"_tmp_{os.path.basename(caminho)}")
        shutil.copy2(caminho, copia)
        return copia




class ExcelMestre:
    def __init__(self, caminho: str):
        self.caminho = caminho
        self.excel   = None
        self.wb      = None

    def conectar(self):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.DisplayAlerts = False
        nome = os.path.basename(self.caminho).lower()
        for i in range(1, self.excel.Workbooks.Count + 1):
            if self.excel.Workbooks(i).Name.lower() == nome:
                self.wb = self.excel.Workbooks(i)
                return
        self.wb = self.excel.Workbooks.Open(self.caminho)

    def injetar_dados(self, nome_aba: str, dados: list[list],
                      n_colunas: int, formatar_data_col: int | None = None):
        abas = [sh.Name for sh in self.wb.Sheets]
        if nome_aba not in abas:
            self.wb.Sheets.Add().Name = nome_aba
        ws = self.wb.Sheets(nome_aba)
        self.excel.Calculation = -4135  # xlManual
        try:
            ultima = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
            if ultima >= 2:
                ws.Range(ws.Cells(2, 1), ws.Cells(ultima, n_colunas)).ClearContents()
            n = len(dados)
            if n == 0:
                return
            ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, n_colunas)).Value = dados
            if formatar_data_col:
                col = chr(64 + formatar_data_col)
                ws.Range(f"{col}2:{col}{n + 1}").NumberFormatLocal = "dd/mm/aaaa"
        finally:
            self._restaurar_calculo()

    def _restaurar_calculo(self):
        for _ in range(5):
            try:
                self.excel.Calculation = -4105
                return
            except Exception:
                time.sleep(0.6)

    def salvar_e_exibir(self):
        try:
            self.wb.Calculate()
        except Exception:
            pass
        for tentativa in range(3):
            try:
                self.wb.Save()
                break
            except Exception:
                if tentativa == 2:
                    raise
                time.sleep(1)
        self.excel.Visible = True




def consolidar_zco1(pasta_destino: str, estado: Estado) -> list[list]:
    """
    Lê os ZCO1_*.xlsx e atualiza o progresso por arquivo lido.
    O status muda a cada arquivo — usuário vê "Consolidando 1/29", "2/29"...

    IMPORTANTE: usa header=None + skiprows=1 para ignorar completamente o
    cabeçalho — apenas a POSIÇÃO da coluna importa. Isso evita que qualquer
    edição acidental no cabeçalho pelo usuário quebre o processamento.
    """
    arquivos = sorted(glob.glob(os.path.join(pasta_destino, "ZCO1_*.xlsx")))
    total    = len(arquivos)
    if total == 0:
        estado.log("[!] Nenhum arquivo ZCO1 encontrado.")
        return []

    # Reserva slots de progresso para a consolidação
    estado.progresso_total += total
    dfs = []
    for idx, arq in enumerate(arquivos, 1):
        estado.avancar(status=f"Consolidando ZCO1... ({idx}/{total})")
        try:
            dfs.append(pd.read_excel(ler_excel_blindado(arq), dtype=str,
                                     header=None, skiprows=1).iloc[:, :18])
        except Exception as e:
            estado.log(f"[!] Ignorando {Path(arq).name}: {e}")

    if not dfs:
        return []

    df = pd.concat(dfs, ignore_index=True)

    # Conversão da coluna de data (posição 7, índice 0-based) para serial Excel.
    # Com header=None as colunas se chamam 0,1,2...17 — acesso por label df[7]
    # permite coerção de tipo (int64 dentro de DataFrame str). iloc recusa.
    if df.shape[1] >= 8:
        serie = pd.to_datetime(
            df[7].astype(str).str.replace(".", "/", regex=False),
            dayfirst=True, errors="coerce"
        )
        base = pd.Timestamp("1899-12-30")
        df[7] = (serie - base).dt.days
        df[7] = df[7].where(df[7].notna(), None)

    # tratar_valor em todas as colunas, exceto a de data (posição 7)
    for i in range(df.shape[1]):
        if df.shape[1] >= 8 and i == 7:
            continue
        df[i] = df[i].map(tratar_valor)

    return df.values.tolist()


def consolidar_extra(pasta_destino: str, nome: str, n_colunas: int, estado: Estado) -> list[list]:
    """
    Lê MB51.xlsx / PROD.xlsx ignorando o cabeçalho (header=None + skiprows=1).
    Apenas a posição da coluna importa — o texto do cabeçalho é irrelevante.
    """
    caminho = os.path.join(pasta_destino, nome)
    if not Path(caminho).exists():
        estado.log(f"[!] {nome} não encontrado.")
        return []
    df = pd.read_excel(ler_excel_blindado(caminho), dtype=str,
                       header=None, skiprows=1).iloc[:, :n_colunas]
    df = df.map(tratar_valor)
    return df.values.tolist()



def worker_zco1_paralelo(sessao_idx: int, datas_chunk: list[str], pasta_destino: str, estado: Estado):
    if not datas_chunk:
        return
    time.sleep(sessao_idx * 2.5)
    pythoncom.CoInitialize()
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)
        session_local = conn.Children(sessao_idx)
        arq_ordens = f"lista_ordens_{sessao_idx}.txt"
        extrair_zco1(session_local, datas_chunk, pasta_destino, estado, arq_ordens)
    except Exception as e:
        estado.log(f"[ERRO] Falha na Thread ZCO1 (Sessão {sessao_idx}): {e}")
    finally:
        pythoncom.CoUninitialize()



def worker_principal(estado: Estado, cfg: dict):
    pythoncom.CoInitialize()
    session = None
    try:
        pasta_destino = preparar_pasta_temp(cfg["pasta_sap"])
        estado.log(f"[→] Pasta pronta.")

        extrair_ordens(cfg["mestre_path"], pasta_destino)
        estado.log(f"[OK] Ordens extraídas.")

        if cfg["puxar_mb51"]:
            Path(os.path.join(pasta_destino, "depositos_mb51.txt")).write_text(
                "\n".join(DEPOSITOS_MB51), encoding="utf-8")
            Path(os.path.join(pasta_destino, "movimentos_mb51.txt")).write_text(
                "\n".join(MOVIMENTOS_MB51), encoding="utf-8")
        if cfg["puxar_prod"]:
            Path(os.path.join(pasta_destino, "movimentos_prod.txt")).write_text(
                "\n".join(MOVIMENTOS_PROD), encoding="utf-8")

        dt_ini, dt_fim = cfg["dt_ini"], cfg["dt_fim"]
        datas = [(dt_ini + timedelta(days=x)).strftime("%d.%m.%Y")
                 for x in range((dt_fim - dt_ini).days + 1)]

        estado.status = "Conectando ao SAP..."
        session = conectar_sap(estado)
        estado.log("[OK] SAP conectado.")

        num_threads = min(cfg.get("num_threads", 1), max(1, len(datas)))

        if num_threads == 1:
            estado.progresso_total += len(datas)
            extrair_zco1(session, datas, pasta_destino, estado, "lista_ordens.txt")
        else:
            estado.progresso_total += len(datas)
            estado.log(f"[→] Abrindo {num_threads - 1} nova(s) janela(s) do SAP...")
            session.findById("wnd[0]").maximize()
            conn = session.Parent

            for i in range(1, num_threads):
                session.createSession()
                for _ in range(60):
                    if conn.Children.Count > i:
                        break
                    time.sleep(0.5)
                time.sleep(2)

            tamanho_chunk = math.ceil(len(datas) / num_threads)
            chunks = [datas[i:i + tamanho_chunk] for i in range(0, len(datas), tamanho_chunk)]

            estado.log(f"[→] Dividindo ZCO1 em {len(chunks)} pacote(s)...")

            arq_original = os.path.join(pasta_destino, "lista_ordens.txt")
            for i in range(num_threads):
                shutil.copy2(arq_original, os.path.join(pasta_destino, f"lista_ordens_{i}.txt"))

            threads = []
            for i, chunk in enumerate(chunks):
                t = threading.Thread(
                    target=worker_zco1_paralelo,
                    args=(i, chunk, pasta_destino, estado),
                    daemon=True
                )
                threads.append(t)
                t.start()
                time.sleep(0.5)

            for t in threads:
                t.join()

        estado.log("[→] Verificando remanescentes...")
        fechar_remanescentes_sap(pasta_destino, "ZCO1_*.xlsx", estado)

        estado.progresso_total += (2 if cfg["puxar_mb51"] and cfg["puxar_prod"]
                                   else 1 if cfg["puxar_mb51"] or cfg["puxar_prod"]
                                   else 0)

        if cfg["puxar_mb51"]:
            extrair_mb51(session, cfg["data_ini_sap"], pasta_destino, estado)
        if cfg["puxar_prod"]:
            extrair_producao(session, cfg["data_ini_sap"], pasta_destino, estado)

        if cfg["puxar_mb51"]:
            fechar_remanescentes_sap(pasta_destino, "MB51.xlsx", estado)
        if cfg["puxar_prod"]:
            fechar_remanescentes_sap(pasta_destino, "PROD.xlsx", estado)

        # Fecha todas as sessões SAP abertas pela extração
        if session:
            try:
                conn = session.Parent
                # Fecha sessões extras de trás pra frente (evita reindexação)
                for i in range(conn.Children.Count - 1, 0, -1):
                    try:
                        extra = conn.Children(i)
                        extra.findById("wnd[0]").Close()
                        time.sleep(0.5)
                        try:
                            extra.findById("wnd[1]/usr/btnSPOP-OPTION1").Press()
                        except Exception:
                            pass
                        estado.log(f"[OK] Sessão SAP {i} fechada.")
                        time.sleep(0.5)
                    except Exception:
                        pass
            except Exception:
                pass
            fechar_sap(session, estado)
            session = None

        # Consolidação com progress por arquivo
        estado.status = "Consolidando dados..."
        estado.log("[→] Lendo e consolidando arquivos...")

        dados_zco  = consolidar_zco1(pasta_destino, estado)
        dados_mb51 = consolidar_extra(pasta_destino, "MB51.xlsx", 16, estado) if cfg["puxar_mb51"] else []
        dados_prod = consolidar_extra(pasta_destino, "PROD.xlsx", 10, estado) if cfg["puxar_prod"] else []

        estado.status = "Injetando no Excel mestre..."
        estado.log("[→] Conectando à planilha mestre...")

        mestre = ExcelMestre(cfg["mestre_path"])
        mestre.conectar()

        estado.log(f"[→] Injetando ZCO1: {len(dados_zco)} linhas...")
        mestre.injetar_dados("ZCO1", dados_zco, n_colunas=18, formatar_data_col=8)
        estado.log("[OK] ZCO1 injetado.")

        if cfg["puxar_mb51"] and dados_mb51:
            estado.log(f"[→] Injetando MB51: {len(dados_mb51)} linhas...")
            mestre.injetar_dados("MB51", dados_mb51, n_colunas=16)
            estado.log("[OK] MB51 injetado.")

        if cfg["puxar_prod"] and dados_prod:
            estado.log(f"[→] Injetando PRODUÇÃO: {len(dados_prod)} linhas...")
            mestre.injetar_dados("PRODUÇÃO", dados_prod, n_colunas=10)
            estado.log("[OK] PRODUÇÃO injetado.")

        estado.status = "Salvando planilha..."
        mestre.salvar_e_exibir()
        estado.log("[OK] Planilha salva e exibida.")

        estado.status = f"✔  Concluído — {cfg['periodo_label']}"
        estado.log(f"[OK] Extração concluída: {len(dados_zco)} linhas ZCO1.")
        notificar_windows("SAP Bridge", f"Extração concluída — {cfg['periodo_label']}")

    except Exception as e:
        estado.log(f"[ERRO] {type(e).__name__}: {e}")
        estado.erro   = str(e)
        estado.status = "✖  Erro — verifique o log"
        notificar_windows("SAP Bridge", f"Erro: {e}", icone="Error")
        # Tenta fechar o SAP mesmo em caso de erro
        if session:
            try:
                fechar_sap(session, estado)
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()
        estado.concluido = True




def main():
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
        config["pasta_sap"]   = cfg["pasta_sap"]
        config["mestre_path"] = cfg["mestre_path"]
        config["num_threads"] = cfg["num_threads"]
        salvar_config(config)

        total_estimado = (cfg["dt_fim"] - cfg["dt_ini"]).days + 1
        if cfg["puxar_mb51"]:
            total_estimado += 1
        if cfg["puxar_prod"]:
            total_estimado += 1

        estado = Estado(progresso_total=total_estimado)

        janela = JanelaProgresso(
            titulo=f"SAP Bridge v{VERSAO}",
            subtitulo=f"{cfg['periodo_label']}  •  {Path(cfg['mestre_path']).name}",
            estado=estado,
        )
        janela.iniciar(worker_principal, estado, cfg)

        if estado.erro:
            if not messagebox.askyesno("SAP Bridge", f"Erro: {estado.erro}\n\nDeseja tentar novamente?"):
                break
        else:
            if not messagebox.askyesno("SAP Bridge", "Extração concluída!\nDeseja executar novamente?"):
                break


if __name__ == "__main__":
    main()