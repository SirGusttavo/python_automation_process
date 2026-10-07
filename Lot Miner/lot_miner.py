"""
Lot Miner — Análise de Bobinas via SSRS
Arquitetura:
  - JanelaPrincipal  : configuração inicial (wizard visual em 3 seções)
  - Estado           : bridge thread-safe worker ↔ UI
  - worker           : extração sequencial em background thread
  - JanelaProgresso  : acompanhamento em tempo real (polling 350ms)
  - formatar_excel   : formatação do arquivo de saída
"""

import os, re, json, time, logging, threading, datetime, subprocess, tempfile, html
import pandas as pd, xlwings as xw, tkinter as tk
from tkinter import filedialog, ttk
from io import StringIO

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.keys import Keys

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── IDs e URL do portal SSRS ───────────────────────────────────────────────
URL_RELATORIO = (
    "http://reports.example.com/Reports/report/Relatorios/"
    "Tissue/Produ%C3%A7%C3%A3o%20Tissue/Auditoria%20de%20Volume%20Tissue"
)
ID_INPUT   = "ReportViewerControl_ctl04_ctl03_txtValue"
ID_BOTAO   = "ReportViewerControl_ctl04_ctl00"
ID_LOADING = "ReportViewerControl_AsyncWait"
MARCADOR_RELATORIO = "Data/Hora"

# ── Leitura da planilha de entrada ─────────────────────────────────────────
COLUNA_LOTES = "I"
LINHA_INICIO = 2
ABA_LOTES    = "BOBINAS"

# ── Mapeamento de colunas ──────────────────────────────────────────────────
COLUNAS_DIV100 = ["Peso Estimado","Peso Balança","Peso Refugo","Comprimento","Diametro","Largura"]
COLUNAS_INT    = ["Qtde","Pedido","Corrida"]
COLUNAS_STR    = ["Track Num","Article","Cod. Estação","Fabrica"]
COLUNAS_NUM    = set(COLUNAS_DIV100) | set(COLUNAS_INT)

# ── Timeouts e retry ───────────────────────────────────────────────────────
TIMEOUT_RENDER   = 600
MAX_TENTATIVAS   = 2
RELOAD_APOS      = 5

# ── Configuração de Paralelismo (Edite Aqui!) ──────────────────────────────
QTD_EDGES         = 2   # Quantidade de navegadores Edge independentes (workers)
QTD_ABAS_POR_EDGE = 4   # Quantidade de abas dentro de CADA navegador Edge
QTD_ABAS          = 4   # Sinônimo direto (se alterar qualquer um deles, funciona!)

# ── Prefixos sem suporte ───────────────────────────────────────────────────
PREFIXOS_IGNORADOS: dict[str, str] = {
    "DMX": "Bobina Duramax — não disponível neste relatório",  # [REVIEW] Product category names kept
}

# ── Config persistente ─────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".lot_miner.json")

# ── SAP GUI Scripting / MB52 ────────────────────────────────────────────────
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_CONNECTION_NAME = "SAP_CONNECTION_NAME"
SAP_MB52_FILENAME = "MB52_BOBINAS.xlsx"

# ── Paleta ─────────────────────────────────────────────────────────────────
COR_AZUL     = "#1F3864"
COR_AZUL_CLR = "#EBF3FB"
COR_VERDE    = "#1A6B1A"
COR_ALERTA   = "#FFDAD9"
COR_ALERTA_F = "#C00000"
COR_REVISAR  = "#FFF3CD"   # amarelo suave — lotes para revisão manual
COR_REVISAR_F= "#7B5C00"

# Excel
XL_HEADER = "1F3864"
XL_ZEBRA  = "EBF3FB"
XL_NORMAL = "FFFFFF"
XL_ALERTA = "FFDAD9"
XL_ALERTA_F = "C00000"

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s — %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger()


# ── Config persistente ─────────────────────────────────────────────────────

def carregar_config() -> dict:
    defaults = {"saida_dir": "", "sap_saida_dir": "", "headless": True, "extrair_sap": False}
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, encoding="utf-8") as f:
                defaults.update(json.load(f))
    except Exception:
        pass
    return defaults

def salvar_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        log.debug(f"Config não salva: {e}")


# ── Estado compartilhado ───────────────────────────────────────────────────

class Estado:
    """
    Bridge thread-safe entre worker (background) e JanelaProgresso (main thread).
    Campos lidos pela UI; escritos pelo worker via tick(). Lock protege escrita.
    """
    def __init__(self, total: int):
        self.total      = total
        self.atual      = 0
        self.lote_atual = "—"
        self.ultimo_ok  = True
        self.sucesso    = 0
        self.falhas_n   = 0
        self.revisar_n  = 0   # lotes com divergência persistente → revisão manual
        self.media_s    = 0.0
        self.restante_m = 0.0
        self.elapsed_m  = 0.0
        self.concluido  = False
        self.parar      = threading.Event()
        self._lock      = threading.Lock()

    def tick(self, lote: str, ok: bool, revisar: bool,
             media: float, restante: float, elapsed: float):
        with self._lock:
            self.atual      += 1
            self.lote_atual  = lote
            self.ultimo_ok   = ok and not revisar
            self.media_s     = media
            self.restante_m  = restante
            self.elapsed_m   = elapsed
            if ok:   self.sucesso  += 1
            else:    self.falhas_n += 1
            if revisar: self.revisar_n += 1


# ── Clipboard para SAP GUI ──────────────────────────────────────────────────

def EnviarParaAreaDeTransferencia(texto: str):
    """
    Envia texto para o Clipboard do Windows para uso no SAP GUI Scripting.
    """
    import win32clipboard

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


# ── Extração MB52 via SAP GUI ───────────────────────────────────────────────

def extrair_mb52_bobinas_sap(
    pasta_saida: str,
    categoria_professional: bool,
    categoria_higienicos: bool) -> str:
    """
    Executa a MB52 no SAP GUI Scripting e exporta a planilha usada pelo Lot Miner.

    O arquivo MB52_BOBINAS.xlsx salvo na pasta configurada torna-se
    automaticamente a entrada do Lot Miner.
    """
    if not pasta_saida:
        raise ValueError("Pasta de saída do SAP não informada.")

    os.makedirs(pasta_saida, exist_ok=True)
    caminho = os.path.join(pasta_saida, SAP_MB52_FILENAME)

    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "A extração direta do SAP requer o pacote 'pywin32'. "
            "Instale com: pip install pywin32"
        ) from e

    pythoncom.CoInitialize()
    try:
        # Remove a exportação anterior.
        try:
            if os.path.exists(caminho):
                os.remove(caminho)
        except PermissionError as e:
            raise RuntimeError(
                f"Feche o arquivo antes de executar novamente:\n{caminho}"
            ) from e

        # Conecta ao SAP GUI; se necessário, abre o SAP Logon.
        sap_gui = None
        deadline = time.time() + 30

        while time.time() < deadline:
            try:
                sap_gui = win32com.client.GetObject("SAPGUI")
                if sap_gui is not None:
                    break
            except Exception:
                sap_gui = None
            time.sleep(1)

        if sap_gui is None:
            if not os.path.exists(SAP_LOGON_PATH):
                raise RuntimeError(
                    "SAP GUI não foi encontrado e o SAP Logon não existe em:\n"
                    f"{SAP_LOGON_PATH}"
                )

            subprocess.Popen(
                [SAP_LOGON_PATH],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )

            deadline = time.time() + 30
            while time.time() < deadline:
                try:
                    sap_gui = win32com.client.GetObject("SAPGUI")
                    if sap_gui is not None:
                        break
                except Exception:
                    sap_gui = None
                time.sleep(1)

        if sap_gui is None:
            raise RuntimeError(
                "Não foi possível acessar o SAP GUI. "
                "Abra o SAP Logon e tente novamente."
            )

        sap_app = sap_gui.GetScriptingEngine

        if sap_app.Children.Count == 0:
            sap_con = sap_app.OpenConnection(SAP_CONNECTION_NAME, True)
            time.sleep(2)
        else:
            sap_con = sap_app.Children(0)

        if sap_con.Children.Count == 0:
            raise RuntimeError("O SAP GUI foi encontrado, mas não há sessão aberta.")

        session = sap_con.Children(0)

        # ── MB52 ────────────────────────────────────────────────────────
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Limpeza dos filtros.
        for control_id in (
            "wnd[0]/usr/ctxtMATNR-LOW",
            "wnd[0]/usr/ctxtWERKS-LOW",
            "wnd[0]/usr/ctxtLGORT-LOW",
            "wnd[0]/usr/ctxtCHARG-LOW",
            "wnd[0]/usr/ctxtMATART-LOW",
        ):
            try:
                session.findById(control_id).Text = ""
            except Exception:
                pass

        # Centro.
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1000"

        # ── Categorias selecionadas ─────────────────────────────────────────────

        depositos = []

        # Professional e Formatados  # [REVIEW] Product category names kept
        if categoria_professional:
            depositos.append("LP01")

        # Higiênicos e Duramax  # [REVIEW] Product category names kept
        if categoria_higienicos:
            depositos.extend([
                "LI04",
                "LI05",
                "LI06",
                "LI07",
                "LI08",
                "LI12",
                "LI13",
                "LI14",
                "LI15"
            ])

        if not depositos:
            raise RuntimeError(
                "Nenhuma categoria foi selecionada."
            )

        # Gerar arquivo TXT para importação
        txt_depositos_path = os.path.join(pasta_saida, "depositos_mb52.txt")
        with open(txt_depositos_path, "w", encoding="utf-8") as f:
            for dep in depositos:
                f.write(f"{dep}\n")

        # Abre seleção múltipla
        session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press()
        
        # Limpar seleção existente
        session.findById("wnd[1]/tbar[0]/btn[16]").Press()
        
        # Importar do TXT
        session.findById("wnd[1]/tbar[0]/btn[23]").Press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").Text = pasta_saida
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "depositos_mb52.txt"
        session.findById("wnd[2]/tbar[0]/btn[0]").Press()

        log.info(f"Depósitos SAP MB52 inseridos via TXT: {len(depositos)}")

        # Confirmar a seleção múltipla
        session.findById("wnd[1]/tbar[0]/btn[8]").Press()

        # Filtros nativos.
        session.findById("wnd[0]/usr/chkXMCHB").Selected = True
        session.findById("wnd[0]/usr/chkNOZERO").Selected = True

        # Variante.
        session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/USER_LAYOUT"
        session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus()

        # Executa e exporta.
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(43)

        session.findById("wnd[1]/tbar[0]/btn[20]").Press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta_saida
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = SAP_MB52_FILENAME
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()

        # Aguarda o arquivo existir.
        deadline = time.time() + 30
        while time.time() < deadline:
            if os.path.exists(caminho):
                try:
                    if os.path.getsize(caminho) > 0:
                        break
                except OSError:
                    pass
            time.sleep(1)

        if not os.path.exists(caminho):
            raise RuntimeError(
                "O SAP não gerou a MB52 no caminho esperado:\n"
                f"{caminho}"
            )

        # Fecha somente a instância do Excel criada pela exportação do SAP.
        try:
            wb_sap = win32com.client.GetObject(caminho)
            app_sap = wb_sap.Parent
            app_sap.DisplayAlerts = False
            wb_sap.Close(SaveChanges=False)
            try:
                if app_sap.Workbooks.Count == 0:
                    app_sap.Quit()
            except Exception:
                pass
        except Exception as e:
            log.debug(f"Excel da exportação MB52 não precisou ser fechado: {e}")

        # Volta ao menu principal sem encerrar a conexão SAP.
        try:
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            log.debug(f"Não foi possível retornar ao menu SAP: {e}")

        log.info(f"MB52 exportada pelo SAP: {caminho}")
        return caminho

    finally:
        pythoncom.CoUninitialize()


# ── Janela Principal ────────────────────────────────────────────────────────

class JanelaPrincipal:
    """
    Wizard visual.

    Se "Extrair planilha MB52 diretamente do SAP" estiver ligado:
      SAP -> MB52_BOBINAS.xlsx -> Lot Miner -> SSRS.

    Se estiver desligado:
      usuário seleciona manualmente a planilha de entrada.
    """
    def __init__(self):
        cfg = carregar_config()

        self.entrada = cfg.get("entrada", "")
        self.saida_dir = cfg.get("saida_dir", "")
        self.sap_saida_dir = cfg.get("sap_saida_dir", "")
        
        if not isinstance(self.sap_saida_dir, str):
            self.sap_saida_dir = ""

        self.headless = bool(cfg.get("headless", True))
        self.extrair_sap = bool(cfg.get("extrair_sap", False))
        
        self.categoria_professional = bool(cfg.get("categoria_professional", False))
        self.categoria_higienicos = bool(cfg.get("categoria_higienicos", False))

        self.confirmado = False
        self.saida = ""

    def abrir(self) -> bool:
        root = tk.Tk()
        root.title("Lot Miner — Análise de Bobinas")
        root.resizable(False, False)
        root.attributes("-topmost", True)
        root.after(120, lambda: root.attributes("-topmost", False))
        root.lift()
        root.focus_force()
        root.configure(bg="#F4F6FA")

        # ------------------------------------------------------------
        # TAMANHO DA JANELA
        # ------------------------------------------------------------

        w, h = 560, 620

        root.update_idletasks()

        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2

        root.geometry(f"{w}x{h}+{x}+{y}")

        # ------------------------------------------------------------
        # CABEÇALHO
        # ------------------------------------------------------------

        hdr = tk.Frame(
            root,
            bg=COR_AZUL,
            height=60
        )

        hdr.pack(
            fill="x"
        )

        hdr.pack_propagate(False)

        tk.Label(
            hdr,
            text="  🔍  Lot Miner",
            font=("Calibri", 15, "bold"),
            fg="white",
            bg=COR_AZUL,
            anchor="w"
        ).pack(
            side="left",
            padx=16,
            pady=10
        )

        tk.Label(
            hdr,
            text="Análise de Bobinas via SSRS  ",
            font=("Calibri", 9),
            fg="#A8C4E0",
            bg=COR_AZUL,
            anchor="e"
        ).pack(
            side="right",
            padx=14
        )

        # ------------------------------------------------------------
        # CORPO
        # ------------------------------------------------------------

        body = tk.Frame(
            root,
            bg="#F4F6FA"
        )

        body.pack(
            fill="both",
            expand=True,
            padx=22,
            pady=10
        )

        # ------------------------------------------------------------
        # FUNÇÃO DE SEÇÃO
        # ------------------------------------------------------------

        def secao(parent, numero, titulo):

            f = tk.Frame(
                parent,
                bg="#F4F6FA"
            )

            f.pack(
                fill="x",
                pady=(8, 0)
            )

            tk.Label(
                f,
                text=f" {numero} ",
                font=("Calibri", 9, "bold"),
                fg="white",
                bg=COR_AZUL,
                width=2
            ).pack(
                side="left"
            )

            tk.Label(
                f,
                text=f"  {titulo}",
                font=("Calibri", 10, "bold"),
                fg=COR_AZUL,
                bg="#F4F6FA"
            ).pack(
                side="left"
            )

            tk.Frame(
                parent,
                bg="#C5D5E8",
                height=1
            ).pack(
                fill="x",
                pady=(2, 4)
            )

        # ============================================================
        # SEÇÃO 1 — ORIGEM
        # ============================================================

        secao(
            body,
            "1",
            "ORIGEM DA PLANILHA DE LOTES"
        )

        frm_origem = tk.Frame(
            body,
            bg="#F4F6FA"
        )

        frm_origem.pack(
            fill="x"
        )

        # ------------------------------------------------------------
        # VARIÁVEIS
        # ------------------------------------------------------------

        var_sap = tk.BooleanVar(
            value=self.extrair_sap
        )

        var_professional = tk.BooleanVar(
            value=self.categoria_professional
        )

        var_higienicos = tk.BooleanVar(
            value=self.categoria_higienicos
        )

        var_origem = tk.StringVar(
            value=(
                "🔵  Utilize uma planilha MB52 já existente"
                if not self.extrair_sap
                else
                "🟢  A MB52 será extraída diretamente do SAP"
            )
        )

        var_info = tk.StringVar(
            value=""
        )

        var_sap_info = tk.StringVar(
            value=""
        )

        # ------------------------------------------------------------
        # STATUS DA ORIGEM
        # ------------------------------------------------------------

        lbl_origem = tk.Label(
            frm_origem,
            textvariable=var_origem,
            font=("Calibri", 7),
            fg="#555",
            bg="#F4F6FA",
            anchor="w"
        )

        # ------------------------------------------------------------
        # ÁREA PRINCIPAL DOS MODOS
        #
        # Altura fixa:
        # isso impede que a Seção 2 fique pulando quando o checkbox
        # SAP for marcado/desmarcado.
        # ------------------------------------------------------------

        frm_modos = tk.Frame(
            frm_origem,
            bg="#F4F6FA",
            height=38
        )

        frm_modos.pack(
            fill="x",
            pady=(5, 0)
        )

        frm_modos.pack_propagate(False)

        # ============================================================
        # ÁREA FIXA — CONFIGURAÇÃO SAP
        # ============================================================

        frm_config_sap = tk.Frame(
            frm_origem,
            bg="#F4F6FA"
        )

        frm_config_sap.pack(
            fill="x",
            pady=(4, 0)
        )

        # ============================================================
        # MODO MANUAL
        # ============================================================

        frm_manual = tk.Frame(
            frm_modos,
            bg="#F4F6FA"
        )

        frm_manual.place(
            x=0,
            y=0,
            relwidth=1.0,
            height=38
        )

        frm_arquivo = tk.Frame(
            frm_manual,
            bg="white",
            relief="solid",
            bd=1,
            height=38
        )

        frm_arquivo.pack(
            fill="x"
        )

        frm_arquivo.pack_propagate(False)

        var_entrada = tk.StringVar(
            value=(
                os.path.basename(self.entrada)
                if self.entrada
                else "📄  Nenhum arquivo selecionado"
            )
        )

        lbl_arq = tk.Label(
            frm_arquivo,
            textvariable=var_entrada,
            fg="#555",
            anchor="w",
            font=("Calibri", 8),
            bg="white",
            padx=6
        )

        lbl_arq.pack(
            side="left",
            fill="both",
            expand=True
        )

        def escolher_entrada():

            p = filedialog.askopenfilename(
                parent=root,
                title="Selecione a planilha de lotes",
                filetypes=[
                    ("Excel", "*.xlsx *.xlsm *.xls")
                ]
            )

            if not p:
                return

            self.entrada = p

            var_entrada.set(
                f"📄  {os.path.basename(p)}"
            )

            lbl_arq.config(
                fg="#222"
            )

            var_info.set(
                "⏳  Detectando coluna..."
            )

            def _info_bg():

                try:

                    col, n = detectar_coluna_lotes_fast(
                        p
                    )

                    msg = (
                        f"✓  {n} lotes · "
                        f"coluna {col} detectada automaticamente"
                    )

                except Exception:

                    msg = (
                        "✓  Arquivo selecionado"
                    )

                root.after(
                    0,
                    lambda: var_info.set(msg)
                )

            threading.Thread(
                target=_info_bg,
                daemon=True
            ).start()

        tk.Button(
            frm_arquivo,
            text="Procurar...",
            command=escolher_entrada,
            width=9,
            font=("Calibri", 8),
            bg="#E8EEF7",
            relief="flat",
            cursor="hand2"
        ).pack(
            side="right",
            padx=4
        )

        # ============================================================
        # MODO SAP
        # ============================================================

        frm_sap = tk.Frame(
            frm_modos,
            bg="white",
            relief="solid",
            bd=1
        )

        var_sap_saida = tk.StringVar(
            value=(
                self.sap_saida_dir
                if self.sap_saida_dir
                else "📁  Nenhuma pasta selecionada"
            )
        )

        lbl_sap_saida = tk.Label(
            frm_sap,
            textvariable=var_sap_saida,
            fg="#555",
            width=48,
            anchor="w",
            font=("Calibri", 8),
            bg="white",
            padx=6
        )

        lbl_sap_saida.pack(
            side="left",
            fill="x",
            expand=True
        )

        def escolher_sap_saida():

            p = filedialog.askdirectory(
                parent=root,
                title="Pasta onde a MB52 do SAP será salva",
                initialdir=(
                    self.sap_saida_dir
                    if self.sap_saida_dir and os.path.isdir(self.sap_saida_dir)
                    else None
                )
            )

            if not p:
                return

            self.sap_saida_dir = p

            var_sap_saida.set(
                p
            )

            lbl_sap_saida.config(
                fg="#222"
            )

            var_sap_salvo.set(
                "  ✓ pasta SAP será salva"
            )

        tk.Button(
            frm_sap,
            text="Procurar...",
            command=escolher_sap_saida,
            width=9,
            font=("Calibri", 8),
            bg="#E8EEF7",
            relief="flat",
            cursor="hand2"
        ).pack(
            side="right",
            padx=4
        )

        # ------------------------------------------------------------
        # Indicador de pasta SAP salva
        # ------------------------------------------------------------

        var_sap_salvo = tk.StringVar(
            value=(
                "  ✓ pasta salva da última sessão"
                if self.sap_saida_dir
                else ""
            )
        )

        lbl_sap_salvo = tk.Label(
            frm_config_sap,
            textvariable=var_sap_salvo,
            font=("Calibri", 7),
            fg=COR_VERDE,
            bg="#F4F6FA",
            anchor="w"
        )

        # ============================================================
        # CATEGORIAS MB52
        # ============================================================

        frm_categorias = tk.LabelFrame(
            frm_config_sap,
            text=" Categorias da MB52 ",
            font=("Calibri", 8, "bold"),
            fg=COR_AZUL,
            bg="#F4F6FA",
            padx=8,
            pady=4
        )

        frm_cat_itens = tk.Frame(
            frm_categorias,
            bg="#F4F6FA"
        )

        frm_cat_itens.pack(
            fill="x"
        )

        # Col 0: Professional  # [REVIEW] Product category names kept
        frm_prof = tk.Frame(
            frm_cat_itens,
            bg="#F4F6FA"
        )

        frm_prof.pack(
            side="left",
            padx=(0, 12)
        )

        tk.Checkbutton(
            frm_prof,
            text="Professional e Formatados",  # [REVIEW] Product category names kept
            variable=var_professional,
            bg="#F4F6FA",
            font=("Calibri", 8),
            cursor="hand2",
            command=lambda: setattr(
                self,
                "categoria_professional",
                var_professional.get()
            )
        ).pack(
            anchor="w"
        )

        tk.Label(
            frm_prof,
            text="LP01",
            font=("Calibri", 7),
            fg="#888",
            bg="#F4F6FA"
        ).pack(
            anchor="w",
            padx=(16, 0)
        )

        # Col 1: Higiênicos  # [REVIEW] Product category names kept
        frm_higien = tk.Frame(
            frm_cat_itens,
            bg="#F4F6FA"
        )

        frm_higien.pack(
            side="left"
        )

        tk.Checkbutton(
            frm_higien,
            text="Higiênicos e Duramax",  # [REVIEW] Product category names kept
            variable=var_higienicos,
            bg="#F4F6FA",
            font=("Calibri", 8),
            cursor="hand2",
            command=lambda: setattr(
                self,
                "categoria_higienicos",
                var_higienicos.get()
            )
        ).pack(
            anchor="w"
        )

        tk.Label(
            frm_higien,
            text="LI04–LI08 · LI12–LI15",
            font=("Calibri", 7),
            fg="#888",
            bg="#F4F6FA"
        ).pack(
            anchor="w",
            padx=(16, 0)
        )

        # ============================================================
        # LÓGICA DE ALTERNÂNCIA
        # ============================================================

        def atualizar_origem(*_):

            usar_sap = bool(
                var_sap.get()
            )

            self.extrair_sap = usar_sap

            # --------------------------------------------------------
            # Remove os elementos da área variável
            # --------------------------------------------------------

            frm_manual.place_forget()
            frm_sap.place_forget()
            lbl_info.pack_forget()
            lbl_sap_salvo.pack_forget()
            frm_categorias.pack_forget()

            # --------------------------------------------------------
            # MODO SAP
            # --------------------------------------------------------

            if usar_sap:

                var_origem.set(
                    "🟢  A MB52 será extraída diretamente do SAP"
                )

                frm_sap.place(
                    x=0,
                    y=0,
                    relwidth=1.0,
                    height=38
                )

                lbl_sap_salvo.pack(
                    fill="x",
                    pady=(1, 0)
                )

                frm_categorias.pack(
                    fill="x",
                    pady=(6, 0)
                )
                
                # Expandir janela para comportar as categorias
                try:
                    root.geometry(f"560x620")
                except Exception:
                    pass

            # --------------------------------------------------------
            # MODO MANUAL
            # --------------------------------------------------------

            else:

                var_origem.set(
                    "🔵  Utilize uma planilha MB52 já existente"
                )

                frm_manual.place(
                    x=0,
                    y=0,
                    relwidth=1.0,
                    height=38
                )

                lbl_info.pack(
                    fill="x",
                    pady=(2, 0)
                )
                
                # Forçar o frame vazio a encolher (contorno de comportamento do Tkinter)
                frm_config_sap.config(height=1)

                # Encolher janela para remover o espaço vazio das categorias
                try:
                    root.geometry(f"560x540")
                except Exception:
                    pass

        # ------------------------------------------------------------
        # Status/informação da planilha manual
        # ------------------------------------------------------------

        lbl_info = tk.Label(
            frm_origem,
            textvariable=var_info,
            font=("Calibri", 7),
            fg="#555",
            bg="#F4F6FA",
            anchor="w"
        )

        # ============================================================
        # CHECKBOX PRINCIPAL
        # ============================================================

        tk.Checkbutton(
            frm_origem,
            text="Extrair planilha MB52 diretamente do SAP",
            variable=var_sap,
            command=atualizar_origem,
            font=("Calibri", 9, "bold"),
            bg="#F4F6FA",
            activebackground="#F4F6FA",
            cursor="hand2"
        ).pack(
            anchor="w",
            pady=(2, 3)
        )

        # Status
        lbl_origem.pack(
            fill="x"
        )

        # ============================================================
        # TEXTO DA COLUNA DE LOTES
        # ============================================================

        tk.Label(
            frm_origem,
            text=(
                f"Coluna de lotes detectada automaticamente · "
                f"aba '{ABA_LOTES}' (ou 1ª aba)"
            ),
            font=("Calibri", 7),
            fg="#999",
            bg="#F4F6FA",
            anchor="w"
        ).pack(
            fill="x",
            pady=(2, 0)
        )

        # Estado inicial
        atualizar_origem()

        # ============================================================
        # SEÇÃO 2 — SAÍDA DO LOT MINER
        # ============================================================

        secao(
            body,
            "2",
            "PASTA DE SAÍDA DO LOT MINER"
        )

        frm2 = tk.Frame(
            body,
            bg="white",
            relief="solid",
            bd=1
        )

        frm2.pack(
            fill="x",
            ipady=4
        )

        saida_val = (
            self.saida_dir
            if self.saida_dir
            else "📁  Nenhuma pasta selecionada"
        )

        var_saida = tk.StringVar(
            value=saida_val
        )

        lbl_pasta = tk.Label(
            frm2,
            textvariable=var_saida,
            fg="#555",
            width=50,
            anchor="w",
            font=("Calibri", 8),
            bg="white",
            padx=6
        )

        lbl_pasta.pack(
            side="left"
        )

        var_salvo = tk.StringVar(
            value=(
                "✓  Pasta salva da última sessão"
                if self.saida_dir
                else ""
            )
        )

        tk.Label(
            body,
            textvariable=var_salvo,
            font=("Calibri", 7),
            fg=COR_VERDE,
            bg="#F4F6FA",
            anchor="w"
        ).pack(
            fill="x",
            pady=(1, 0)
        )

        def escolher_saida():

            p = filedialog.askdirectory(
                parent=root,
                title="Pasta de saída do Lot Miner",
                initialdir=(
                    self.saida_dir
                    if self.saida_dir
                    else None
                )
            )

            if p:

                self.saida_dir = p

                var_saida.set(
                    p
                )

                lbl_pasta.config(
                    fg="#222"
                )

                var_salvo.set(
                    "✓  Será salvo ao iniciar"
                )

        tk.Button(
            frm2,
            text="Procurar...",
            command=escolher_saida,
            width=9,
            font=("Calibri", 8),
            bg="#E8EEF7",
            relief="flat",
            cursor="hand2"
        ).pack(
            side="right",
            padx=4
        )

        # ============================================================
        # SEÇÃO 3 — OPÇÕES
        # ============================================================

        secao(
            body,
            "3",
            "OPÇÕES"
        )

        frm3 = tk.Frame(
            body,
            bg="#F4F6FA"
        )

        frm3.pack(
            fill="x"
        )

        var_h = tk.BooleanVar(
            value=self.headless
        )

        var_modo = tk.StringVar(
            value=(
                "🟢 Headless ativo — sem janela do browser (recomendado)"
                if self.headless
                else
                "🔵 Com janela — abre o Edge normalmente"
            )
        )

        def toggle_headless():

            var_modo.set(
                "🟢 Headless ativo — sem janela do browser (recomendado)"
                if var_h.get()
                else
                "🔵 Com janela — abre o Edge normalmente"
            )

        tk.Checkbutton(
            frm3,
            textvariable=var_modo,
            variable=var_h,
            command=toggle_headless,
            font=("Calibri", 9),
            bg="#F4F6FA",
            activebackground="#F4F6FA",
            cursor="hand2"
        ).pack(
            anchor="w",
            pady=4
        )

        tk.Label(
            frm3,
            text=(
                "  ⚡ Lotes com divergência de peso "
                "são retentados automaticamente"
            ),
            font=("Calibri", 7),
            fg="#888",
            bg="#F4F6FA"
        ).pack(
            anchor="w"
        )

        # ============================================================
        # STATUS
        # ============================================================

        tk.Frame(
            body,
            bg="#C5D5E8",
            height=1
        ).pack(
            fill="x",
            pady=(10, 4)
        )

        var_status = tk.StringVar(
            value=""
        )

        tk.Label(
            body,
            textvariable=var_status,
            fg=COR_ALERTA_F,
            font=("Calibri", 8),
            bg="#F4F6FA",
            wraplength=510
        ).pack()

        # ============================================================
        # BOTÕES
        # ============================================================

        frm_btn = tk.Frame(
            body,
            bg="#F4F6FA"
        )

        frm_btn.pack(
            pady=(6, 0)
        )

        # ============================================================
        # INICIAR
        # ============================================================

        def iniciar():
            self.extrair_sap = bool(var_sap.get())
            self.headless = bool(var_h.get())

            if self.extrair_sap:
                if not self.sap_saida_dir:
                    var_status.set("⚠  Selecione a pasta onde a MB52 será salva pelo SAP.")
                    return
                if not os.path.isdir(self.sap_saida_dir):
                    var_status.set("⚠  A pasta selecionada para o SAP não existe mais.")
                    return

                self.categoria_professional = bool(var_professional.get())
                self.categoria_higienicos = bool(var_higienicos.get())

                if not self.categoria_professional and not self.categoria_higienicos:
                    var_status.set("⚠  Selecione Professional, Higiênicos ou as duas categorias.")
                    return

                self.entrada = ""
            else:
                if not self.entrada or not os.path.exists(self.entrada):
                    var_status.set("⚠  Selecione a planilha de lotes.")
                    return

                self.categoria_professional = False
                self.categoria_higienicos = False

            if not self.saida_dir:
                var_status.set("⚠  Selecione a pasta de saída do Lot Miner.")
                return

            if not os.path.isdir(self.saida_dir):
                var_status.set("⚠  A pasta de saída do Lot Miner não existe mais.")
                return

            salvar_config({
                "saida_dir": self.saida_dir,
                "sap_saida_dir": self.sap_saida_dir,
                "headless": self.headless,
                "extrair_sap": self.extrair_sap,
                "categoria_professional": self.categoria_professional,
                "categoria_higienicos": self.categoria_higienicos,
            })

            self.confirmado = True
            root.destroy()

        # ============================================================
        # BOTÃO CANCELAR
        # ============================================================

        tk.Button(
            frm_btn,
            text="Cancelar",
            command=root.destroy,
            width=10,
            pady=5,
            relief="flat",
            bg="#F4F6FA",
            cursor="hand2"
        ).pack(
            side="left",
            padx=6
        )

        # ============================================================
        # BOTÃO INICIAR
        # ============================================================

        tk.Button(
            frm_btn,
            text="  ▶  Iniciar extração  ",
            command=iniciar,
            bg=COR_AZUL,
            fg="white",
            font=("Calibri", 10, "bold"),
            padx=12,
            pady=5,
            relief="flat",
            cursor="hand2",
            activebackground="#2D5299"
        ).pack(
            side="left",
            padx=6
        )

        # ============================================================
        # LOOP PRINCIPAL DO TKINTER
        # ============================================================

        root.mainloop()

        # ============================================================
        # APÓS FECHAR A INTERFACE
        # ============================================================

        if self.confirmado:

            now = datetime.datetime.now()

            nome = (
                f"Análise de Bobinas - "
                f"{now.day:02d}."
                f"{now.month:02d}.xlsx"
            )

            self.saida = os.path.join(
                self.saida_dir,
                nome
            )

        return self.confirmado

# ── Janela de Progresso ────────────────────────────────────────────────────

class JanelaProgresso:
    """
    Janela em tempo real executada na main thread via polling (350ms).
    Mostra: barra, %, lote atual, métricas, contadores e botão Parar.
    Lotes com divergência persistente são contados separadamente (⚠️).
    """
    def __init__(self, estado: Estado):
        self.estado = estado

    def run(self):
        root = tk.Tk()
        root.title("Lot Miner — Extração")
        root.resizable(False, False)
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
        root.lift()
        root.configure(bg="#F4F6FA")

        w, h = 500, 320
        root.update_idletasks()
        x = (root.winfo_screenwidth()  - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

        # Cabeçalho dinâmico
        hdr = tk.Frame(root, bg=COR_AZUL, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        lbl_titulo = tk.Label(hdr, text="  ⚙  Extração em andamento",
                 font=("Calibri", 11, "bold"), fg="white",
                 bg=COR_AZUL, anchor="w")
        lbl_titulo.pack(side="left", padx=14, pady=8)

        body = tk.Frame(root, bg="#F4F6FA")
        body.pack(fill="both", expand=True, padx=22, pady=10)

        # Barra de progresso + %
        frm_bar = tk.Frame(body, bg="#F4F6FA")
        frm_bar.pack(fill="x", pady=(2, 2))
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("LM.Horizontal.TProgressbar",
                         troughcolor="#D0DCF0",
                         background=COR_AZUL,
                         thickness=20)
        pbar = ttk.Progressbar(frm_bar, length=390, mode="determinate",
                               maximum=self.estado.total,
                               style="LM.Horizontal.TProgressbar")
        pbar.pack(side="left")
        var_pct = tk.StringVar(value="0%")
        tk.Label(frm_bar, textvariable=var_pct,
                 font=("Calibri", 9, "bold"), fg=COR_AZUL,
                 bg="#F4F6FA", width=5).pack(side="left", padx=6)

        # Contador de lotes
        var_count = tk.StringVar(value="0 / 0  lotes")
        tk.Label(body, textvariable=var_count,
                 font=("Calibri", 12, "bold"), fg=COR_AZUL,
                 bg="#F4F6FA").pack(anchor="w", pady=(4, 0))

        # Último lote (cor dinâmica: verde ok, amarelo revisar, vermelho falha)
        var_lote = tk.StringVar(value="Aguardando primeiro lote...")
        lbl_lote = tk.Label(body, textvariable=var_lote,
                 font=("Calibri", 9), fg="#666",
                 bg="#F4F6FA", anchor="w")
        lbl_lote.pack(fill="x", pady=(2, 6))

        # Grid de métricas
        frm_stats = tk.Frame(body, bg="#EDF2F7", relief="flat")
        frm_stats.pack(fill="x", pady=2)
        var_media   = tk.StringVar(value="—")
        var_eta     = tk.StringVar(value="—")
        var_elapsed = tk.StringVar(value="—")

        def _stat(txt, var, col):
            tk.Label(frm_stats, text=txt, font=("Calibri", 7),
                     fg="#777", bg="#EDF2F7").grid(
                         row=0, column=col*2, padx=(10,2), pady=7)
            tk.Label(frm_stats, textvariable=var,
                     font=("Calibri", 9, "bold"), fg=COR_AZUL,
                     bg="#EDF2F7").grid(
                         row=0, column=col*2+1, padx=(0,10), pady=7)

        _stat("Média/lote:", var_media,   0)
        _stat("Restante:",   var_eta,     1)
        _stat("Decorrido:",  var_elapsed, 2)

        # Contadores sucesso / falha / revisar
        var_ok = tk.StringVar(value="✅ 0   ❌ 0")
        lbl_ok = tk.Label(body, textvariable=var_ok,
                 font=("Calibri", 9), fg="#333", bg="#F4F6FA")
        lbl_ok.pack(pady=(4, 2))

        def parar():
            btn_parar.config(state="disabled",
                             text="⏳  Aguardando lote atual...")
            lbl_titulo.config(text="  ⏸  Parando — aguarde o lote atual")
            self.estado.parar.set()

        btn_parar = tk.Button(root, text="⏹  Parar e Salvar Parcial",
                               command=parar, bg=COR_ALERTA_F, fg="white",
                               font=("Calibri", 10, "bold"), padx=14, pady=6,
                               relief="flat", cursor="hand2",
                               activebackground="#900000")
        btn_parar.pack(pady=(0, 14))

        def poll():
            e  = self.estado
            n, t = e.atual, e.total
            pct = int(n / t * 100) if t else 0

            pbar["value"] = n
            var_pct.set(f"{pct}%")
            var_count.set(f"{n} / {t}  lotes")
            var_media.set(f"{e.media_s:.1f}s")
            var_eta.set(f"~{e.restante_m:.1f} min")
            var_elapsed.set(f"{e.elapsed_m:.1f} min")

            ok_txt = f"✅ {e.sucesso}   ❌ {e.falhas_n}"
            if e.revisar_n:
                ok_txt += f"   ⚠️  {e.revisar_n} para revisar"
            var_ok.set(ok_txt)

            if n > 0:
                if e.revisar_n > 0 and e.ultimo_ok:
                    # último tick foi sucesso mas com divergência
                    lbl_lote.config(fg="#7B5C00",
                        text=f"⚠️   Último: {e.lote_atual}  (revisar manualmente)")
                elif e.ultimo_ok:
                    lbl_lote.config(fg=COR_VERDE,
                        text=f"✅  Último: {e.lote_atual}")
                else:
                    lbl_lote.config(fg=COR_ALERTA_F,
                        text=f"❌  Último: {e.lote_atual}  (falhou)")

            if e.concluido:
                style.configure("LM.Horizontal.TProgressbar", background="#1A6B1A")
                lbl_titulo.config(text="  ✅  Extração concluída!")
                btn_parar.config(state="disabled", text="✅  Concluído")
                root.after(1800, root.destroy)
                return

            if e.parar.is_set() and n >= t:
                style.configure("LM.Horizontal.TProgressbar", background="#E07000")
                root.after(800, root.destroy)
                return

            root.after(350, poll)

        root.after(350, poll)
        root.protocol("WM_DELETE_WINDOW", parar)
        root.mainloop()


# ── Notificação Windows ────────────────────────────────────────────────────

def notificar_windows(titulo: str, mensagem: str, arquivo: str = ""):
    """
    Exibe uma notificação nativa do Windows, com fallback para Balloon Tip.
    Se arquivo for informado, permite abrir a pasta/arquivo no Explorer.
    """
    try:
        t = str(titulo or "").replace("\r", " ").replace("\n", " ").strip()
        m = str(mensagem or "").replace("\r", " ").replace("\n", " ").strip()

        t_xml = html.escape(t, quote=False)
        m_xml = html.escape(m, quote=False)

        # ------------------------------------------------------------
        # COMANDO OPCIONAL PARA ABRIR O ARQUIVO
        # ------------------------------------------------------------

        abrir = ""

        if arquivo:
            arquivo_ps = str(arquivo).replace("'", "''")

            abrir = f"""
try {{
    Start-Process explorer.exe -ArgumentList "/select,`"{arquivo_ps}`""
}} catch {{}}
"""

        # ------------------------------------------------------------
        # POWERSHELL
        # ------------------------------------------------------------

        ps = f"""$ErrorActionPreference = 'SilentlyContinue'

$t = @'
{t_xml}
'@

$m = @'
{m_xml}
'@

$ok = $false

# ============================================================
# TOAST WINDOWS 10/11
# ============================================================

try {{

    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]

    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType=WindowsRuntime]

    $AppId = '{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe'

    # --------------------------------------------------------
    # XML CORRETO DO TOAST
    # --------------------------------------------------------

    $xml = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>$t</text>
            <text>$m</text>
        </binding>
    </visual>
</toast>
"@

    $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
    $doc.LoadXml($xml)

    $toast = New-Object Windows.UI.Notifications.ToastNotification $doc

    $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)

    $notifier.Show($toast)

    $ok = $true

    Start-Sleep -Seconds 6

}} catch {{

    $ok = $false

}}

# ============================================================
# FALLBACK — BALLOON TIP
# ============================================================

if (-not $ok) {{

    try {{

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $n = New-Object System.Windows.Forms.NotifyIcon

        $n.Icon = [System.Drawing.SystemIcons]::Information

        $n.BalloonTipTitle = $t
        $n.BalloonTipText = $m
        $n.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info

        $n.Visible = $true

        $n.ShowBalloonTip(9000)

        [System.Windows.Forms.Application]::DoEvents()

        Start-Sleep -Seconds 10

        $n.Visible = $false
        $n.Dispose()

    }} catch {{}}

}}

# ============================================================
# ABRIR ARQUIVO/PASTA, SE INFORMADO
# ============================================================

{abrir}

# ============================================================
# REMOVE O SCRIPT TEMPORÁRIO
# ============================================================

Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""

        fd, path = tempfile.mkstemp(suffix=".ps1", prefix="lotm_")
        os.close(fd)

        # IMPORTANTE: Windows PowerShell 5.1 precisa do BOM para UTF-8.
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(ps)

        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive",
             "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", path],
            creationflags=flags, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except Exception as e:
        try:
            log.debug(f"Notif: {e}")
        except Exception:
            pass


# ── Leitura de lotes ───────────────────────────────────────────────────────

_PADRAO_LOTE = re.compile(r'^[A-Z]{8}$')

def detectar_coluna_lotes_fast(caminho: str) -> tuple[str, int]:
    """
    Detecta a coluna de lotes sem COM/xlwings.
    Lote = exatamente 8 letras maiusculas (SKAABBXD, etc).
    Retorna (letra_coluna, n_lotes). Fallback para COLUNA_LOTES.
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    try:
        wb = load_workbook(caminho, read_only=True, data_only=True)
        ws = wb[ABA_LOTES] if ABA_LOTES in wb.sheetnames else wb.worksheets[0]
        hits = {}
        for row in ws.iter_rows(min_row=LINHA_INICIO, max_row=202, values_only=True):
            for ci, v in enumerate(row):
                if not v: continue
                s = str(v).strip()
                if s in ('None','','nan'): continue
                hits.setdefault(ci, [0, 0])[1] += 1
                if _PADRAO_LOTE.match(s):
                    hits[ci][0] += 1
        wb.close()
        best_ci, best_score, best_n = None, 0.0, 0
        for ci, (m, t) in hits.items():
            if t < 3: continue
            sc = m / t
            if sc >= 0.8 and sc > best_score:
                best_score, best_ci, best_n = sc, ci, m
        if best_ci is not None:
            return get_column_letter(best_ci + 1), best_n
    except Exception as e:
        log.debug(f"Detect coluna: {e}")
    return COLUNA_LOTES, 0


def ler_lotes(caminho: str) -> list[str]:
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string
    col, _ = detectar_coluna_lotes_fast(caminho)
    log.info(f"Coluna detectada: {col}")
    
    col_idx = column_index_from_string(col)
    lotes = []
    
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb[ABA_LOTES] if ABA_LOTES in wb.sheetnames else wb.worksheets[0]
        log.info(f"Aba: '{ws.title}'")
        
        for row in ws.iter_rows(min_row=LINHA_INICIO, min_col=col_idx, max_col=col_idx, values_only=True):
            v = row[0]
            if v is not None:
                s = str(v).strip()
                if s and s not in ("None", ""):
                    lotes.append(s)
                    
        log.info(f"{len(lotes)} lotes (col {col})")
        return lotes
    finally:
        wb.close()


def prefixo_ignorado(lote: str) -> str | None:
    return PREFIXOS_IGNORADOS.get(lote[:3].upper())


# ── Driver ─────────────────────────────────────────────────────────────────

def criar_driver(headless: bool) -> webdriver.Edge:
    opts = EdgeOptions()
    for arg in ["--ignore-certificate-errors", "--ignore-ssl-errors",
                "--log-level=3", "--disable-extensions",
                "--no-first-run", "--no-default-browser-check",
                "--disable-features=HttpsUpgrades"]:
        opts.add_argument(arg)
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--force-device-scale-factor=1")
        opts.add_argument("--run-all-compositor-stages-before-draw")
    d = webdriver.Edge(options=opts)
    if not headless:
        d.maximize_window()
    return d


# ── JavaScript ─────────────────────────────────────────────────────────────

_JS_FRAME_INPUT = """
var frames = document.querySelectorAll('iframe');
for (var i = 0; i < frames.length; i++) {
    try {
        var doc = frames[i].contentDocument || frames[i].contentWindow.document;
        if (doc && doc.getElementById(arguments[0])) return i;
    } catch(e) {}
}
return -1;
"""

_JS_SNIPER = """
var frames = document.querySelectorAll('iframe');
for (var i = 0; i < frames.length; i++) {
    try {
        var doc = frames[i].contentDocument || frames[i].contentWindow.document;
        var tds = doc.querySelectorAll('td');
        for (var j = 0; j < tds.length; j++) {
            if (tds[j].innerText.trim() === arguments[0]) {
                var tabela = tds[j].closest('table');
                if (tabela) return tabela.outerHTML;
            }
        }
    } catch(e) {}
}
return null;
"""

_JS_SMART_WAIT = """
var frames = document.querySelectorAll('iframe');
for (var i = 0; i < frames.length; i++) {
    try {
        var doc = frames[i].contentDocument || frames[i].contentWindow.document;
        if (doc && doc.getElementById(arguments[0])) return true;
    } catch(e) {}
}
return false;
"""


# ── SSRS sync ─────────────────────────────────────────────────────────────

def esperar_ssrs(driver, headless: bool = False):
    """
    Aguarda spinner aparecer e depois sumir.
    Em headless adiciona 1s de estabilização: o SSRS continua executando
    JS de pós-render por ~0.5–1s após o spinner desaparecer.
    """
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, ID_LOADING)))
    except Exception:
        pass
    try:
        WebDriverWait(driver, TIMEOUT_RENDER).until_not(
            EC.presence_of_element_located((By.ID, ID_LOADING)))
    except Exception:
        log.warning("Timeout SSRS — continuando")
    if headless:
        time.sleep(1.0)


# ── Limpeza e conversão de tipos ───────────────────────────────────────────

def limpar_df(df_bruto: pd.DataFrame, lote: str) -> pd.DataFrame | None:
    """
    Converte tabela bruta do SSRS em DataFrame estruturado.
    Encontra o header pelo match exato de 'Data/Hora'; descarta tudo acima.
    Aplica conversões de tipo em todas as colunas para eliminar avisos
    de 'Texto para coluna' no Excel.
    """
    header_idx = None
    for idx, row in df_bruto.iterrows():
        if (row.astype(str).str.strip() == MARCADOR_RELATORIO).any():
            header_idx = idx
            break
    if header_idx is None:
        return None

    headers = df_bruto.iloc[header_idx].astype(str).str.strip().tolist()
    df = df_bruto.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.loc[:, df.columns.notna()]
    df = df.loc[:, ~df.columns.str.contains(r'^Unnamed|^$', regex=True)]

    if MARCADOR_RELATORIO not in df.columns:
        return None

    mascara = df[MARCADOR_RELATORIO].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")
    df = df[mascara].copy()
    if df.empty:
        return None

    df.insert(0, "Lote_Base", lote)

    # Decimais: SSRS exporta sem separador (224200 → 2242.00)
    for col in COLUNAS_DIV100:
        if col in df.columns:
            df[col] = (df[col].astype(str)
                       .str.replace(".", "", regex=False)
                       .str.replace(",", ".", regex=False))
            df[col] = pd.to_numeric(df[col], errors="coerce").div(100).round(2)

    for col in COLUNAS_INT:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # COLUNAS_STR: garantir texto puro para evitar aviso de conversão no Excel
    for col in COLUNAS_STR:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": "", "None": ""})

    # Demais colunas de texto: limpar "nan"
    texto_cols = [c for c in df.columns
                  if c not in COLUNAS_NUM
                  and c not in (MARCADOR_RELATORIO, "Lote_Base")]
    for col in texto_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": "", "None": ""})

    return df


def html_para_df(html: str, lote: str) -> pd.DataFrame | None:
    try:
        tabelas = pd.read_html(StringIO(html), header=None, decimal=",", thousands=".")
        if not tabelas:
            return None
        return limpar_df(tabelas[0], lote)
    except Exception as e:
        log.warning(f"[{lote}] Erro no processamento: {e}")
        return None


def aguardar_iframe_aba(driver, wait, lote=""):
    for tentativa in range(1, 4):
        try:
            wait.until(lambda d: d.execute_script(_JS_SMART_WAIT, ID_INPUT))
            return True
        except Exception:
            if tentativa < 3:
                log.warning(f"[{lote}] Iframe travou na tentativa {tentativa}. Dando F5...")
                driver.refresh()
    return False

def injetar_lote(driver, wait, lote: str):
    driver.switch_to.default_content()
    frame_idx = driver.execute_script(_JS_FRAME_INPUT, ID_INPUT)
    if frame_idx == -1: return False
    
    driver.switch_to.frame(frame_idx)
    try:
        campo = wait.until(EC.element_to_be_clickable((By.ID, ID_INPUT)))
        campo.click()
        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(Keys.DELETE)
        campo.send_keys(lote)
        driver.find_element(By.ID, ID_BOTAO).click()
        driver.switch_to.default_content()
        return True
    except Exception:
        driver.switch_to.default_content()
        return False

def extrair_tabela(driver, lote: str, headless: bool):
    esperar_ssrs(driver, headless)
    html = None
    for _ in range(15):
        html = driver.execute_script(_JS_SNIPER, MARCADOR_RELATORIO)
        if html: break
        time.sleep(1)
    if not html: return None
    return html_para_df(html, lote)

def processar_lote_sequencial(driver, wait, lote, headless):
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        if not injetar_lote(driver, wait, lote):
            time.sleep(2)
            continue
        df = extrair_tabela(driver, lote, headless)
        if df is not None and not df.empty:
            return df
        log.warning(f"[{lote}] tentativa {tentativa}/{MAX_TENTATIVAS} - Vazio")
        time.sleep(2)
    return None

import queue

def extrair_tabela_segura(driver, handle, lote, lock, headless):
    t_start = time.time()
    
    while time.time() - t_start < 10:
        with lock:
            driver.switch_to.window(handle)
            try:
                driver.find_element(By.ID, ID_LOADING)
                break
            except Exception:
                pass
        time.sleep(0.2)
        
    while time.time() - t_start < TIMEOUT_RENDER:
        with lock:
            driver.switch_to.window(handle)
            try:
                driver.find_element(By.ID, ID_LOADING)
            except Exception:
                break
        time.sleep(0.5)
        
    html = None
    for _ in range(15):
        with lock:
            driver.switch_to.window(handle)
            html = driver.execute_script(_JS_SNIPER, MARCADOR_RELATORIO)
        if html: break
        time.sleep(1)
        
    if not html: return None
    return html_para_df(html, lote)

def get_qtd_abas():
    if "QTD_ABAS" in globals() and QTD_ABAS != 2:
        return QTD_ABAS
    return QTD_ABAS_POR_EDGE

def worker_aba_thread(aba_idx, handle, fila, driver, wait, edge_lock, global_lock, estado, acumulado, falhas, headless, t_worker, tag_aba):
    while not fila.empty():
        if estado.parar.is_set():
            break

        try:
            lote = fila.get_nowait()
        except queue.Empty:
            break

        log.info(f"[{tag_aba}] Iniciando lote: {lote}")

        # Injeção requer uso exclusivo do driver (edge_lock)
        with edge_lock:
            driver.switch_to.window(handle)
            sucesso = injetar_lote(driver, wait, lote)

        if not sucesso:
            log.warning(f"[{tag_aba}] {lote} falhou ao injetar. Retentando...")
            time.sleep(2)
            with edge_lock:
                driver.switch_to.window(handle)
                sucesso = injetar_lote(driver, wait, lote)

        df = None
        if sucesso:
            df = extrair_tabela_segura(driver, handle, lote, edge_lock, headless)

        # Atualização de estado global (global_lock)
        with global_lock:
            n_feitos = estado.atual + 1
            elapsed = time.time() - t_worker
            media = elapsed / n_feitos if n_feitos > 0 else 0
            restante = media * (estado.total - n_feitos) / 60

            ok = df is not None and not df.empty
            try:
                estado.tick(lote, ok, False, media, restante, elapsed / 60)
            except Exception:
                estado.tick(lote, ok, media, restante, elapsed / 60)

            if ok:
                acumulado.append(df)
                log.info(f"[{tag_aba}] {lote} finalizado | {len(df)} linhas")
            else:
                falhas.append(lote)
                log.warning(f"[{tag_aba}] {lote} falhou (sem dados).")

        fila.task_done()

def worker_edge_manager(edge_idx, fila, global_lock, estado, acumulado, falhas, headless, t_worker):
    tag_edge = f"EDGE {edge_idx+1}"
    log.info(f"[{tag_edge}] Iniciando novo navegador...")
    
    try:
        driver = criar_driver(headless)
    except Exception as e:
        log.error(f"[{tag_edge}] Falha ao criar driver: {e}")
        return

    wait = WebDriverWait(driver, 20)
    edge_lock = threading.Lock()
    handles = []

    try:
        log.info(f"[{tag_edge}] Abrindo {get_qtd_abas()} abas...")
        for i in range(get_qtd_abas()):
            if i == 0:
                driver.get(URL_RELATORIO)
            else:
                driver.execute_script(f"window.open('{URL_RELATORIO}', '_blank');")
                driver.switch_to.window(driver.window_handles[-1])
            handles.append(driver.window_handles[-1])

        log.info(f"[{tag_edge}] Aguardando carregamento do portal nas {get_qtd_abas()} abas...")
        for i in range(get_qtd_abas()):
            driver.switch_to.window(handles[i])
            tag_aba = f"EDGE {edge_idx+1} - ABA {i+1}"
            if not aguardar_iframe_aba(driver, wait, tag_aba):
                log.error(f"[{tag_aba}] Não carregou o portal.")

        # Lança as threads para cada aba deste Edge
        abas_threads = []
        for i in range(get_qtd_abas()):
            tag_aba = f"EDGE {edge_idx+1}-Aba{i+1}"
            t = threading.Thread(
                target=worker_aba_thread,
                args=(i, handles[i], fila, driver, wait, edge_lock, global_lock, estado, acumulado, falhas, headless, t_worker, tag_aba),
                name=tag_aba
            )
            abas_threads.append(t)
            t.start()

        for t in abas_threads:
            t.join()

    finally:
        log.info(f"[{tag_edge}] Encerrando navegador Edge...")
        try:
            driver.quit()
        except Exception:
            pass

def worker(lotes: list[str], headless: bool,
           estado: Estado, acumulado: list,
           falhas: list, lock: threading.Lock):
    t_worker = time.time()
    log.info(f"Iniciando {QTD_EDGES} navegadores com {get_qtd_abas()} abas cada (Total: {QTD_EDGES * get_qtd_abas()} extrações simultâneas)...")

    fila = queue.Queue()
    for lote in lotes:
        fila.put(lote)

    edge_threads = []
    for i in range(QTD_EDGES):
        t = threading.Thread(
            target=worker_edge_manager,
            args=(i, fila, lock, estado, acumulado, falhas, headless, t_worker),
            name=f"EdgeMgr-{i+1}"
        )
        edge_threads.append(t)
        t.start()

    for t in edge_threads:
        t.join()

# ── Formatação Excel ──
def formatar_excel(caminho: str) -> None:
    """
    Aplica formatação visual e corrige tipos de coluna para eliminar
    avisos 'Texto para coluna' do Excel.
    - Todas as colunas de texto recebem number_format='@' (força texto)
    - Números: '#,##0.00' ou '#,##0'
    - Zebra + alerta vermelho para divergência de peso
    """
    wb = load_workbook(caminho)
    ws = wb["Dados"]

    fnt_h = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    fnt_n = Font(name="Calibri", size=9)
    aln_c = Alignment(horizontal="center", vertical="center")
    aln_l = Alignment(horizontal="left",   vertical="center")
    aln_r = Alignment(horizontal="right",  vertical="center")
    thin  = Side(style="thin", color="BDD7EE")
    brd   = Border(left=thin, right=thin, top=thin, bottom=thin)

    col_names = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    for cell in ws[1]:
        cell.font      = fnt_h
        cell.fill      = PatternFill("solid", fgColor=XL_HEADER)
        cell.alignment = aln_c
        cell.border    = brd
    ws.row_dimensions[1].height = 22

    for r in range(2, ws.max_row + 1):
        fill_base = XL_ZEBRA if r % 2 == 0 else XL_NORMAL

        for c in range(1, ws.max_column + 1):
            cell     = ws.cell(r, c)
            col_name = col_names[c-1] if c <= len(col_names) else ""
            cell.border = brd
            cell.fill   = PatternFill("solid", fgColor=fill_base)
            cell.font   = fnt_n

            if col_name in COLUNAS_NUM:
                cell.alignment     = aln_r
                cell.number_format = ('#,##0' if col_name in set(COLUNAS_INT)
                                      else '#,##0.00')
            elif col_name == MARCADOR_RELATORIO:
                cell.alignment     = aln_c
                cell.number_format = '@'
            else:
                cell.alignment     = aln_l
                cell.number_format = '@'

        ws.row_dimensions[r].height = 15

    for c, col_name in enumerate(col_names, 1):
        max_w = len(str(col_name or "")) + 3
        for r in range(2, min(ws.max_row + 1, 202)):
            v = ws.cell(r, c).value
            if v:
                max_w = max(max_w, len(str(v)) + 2)
        ws.column_dimensions[get_column_letter(c)].width = min(max_w, 32)

    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho)
    log.info("  -> Excel formatado.")


# ── Salvar resultado ───────────────────────────────────────────────────────

def salvar_excel(resultados: list, ignorados: list,
                 falhas: list,
                 caminho: str) -> pd.DataFrame:
    """
    Abas geradas:
    - Dados       → todos os registros extraídos
    - Revisar     → lotes com divergência persistente (revisar manualmente)
    - Falhas      → lotes que falharam na extração
    - Ignorados   → lotes pulados por prefixo (DMX etc.)  # [REVIEW] Product category names kept
    """
    df_final = pd.concat(resultados, ignore_index=True)
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Dados", index=False)
        if falhas:
            pd.DataFrame({"Lote": falhas}).to_excel(
                writer, sheet_name="Falhas", index=False)
        if ignorados:
            pd.DataFrame(ignorados).to_excel(
                writer, sheet_name="Ignorados", index=False)
    formatar_excel(caminho)
    return df_final


# ── Fluxo principal ────────────────────────────────────────────────────────

def processar(lotes: list[str], saida: str, headless: bool) -> None:
    resultados:    list[pd.DataFrame] = []
    falhas:        list[str]          = []
    ignorados:     list[dict]         = []

    lotes_processar = []
    for lote in lotes:
        motivo = prefixo_ignorado(lote)
        if motivo:
            ignorados.append({"Lote": lote, "Motivo": motivo})
        else:
            lotes_processar.append(lote)

    total = len(lotes_processar)

    if ignorados:
        log.warning("=" * 55)
        log.warning(f"⚠️  {len(ignorados)} lote(s) ignorados:")
        for item in ignorados:
            log.warning(f"   {item['Lote']}  →  {item['Motivo']}")
        log.warning("=" * 55)

    if total == 0:
        log.error("Nenhum lote válido.")
        return

    estado   = Estado(total=total)
    lock     = threading.Lock()
    t_inicio = time.time()

    def run_worker():
        worker(lotes_processar, headless, estado,
               resultados, falhas, lock)
        estado.concluido = True

    t = threading.Thread(target=run_worker, daemon=True, name="worker-ssrs")
    t.start()
    JanelaProgresso(estado).run()
    t.join(timeout=10)

    elapsed       = time.time() - t_inicio
    sufixo        = "_PARCIAL" if estado.parar.is_set() else ""
    caminho_final = saida.replace(".xlsx", f"{sufixo}.xlsx")

    log.info("=" * 55)
    log.info(f"Total            : {len(lotes)}")
    log.info(f"Ignorados        : {len(ignorados)}")
    log.info(f"Processados      : {total}")
    log.info(f"  Sucesso        : {len(resultados)}")
    log.info(f"  Falhas         : {len(falhas)}")
    if falhas:
        log.warning(f"Lotes com erro: {falhas}")

    if not resultados:
        log.error("Nenhum dado capturado.")
        notificar_windows("Lot Miner — Erro", "Nenhum dado foi capturado.")
        return

    df_final = salvar_excel(resultados, ignorados, falhas,
                            caminho_final)

    log.info("=" * 55)
    log.info(f"Concluído em {elapsed/60:.1f} min  ({elapsed/total:.1f}s/lote)")
    log.info(f"Total de linhas : {len(df_final)}")
    log.info(f"Arquivo salvo   : {caminho_final}")
    abas = ["Dados"]
    if falhas:        abas.append("Falhas")
    if ignorados:     abas.append("Ignorados")
    log.info(f"Abas geradas    : {' | '.join(abas)}")

    status = "Parcial" if estado.parar.is_set() else "Concluida"
    msg = (f"{len(resultados)}/{total} lotes extraídos | "
           f"{len(df_final)} linhas | {elapsed/60:.1f} min")

    notificar_windows(
        f"Lot Miner — {status}",
        msg,
        arquivo=caminho_final
    )


# ── Entry point ────────────────────────────────────────────────────────────

def executar():
    janela = JanelaPrincipal()
    if not janela.abrir():
        return

    if janela.extrair_sap:
        try:
            log.info("Iniciando extração direta da MB52 pelo SAP...")
            log.info(f"Pasta SAP: {janela.sap_saida_dir}")
            log.info(f"Categoria Professional: {janela.categoria_professional}")
            log.info(f"Categoria Higiênicos: {janela.categoria_higienicos}")

            janela.entrada = extrair_mb52_bobinas_sap(
                janela.sap_saida_dir,
                janela.categoria_professional,
                janela.categoria_higienicos
            )
            log.info(f"Entrada do Lot Miner definida pelo SAP: {janela.entrada}")
        except Exception as e:
            log.exception("Falha na extração direta da MB52")
            from tkinter import messagebox
            root_erro = tk.Tk()
            root_erro.withdraw()
            messagebox.showerror(
                "Lot Miner — Erro no SAP",
                f"Não foi possível extrair a MB52 diretamente do SAP.\n\n{e}"
            )
            root_erro.destroy()
            return

    try:
        lotes = ler_lotes(janela.entrada)
    except Exception as e:
        log.exception("Falha ao ler a planilha de entrada")
        import tkinter.messagebox as messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Lot Miner — Erro na planilha",
            f"Não foi possível ler a planilha de entrada.\n\n{e}"
        )
        root.destroy()
        return

    log.info(
        f"{len(lotes)} lotes | headless={janela.headless} | "
        f"origem={'SAP' if janela.extrair_sap else 'arquivo manual'}"
    )

    processar(lotes, janela.saida, janela.headless)


def perguntar_executar_novamente(parent=None) -> bool:
    """
    Pergunta ao usuário se deseja executar o Lot Miner novamente.

    Retorna:
        True  -> usuário clicou em Sim
        False -> usuário clicou em Não
    """

    from tkinter import messagebox

    # Ocultar a janela principal do Tk para a caixa de diálogo se parent for None
    root = None
    if parent is None:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()

    resposta = messagebox.askyesno(
        title="Lot Miner",
        message=(
            "A extração foi concluída.\n\n"
            "Deseja executar o Lot Miner novamente?"
        ),
        icon="question",
        parent=parent
    )

    if root:
        root.destroy()

    return resposta


if __name__ == "__main__":
    while True:
        executar()
        if not perguntar_executar_novamente():
            break

