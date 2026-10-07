"""
plan_explorer_v3.py
Automação de Apontamentos de Produção — ACME Corp
MB51 → Excel COM → COR3 → Base Ordens → Tabela Dinâmica

Desenvolvido por: Author

CORREÇÕES v6:
  - Threading: UI nunca trava, processo roda em background
  - Busca automática da planilha mais recente (sem diálogo)
  - Coluna F = Ordem (corrigido de C)
  - Coluna K = Linha (flag "NÃO ENCONTRADO")
  - Tela de confirmação antes de executar

BLINDAGENS v6.1:
  - Pré-configuração de mês/ano na interface (virada de mês)
  - Range de datas inteligente: dia 01 usa mês anterior, dia 02+ usa mês atual
  - Criação automática de pasta do dia de destino se não existir
  - Verificação de arquivo em uso na rede antes de processar
  - Reset automático do cache gen_py do win32com em caso de falha COM
  - Eliminação de instâncias zumbi do Excel no bloco finally
  - Timeout no aguardo de abertura do SAP Logon
"""

import os
import sys
import re
import json
import time
import html
import shutil
import tempfile
import threading
import subprocess
import traceback
import calendar
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta
import win32com.client as win32
import win32clipboard
import pythoncom
from pathlib import Path

# ─────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO PERSISTENTE
# ─────────────────────────────────────────────────────────────────
_CONFIG_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "plan_explorer_config.json"
)

def _carregar_config():
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _salvar_config(dados: dict):
    try:
        cfg = _carregar_config()
        cfg.update(dados)
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────
# PALETA
# ─────────────────────────────────────────────────────────────────
BG_ROOT      = "#0f0f0f"
BG_HEADER    = "#141414"
BG_CARD      = "#1a1a1a"
BG_FOOTER    = "#111111"
BG_INFO      = "#151f18"   # card de info da planilha detectada
ACCENT       = "#00c875"
ACCENT_DIM   = "#007a47"
TEXT_PRIMARY = "#f0f0f0"
TEXT_MUTED   = "#6b6b6b"
TEXT_WARN    = "#f5a623"
TEXT_ERROR   = "#e05252"
TEXT_LOG     = "#b8f0cc"
TEXT_INFO    = "#7ecfa0"

FNT_TITLE    = ("Segoe UI", 13, "bold")
FNT_SUB      = ("Segoe UI", 9)
FNT_LOG      = ("Cascadia Code", 9)
FNT_STATUS   = ("Segoe UI", 8)
FNT_LABEL    = ("Segoe UI", 8, "bold")
FNT_BTN      = ("Segoe UI", 10, "bold")
FNT_INFO     = ("Segoe UI", 9)
FNT_INFO_B   = ("Segoe UI", 9, "bold")

# ─────────────────────────────────────────────────────────────────
# CONSTANTES DE NEGÓCIO
# ─────────────────────────────────────────────────────────────────
def _resolver_raiz_producao(ano: int) -> str:
    """
    Resolve o caminho base da pasta de produção de forma segura para o ano solicitado.
    Procura pela pasta mestre (ex: '02 - Produção' ou subpastas de meses).
    Utiliza o diretório-base de produção configurado pelo usuário na aplicação.
    """
    cfg = _carregar_config()
    cfg_base = cfg.get(f"raiz_producao_{ano}", "") or cfg.get("diretorio_base", "")

    if cfg_base and os.path.exists(cfg_base):
        # 1. Se o diretório configurado já for diretamente a pasta mestre (contém subpastas de mês)
        try:
            subs = [f for f in os.listdir(cfg_base) if os.path.isdir(os.path.join(cfg_base, f))]
            if any(s.startswith(("01 -", "02 -", "03 -", "04 -", "05 -", "06 -", "07 -", "08 -", "09 -", "10 -", "11 -", "12 -")) for s in subs):
                return cfg_base
        except Exception:
            pass

        # 2. Se o diretório configurado for a raiz, busca pela estrutura padrão:
        #    <diretório-base>\Controle de Produção\Apontamentos Diários\{ano}
        cand_anos = [
            os.path.join(cfg_base, "Controle de Produção", "Apontamentos Diários", str(ano)),
            os.path.join(cfg_base, str(ano)),
            cfg_base
        ]

        for caminho_ano in cand_anos:
            if not os.path.exists(caminho_ano):
                continue
            try:
                subpastas = [f for f in os.listdir(caminho_ano) if os.path.isdir(os.path.join(caminho_ano, f))]
                # Prioriza pasta '02 - Produção'
                if "02 - Produção" in subpastas:
                    return os.path.join(caminho_ano, "02 - Produção")
                # Heurística: procura por pastas contendo 'produ'
                cand = [f for f in subpastas if "produ" in f.lower()]
                if cand:
                    return os.path.join(caminho_ano, cand[0])
                # Se contém subpastas mensais diretamente
                if any(s.startswith(("01 -", "02 -", "03 -", "04 -", "05 -", "06 -", "07 -", "08 -", "09 -", "10 -", "11 -", "12 -")) for s in subpastas):
                    return caminho_ano
            except Exception:
                continue

    # Caso não seja configurado ou não seja encontrado
    base_str = cfg_base if cfg_base else "<diretório-base não configurado>"
    raise FileNotFoundError(
        f"ERRO na detecção: Não foi possível localizar a pasta mestre do Prod Forge para {ano}.\n\n"
        f"Estrutura esperada:\n"
        f"{base_str}\\Controle de Produção\\Apontamentos Diários\\{ano}\n\n"
        f"Possíveis causas:\n"
        f"- O diretório-base não foi configurado corretamente nas configurações da aplicação.\n"
        f"- A pasta mestre foi renomeada ou movida.\n"
        f"- A estrutura de diretórios foi alterada.\n"
        f"- A unidade de rede ou pasta local está indisponível ou sem permissão de acesso."
    )

DEPARA_LINHAS = {
    "B-CONVAM": "PH.04",
    "B-CONVBR": "PH.06",
    "B-CONVCR": "PH.07",
    "B-CONVCB": "FOR.01",
    "B-CONVC2": "FOR.02",
    "B-CONVP7": "DMX.05",
    "B-CONVTA": "PROF.01",
    "B-CONVL4": "PROF.02",
    "B-CONVL5": "PH.08",
    "B-CONVL6": "PH.12",
    "B-CONVL7": "PH.13",
    "B-CONVL8": "PH.14",
    "B-CONV12": "PH.15",
    "B-CONV13": "FOR.03",
    "B-CONV14": "PROF.03",
    "B-CONV15": "PROF.04",
}

MESES = {
    1:"01 - Janeiro", 2:"02 - Fevereiro", 3:"03 - Março",
    4:"04 - Abril",   5:"05 - Maio",      6:"06 - Junho",
    7:"07 - Julho",   8:"08 - Agosto",    9:"09 - Setembro",
    10:"10 - Outubro",11:"11 - Novembro", 12:"12 - Dezembro",
}

# Variáveis globais de configuração de mês/ano (ajustadas pela UI)
_mes_selecionado  = datetime.today().month
_ano_selecionado  = datetime.today().year

ETAPAS = [
    "Conexão SAP",
    "Extração MB51",
    "Tratamento Excel",
    "Verificação COR3",
    "Base Ordens",
    "Verificação MM03",
    "Tabela Dinâmica",
    "Salvamento",
]

# ─────────────────────────────────────────────────────────────────
# BUSCA AUTOMÁTICA DA PLANILHA MAIS RECENTE
# ─────────────────────────────────────────────────────────────────
def _pasta_dia_para_data(nome, ano_ref=None):
    """Converte nome de pasta DD.MM em datetime. Aceita sufixos opcionais."""
    m = re.match(r"^(\d{2})\.(\d{2})", nome)
    if not m:
        return None
    dia, mes = int(m.group(1)), int(m.group(2))
    ano = ano_ref if ano_ref else datetime.today().year
    try:
        return datetime(ano, mes, dia)
    except ValueError:
        return None


def _checar_arquivo_em_uso(caminho: str) -> bool:
    """Retorna True se o arquivo estiver aberto/bloqueado por outro processo."""
    if not os.path.exists(caminho):
        return False
    try:
        os.rename(caminho, caminho)
        return False
    except OSError:
        return True


def _resetar_cache_com():
    """Apaga o cache gen_py do win32com para forçar regeneração limpa."""
    import tempfile
    gen_path = os.path.join(tempfile.gettempdir(), "gen_py")
    if os.path.isdir(gen_path):
        try:
            shutil.rmtree(gen_path)
            return True
        except Exception:
            return False
    return False


def _ensure_excel_dispatch():
    """
    Tenta criar o objeto Excel via EnsureDispatch (Early Binding).
    Em caso de falha por cache corrompido, reseta o gen_py e tenta novamente.
    """
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        registrar_log("Cache COM corrompido detectado. Resetando gen_py e tentando novamente…", "aviso")
        _resetar_cache_com()
        try:
            return win32.gencache.EnsureDispatch("Excel.Application")
        except Exception as e:
            raise Exception(f"Falha crítica ao conectar ao Excel via COM após reset: {e}")


DATA_MINIMA_PRODUCAO = datetime(2026, 6, 1)

def calcular_range_datas(mes: int, ano: int):
    """
    Calcula o range de datas (inicio, fim) para a extração MB51.
    Limite mínimo dos dados: 01/06/2026.

    Regra de negócio:
      - Se hoje é dia 01 do mês selecionado → usa o mês ANTERIOR completo
        (ex: 01/08 → range de 01/07 a 01/08)
      - Caso contrário → usa do dia 01 do mês selecionado até hoje
        (ex: 13/08 → range de 01/08 a 13/08)

    Retorna (data_inicio: datetime, data_fim: datetime, mes_base: int, ano_base: int)
    onde mes_base/ano_base indicam o mês de onde buscar a planilha.
    """
    hoje = datetime.today()
    data_ref = datetime(ano, mes, hoje.day if (ano == hoje.year and mes == hoje.month) else
                        calendar.monthrange(ano, mes)[1])

    # Ajuste: se o dia de referência é o dia 1 do mês selecionado,
    # o range vai do primeiro dia do mês anterior até hoje
    if data_ref.day == 1 and data_ref.month == mes and data_ref.year == ano:
        # Mês anterior
        primeiro_mes_ant = (data_ref.replace(day=1) - timedelta(days=1)).replace(day=1)
        inicio = primeiro_mes_ant
        mes_base = primeiro_mes_ant.month
        ano_base = primeiro_mes_ant.year
    else:
        inicio = data_ref.replace(day=1)
        mes_base = mes
        ano_base = ano

    # Aplica a trava de segurança: data inicial nunca antes de 01/06/2026
    if inicio < DATA_MINIMA_PRODUCAO:
        inicio = DATA_MINIMA_PRODUCAO

    return inicio, data_ref, mes_base, ano_base


def localizar_planilha_base(mes_ref: int = None, ano_ref: int = None):
    """
    Localiza a planilha base do dia anterior mais recente.
    Usa mes_ref/ano_ref para determinar em qual pasta mensal buscar.
    Cria a pasta do dia de destino se ela não existir.
    """
    hoje = datetime.today()
    if mes_ref is None:
        mes_ref = _mes_selecionado
    if ano_ref is None:
        ano_ref = _ano_selecionado

    # Calcula o range e descobre em qual mês a planilha base está
    data_inicio, data_fim, mes_base, ano_base = calcular_range_datas(mes_ref, ano_ref)

    # Valida se o mês solicitado é anterior ao limite operacional (01/06/2026)
    if datetime(ano_base, mes_base, 1) < DATA_MINIMA_PRODUCAO:
        raise FileNotFoundError(
            "Dados de produção disponíveis somente a partir de Junho/2026 (01/06/2026)."
        )

    nome_mes = MESES.get(mes_base)
    if not nome_mes:
        raise FileNotFoundError(f"Mês inválido: {mes_base}")

    try:
        raiz_producao = _resolver_raiz_producao(ano_base)
    except PermissionError as e:
        # Repassa o erro de rede claro para a interface
        raise FileNotFoundError(str(e))

    pasta_mes = os.path.join(raiz_producao, nome_mes)

    # Cria a pasta do mês caso não exista (útil em virada de mês)
    if not os.path.isdir(pasta_mes):
        try:
            os.makedirs(pasta_mes, exist_ok=True)
        except Exception as e:
            raise FileNotFoundError(
                f"Pasta do mês não encontrada e não há permissão para criá-la:\n{pasta_mes}\n\n"
                f"Erro do sistema: {e}"
            )

    # Lista subpastas com padrão DD.MM
    subpastas = []
    for item in os.scandir(pasta_mes):
        if not item.is_dir():
            continue
        dt = _pasta_dia_para_data(item.name, ano_ref=ano_base)
        # Só considera pastas anteriores à data_fim
        if dt and dt.date() < data_fim.date():
            subpastas.append((dt, item.path, item.name))

    if not subpastas:
        raise FileNotFoundError(
            f"Nenhuma pasta de dia anterior encontrada em:\n{pasta_mes}\n"
            f"(Buscando dias antes de {data_fim.strftime('%d/%m/%Y')})"
        )

    # Pega a pasta mais recente
    subpastas.sort(key=lambda x: x[0], reverse=True)
    _, pasta_dia_orig, dia_ref = subpastas[0]

    # Localiza a planilha principal
    candidatos = [
        f for f in os.listdir(pasta_dia_orig)
        if f.lower().endswith((".xlsx", ".xlsm"))
        and not f.startswith("~$")
        and not f.upper().startswith("EXPORT_")
    ]
    if not candidatos:
        raise FileNotFoundError(f"Nenhuma planilha válida encontrada em:\n{pasta_dia_orig}")

    candidatos.sort(key=len, reverse=True)
    arquivo_base = os.path.join(pasta_dia_orig, candidatos[0])

    return {
        "arquivo_base":   arquivo_base,
        "pasta_mes":      pasta_mes,
        "pasta_dia_orig": pasta_dia_orig,
        "dia_ref":        dia_ref,
        "data_inicio":    data_inicio,
        "data_fim":       data_fim,
        "mes_base":       mes_base,
        "ano_base":       ano_base,
    }


def gerar_caminho_destino(info):
    """
    Gera o caminho de destino e nome do arquivo com total coerência.
    Se for o mês atual: 'Produção - DD.MM.xlsx' dentro de '08 - Agosto\\DD.MM'
    Se for um mês histórico (ex: Janeiro): 'Produção Janeiro - DD.MM.xlsx' dentro de '01 - Janeiro\\DD.MM'
    """
    hoje = datetime.today()
    dia_s = hoje.strftime("%d.%m")
    mes_base = info.get("mes_base", hoje.month)
    ano_base = info.get("ano_base", hoje.year)
    pasta_mes_base = info["pasta_mes"]

    raw_mes = MESES.get(mes_base, "")
    nome_mes_str = raw_mes.split("-")[-1].strip() if "-" in raw_mes else raw_mes

    # A pasta de destino SEMPRE fica dentro da pasta do mês da base selecionada
    pasta_destino = os.path.join(pasta_mes_base, dia_s)

    # Nomenclatura coerente
    if mes_base == hoje.month and ano_base == hoje.year:
        nome_novo = f"Produção - {dia_s}.xlsx"
    else:
        nome_novo = f"Produção {nome_mes_str} - {dia_s}.xlsx"

    arquivo_novo = os.path.join(pasta_destino, nome_novo)
    nome_pasta_mes = os.path.basename(pasta_mes_base)
    rel_destino = os.path.join(nome_pasta_mes, dia_s, nome_novo)

    return {
        "pasta_destino": pasta_destino,
        "nome_novo":     nome_novo,
        "arquivo_novo":  arquivo_novo,
        "rel_destino":   rel_destino,
    }


# ══════════════════════════════════════════════════════════════════
#  JANELA PRINCIPAL
# ══════════════════════════════════════════════════════════════════
root = tk.Tk()
root.title("Automação de Produção — ACME Corp")
root.geometry("1100x760")
root.resizable(True, True)
root.configure(bg=BG_ROOT)
root.attributes("-topmost", True)

root.update_idletasks()
sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
root.geometry(f"1100x760+{(sw-1100)//2}+{(sh-760)//2}")

# ── CABEÇALHO ────────────────────────────────────────────────────
header = tk.Frame(root, bg=BG_HEADER, height=68)
header.pack(fill=tk.X)
header.pack_propagate(False)

tk.Frame(header, bg=ACCENT, width=4).pack(side=tk.LEFT, fill=tk.Y)

tk.Label(header, text="AUTOMAÇÃO DE PRODUÇÃO", font=FNT_TITLE,
         fg=TEXT_PRIMARY, bg=BG_HEADER, padx=16).pack(side=tk.LEFT, pady=14)

tk.Label(header, text="MB51  ›  COR3  ›  Base Ordens  ›  Tabela Dinâmica",
         font=FNT_SUB, fg=TEXT_MUTED, bg=BG_HEADER).pack(side=tk.LEFT, pady=14)

lbl_hora = tk.Label(header, text="", font=FNT_STATUS, fg=TEXT_MUTED,
                    bg=BG_HEADER, padx=16)
lbl_hora.pack(side=tk.RIGHT, pady=14)

def _tick():
    lbl_hora.config(text=datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))
    root.after(1000, _tick)
_tick()

# ── CARD DE PRÉ-CONFIGURAÇÃO DE MÊS ─────────────────────────────

# Estilo do combobox definido ANTES da criação dos widgets
style = ttk.Style()
style.theme_use("clam")
style.configure("TCombobox",
    fieldbackground="#1e2a35", background="#1e2a35",
    foreground=TEXT_PRIMARY, selectbackground="#3a7bd5",
    selectforeground=TEXT_PRIMARY, bordercolor="#3a7bd5",
    arrowcolor=TEXT_PRIMARY,
)
style.map("TCombobox",
    fieldbackground=[("readonly", "#1e2a35")],
    foreground=[("readonly", TEXT_PRIMARY)],
)

frame_cfg = tk.Frame(root, bg="#12181e", pady=0)
frame_cfg.pack(fill=tk.X, padx=20, pady=(10, 0))

tk.Frame(frame_cfg, bg="#3a7bd5", height=2).pack(fill=tk.X)

cfg_inner = tk.Frame(frame_cfg, bg="#12181e", padx=14, pady=6)
cfg_inner.pack(fill=tk.X)

# ── Linha 1: seletores ─────────────────────────────────────────
cfg_row1 = tk.Frame(cfg_inner, bg="#12181e")
cfg_row1.pack(fill=tk.X)

tk.Label(cfg_row1, text="CONFIGURAÇÃO DO PERÍODO", font=FNT_LABEL,
         fg="#3a7bd5", bg="#12181e").pack(side=tk.LEFT)

tk.Label(cfg_row1, text="Mês:", font=FNT_INFO, fg=TEXT_MUTED,
         bg="#12181e").pack(side=tk.LEFT, padx=(20, 4))

def _obter_meses_validos(ano: int):
    """Retorna os meses disponíveis. Para 2026, apenas de Junho (6) a Dezembro (12)."""
    if ano <= 2026:
        return [MESES[i] for i in range(6, 13)]
    else:
        return [MESES[i] for i in range(1, 13)]

_meses_iniciais = _obter_meses_validos(_ano_selecionado)
_var_mes = tk.StringVar()
cmb_mes = ttk.Combobox(
    cfg_row1, textvariable=_var_mes,
    values=_meses_iniciais, state="readonly", width=16,
    font=FNT_INFO,
)
cmb_mes.pack(side=tk.LEFT, padx=(0, 16))

# Seleciona o mês atual se válido, senão o primeiro disponível (Junho)
_nome_atual = MESES.get(_mes_selecionado, "")
if _nome_atual in _meses_iniciais:
    cmb_mes.set(_nome_atual)
else:
    cmb_mes.set(_meses_iniciais[0])
    # Atualiza a global para o primeiro mês válido se o atual for inválido (ex: Maio -> Junho)
    for _num, _nome in MESES.items():
        if _nome == _meses_iniciais[0]:
            _mes_selecionado = _num
            break

tk.Label(cfg_row1, text="Ano:", font=FNT_INFO, fg=TEXT_MUTED,
         bg="#12181e").pack(side=tk.LEFT, padx=(0, 4))

_anos_disponiveis = [str(y) for y in range(2026, datetime.today().year + 2)]
_var_ano = tk.StringVar()
cmb_ano = ttk.Combobox(
    cfg_row1, textvariable=_var_ano,
    values=_anos_disponiveis, state="readonly", width=6,
    font=FNT_INFO,
)
cmb_ano.pack(side=tk.LEFT, padx=(0, 0))

try:
    _idx_ano = _anos_disponiveis.index(str(_ano_selecionado))
except ValueError:
    _idx_ano = 0
cmb_ano.current(_idx_ano)

# ── Linha 2: range calculado ───────────────────────────────────
cfg_row2 = tk.Frame(cfg_inner, bg="#12181e")
cfg_row2.pack(fill=tk.X, pady=(4, 0))

lbl_range_datas = tk.Label(
    cfg_row2, text="Calculando range...", font=FNT_INFO,
    fg=TEXT_INFO, bg="#12181e", anchor="w"
)
lbl_range_datas.pack(side=tk.LEFT, fill=tk.X)

def _atualizar_periodo_display(*_):
    """Atualiza globais e label de range ao mudar mês/ano na UI pelo usuário."""
    global _mes_selecionado, _ano_selecionado
    nome_sel = _var_mes.get()
    for num, nome in MESES.items():
        if nome == nome_sel:
            _mes_selecionado = num
            break
    try:
        _ano_selecionado = int(_var_ano.get())
    except ValueError:
        pass
    try:
        ini, fim, _, _ = calcular_range_datas(_mes_selecionado, _ano_selecionado)
        lbl_range_datas.config(
            text=f"▶  Range MB51: {ini.strftime('%d/%m/%Y')} → {fim.strftime('%d/%m/%Y')}"
        )
    except Exception:
        lbl_range_datas.config(text="")
    # Re-detecta planilha base com novo mês (somente quando usuário muda)
    threading.Thread(target=detectar_planilha, daemon=True).start()

def _ao_mudar_ano(*_):
    """Atualiza as opções do combobox de mês quando o ano é alterado."""
    try:
        ano = int(_var_ano.get())
    except ValueError:
        ano = 2026
    novos_meses = _obter_meses_validos(ano)
    cmb_mes['values'] = novos_meses
    if _var_mes.get() not in novos_meses:
        cmb_mes.set(novos_meses[0])
    _atualizar_periodo_display()

# Usa evento de seleção real do combobox (não trace) — evita disparo no startup
cmb_mes.bind("<<ComboboxSelected>>", _atualizar_periodo_display)
cmb_ano.bind("<<ComboboxSelected>>", _ao_mudar_ano)

# Popula o label de range imediatamente no startup com os valores padrão
try:
    _ini_startup, _fim_startup, _, _ = calcular_range_datas(_mes_selecionado, _ano_selecionado)
    lbl_range_datas.config(
        text=f"▶  Range MB51: {_ini_startup.strftime('%d/%m/%Y')} → {_fim_startup.strftime('%d/%m/%Y')}"
    )
except Exception:
    pass


# ── CARD DE CONFIGURAÇÃO DE DIRETÓRIOS ───────────────────────────
frame_export = tk.Frame(root, bg="#12181e", pady=0)
frame_export.pack(fill=tk.X, padx=20, pady=(8, 0))

tk.Frame(frame_export, bg="#e0a030", height=2).pack(fill=tk.X)

export_inner = tk.Frame(frame_export, bg="#12181e", padx=14, pady=6)
export_inner.pack(fill=tk.X)

# Rownt 1: Diretório-Base da Produção
_cfg_startup = _carregar_config()
_pasta_base_startup = _cfg_startup.get("diretorio_base", "") or _cfg_startup.get(f"raiz_producao_{_ano_selecionado}", "")

base_row1 = tk.Frame(export_inner, bg="#12181e")
base_row1.pack(fill=tk.X)

tk.Label(base_row1, text="DIRETÓRIO-BASE DA PRODUÇÃO (REDE / LOCAL)", font=FNT_LABEL,
         fg="#e0a030", bg="#12181e").pack(side=tk.LEFT)

btn_browse_base = tk.Button(
    base_row1, text="📁 Alterar Base", font=("Segoe UI", 8),
    fg=TEXT_PRIMARY, bg="#1e2a35", activebackground="#2a3a50",
    activeforeground=TEXT_PRIMARY, relief=tk.FLAT, padx=8, pady=2,
    cursor="hand2",
)
btn_browse_base.pack(side=tk.RIGHT, padx=(8, 0))

lbl_pasta_base = tk.Label(
    export_inner, text=_pasta_base_startup if _pasta_base_startup else "Nenhum diretório-base configurado — clique em 'Alterar Base'",
    font=FNT_INFO, fg=TEXT_MUTED if _pasta_base_startup else TEXT_WARN,
    bg="#12181e", anchor="w",
)
lbl_pasta_base.pack(fill=tk.X, pady=(2, 6))

def _selecionar_diretorio_base():
    atual = lbl_pasta_base.cget("text")
    initialdir = atual if atual and os.path.isdir(atual) else os.environ.get("USERPROFILE", "C:\\")
    p = filedialog.askdirectory(
        title="Selecione o diretório-base da produção (Rede ou Pasta Local)",
        initialdir=initialdir,
    )
    if p:
        pasta_norm = os.path.normpath(p)
        _salvar_config({"diretorio_base": pasta_norm, f"raiz_producao_{_ano_selecionado}": pasta_norm})
        lbl_pasta_base.config(text=pasta_norm, fg=TEXT_MUTED)
        registrar_log(f"Diretório-base alterado: {pasta_norm}", "destaque")
        threading.Thread(target=detectar_planilha, daemon=True).start()

btn_browse_base.config(command=_selecionar_diretorio_base)

# Row 2: Pasta de Exportação SAP
export_row1 = tk.Frame(export_inner, bg="#12181e")
export_row1.pack(fill=tk.X, pady=(4, 0))

tk.Label(export_row1, text="PASTA DE EXPORTAÇÃO SAP", font=FNT_LABEL,
         fg="#e0a030", bg="#12181e").pack(side=tk.LEFT)

btn_browse_export = tk.Button(
    export_row1, text="📁 Alterar SAP", font=("Segoe UI", 8),
    fg=TEXT_PRIMARY, bg="#1e2a35", activebackground="#2a3a50",
    activeforeground=TEXT_PRIMARY, relief=tk.FLAT, padx=8, pady=2,
    cursor="hand2",
)
btn_browse_export.pack(side=tk.RIGHT, padx=(8, 0))

_pasta_export_startup = _cfg_startup.get("pasta_exportacao_sap", "")

lbl_pasta_export = tk.Label(
    export_inner, text=_pasta_export_startup if _pasta_export_startup else "Nenhuma pasta selecionada — clique em 'Alterar SAP'",
    font=FNT_INFO, fg=TEXT_MUTED if _pasta_export_startup else TEXT_WARN,
    bg="#12181e", anchor="w",
)
lbl_pasta_export.pack(fill=tk.X, pady=(2, 0))

def _selecionar_pasta_export():
    """Abre diálogo para selecionar pasta de exportação e atualiza a interface."""
    atual = lbl_pasta_export.cget("text")
    initialdir = atual if atual and os.path.isdir(atual) else os.environ.get("USERPROFILE", "C:\\")
    p = filedialog.askdirectory(
        title="Selecione a pasta de saída do PROD.xlsx (exportação SAP)",
        initialdir=initialdir,
    )
    if p:
        pasta_norm = os.path.normpath(p)
        _salvar_config({"pasta_exportacao_sap": pasta_norm})
        lbl_pasta_export.config(text=pasta_norm, fg=TEXT_MUTED)
        registrar_log(f"Pasta de exportação alterada: {pasta_norm}", "destaque")

btn_browse_export.config(command=_selecionar_pasta_export)


# ── CARD DE DETECÇÃO DA PLANILHA ─────────────────────────────────
frame_info = tk.Frame(root, bg=BG_INFO, pady=0)
frame_info.pack(fill=tk.X, padx=20, pady=(8, 0))

# borda superior verde
tk.Frame(frame_info, bg=ACCENT, height=2).pack(fill=tk.X)

info_inner = tk.Frame(frame_info, bg=BG_INFO, padx=14, pady=10)
info_inner.pack(fill=tk.X)

# linha 1 — rótulo + nome do arquivo
row1 = tk.Frame(info_inner, bg=BG_INFO)
row1.pack(fill=tk.X)

tk.Label(row1, text="PLANILHA BASE DETECTADA", font=FNT_LABEL,
         fg=ACCENT, bg=BG_INFO).pack(side=tk.LEFT)

lbl_nome_arquivo = tk.Label(row1, text="Aguardando…", font=FNT_INFO_B,
                             fg=TEXT_PRIMARY, bg=BG_INFO)
lbl_nome_arquivo.pack(side=tk.LEFT, padx=(10, 0))

# linha 2 — caminho completo
lbl_caminho = tk.Label(info_inner, text="", font=FNT_INFO,
                        fg=TEXT_MUTED, bg=BG_INFO, anchor="w")
lbl_caminho.pack(fill=tk.X, pady=(2, 0))

# linha 3 — referência + destino
row3 = tk.Frame(info_inner, bg=BG_INFO)
row3.pack(fill=tk.X, pady=(4, 0))

lbl_ref   = tk.Label(row3, text="", font=FNT_INFO, fg=TEXT_INFO, bg=BG_INFO)
lbl_ref.pack(side=tk.LEFT)

lbl_arrow = tk.Label(row3, text="  →  ", font=FNT_INFO, fg=TEXT_MUTED, bg=BG_INFO)
lbl_arrow.pack(side=tk.LEFT)

lbl_dest  = tk.Label(row3, text="", font=FNT_INFO_B, fg=ACCENT, bg=BG_INFO)
lbl_dest.pack(side=tk.LEFT)

def atualizar_card_info(info):
    dest = gerar_caminho_destino(info)
    lbl_nome_arquivo.config(text=os.path.basename(info["arquivo_base"]), fg=TEXT_PRIMARY)
    lbl_caminho.config(text=info["arquivo_base"])
    lbl_ref.config(text=f"Base: {info['dia_ref']} ({info['data_inicio'].strftime('%d/%m')} → {info['data_fim'].strftime('%d/%m/%Y')})")
    lbl_dest.config(text=f"Salvará em: {dest['rel_destino']}")
    lbl_arrow.config(text="  →  ")

def card_erro(msg):
    lbl_nome_arquivo.config(text="NÃO ENCONTRADO", fg=TEXT_ERROR)
    lbl_caminho.config(text=msg)
    lbl_ref.config(text="")
    lbl_arrow.config(text="")
    lbl_dest.config(text="")

# ── PIPELINE DE ETAPAS ───────────────────────────────────────────
frame_pipe = tk.Frame(root, bg=BG_ROOT, pady=8)
frame_pipe.pack(fill=tk.X, padx=20)

step_widgets = {}

for i, nome in enumerate(ETAPAS):
    col = tk.Frame(frame_pipe, bg=BG_ROOT)
    col.pack(side=tk.LEFT, expand=True)
    dot = tk.Label(col, text="●", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_ROOT)
    dot.pack()
    lbl = tk.Label(col, text=nome, font=FNT_STATUS, fg=TEXT_MUTED, bg=BG_ROOT)
    lbl.pack()
    step_widgets[nome] = (dot, lbl)
    if i < len(ETAPAS) - 1:
        tk.Label(frame_pipe, text="──", font=("Segoe UI", 9),
                 fg=TEXT_MUTED, bg=BG_ROOT).pack(side=tk.LEFT, expand=True)

def marcar_etapa(nome, estado="ativo"):
    cores = {
        "ativo":     (ACCENT,     TEXT_PRIMARY),
        "concluido": (ACCENT_DIM, TEXT_MUTED),
        "erro":      (TEXT_ERROR, TEXT_ERROR),
        "aviso":     (TEXT_WARN,  TEXT_WARN),
    }
    cor_dot, cor_lbl = cores.get(estado, (TEXT_MUTED, TEXT_MUTED))
    dot, lbl = step_widgets[nome]
    # Atualiza sempre na thread principal via after
    root.after(0, lambda: (dot.config(fg=cor_dot), lbl.config(fg=cor_lbl)))

# ── BARRA INFERIOR: STATUS + BOTÃO ───────────────────────────────
# DEVE ser declarada ANTES do log (que tem expand=True) para não ser espremida
footer = tk.Frame(root, bg=BG_FOOTER, height=52)
footer.pack(fill=tk.X, side=tk.BOTTOM)
footer.pack_propagate(False)

status_dot = tk.Label(footer, text="●", font=("Segoe UI", 9),
                      fg=ACCENT, bg=BG_FOOTER)
status_dot.pack(side=tk.LEFT, padx=(12, 4), pady=14)

status_var = tk.StringVar(value="Detectando planilha…")
tk.Label(footer, textvariable=status_var, font=FNT_STATUS,
         fg=TEXT_MUTED, bg=BG_FOOTER).pack(side=tk.LEFT)

btn_executar = tk.Button(
    footer, text="▶  EXECUTAR",
    font=FNT_BTN, fg=BG_ROOT, bg=ACCENT,
    activebackground=ACCENT_DIM, activeforeground=TEXT_PRIMARY,
    relief=tk.FLAT, padx=20, pady=6,
    cursor="hand2",
)
btn_executar.pack(side=tk.RIGHT, padx=16, pady=8)

def set_status(texto, cor=None):
    root.after(0, lambda: (
        status_var.set(texto),
        status_dot.config(fg=cor or ACCENT),
    ))

def set_btn(estado):
    """'normal' | 'running' | 'done' | 'erro'"""
    configs = {
        "normal":  dict(text="▶  EXECUTAR",    bg=ACCENT,      fg=BG_ROOT,      state=tk.NORMAL),
        "running": dict(text="⏳  EXECUTANDO…", bg=TEXT_MUTED,  fg=BG_ROOT,      state=tk.DISABLED),
        "done":    dict(text="✓  CONCLUÍDO",   bg=ACCENT_DIM,  fg=TEXT_PRIMARY,  state=tk.DISABLED),
        "erro":    dict(text="✗  ERRO",        bg=TEXT_ERROR,  fg=TEXT_PRIMARY,  state=tk.NORMAL),
    }
    cfg = configs.get(estado, configs["normal"])
    root.after(0, lambda: btn_executar.config(**cfg))


# ── TERMINAL DE LOG ──────────────────────────────────────────────
# Declarado DEPOIS do footer para que expand=True não empurre o footer pra fora
frame_log_outer = tk.Frame(root, bg=BG_CARD, padx=2, pady=2)
frame_log_outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=(6, 6))

frame_log_inner = tk.Frame(frame_log_outer, bg=BG_CARD)
frame_log_inner.pack(fill=tk.BOTH, expand=True)

tk.Label(frame_log_inner, text="  ● TERMINAL DE EXECUÇÃO", font=FNT_LABEL,
         fg=ACCENT, bg=BG_CARD, anchor="w", pady=6).pack(fill=tk.X, padx=8)

tk.Frame(frame_log_inner, bg=ACCENT_DIM, height=1).pack(fill=tk.X, padx=8)

log_frame = tk.Frame(frame_log_inner, bg="#111111")
log_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

txt_log = tk.Text(
    log_frame, wrap=tk.WORD, bg="#111111", fg=TEXT_LOG,
    font=FNT_LOG, bd=0, highlightthickness=0,
    selectbackground=ACCENT_DIM, insertbackground=ACCENT,
    state=tk.DISABLED, padx=8, pady=6,
)
scrollbar = tk.Scrollbar(log_frame, command=txt_log.yview,
                         bg="#1a1a1a", troughcolor="#111111", width=8)
txt_log.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

txt_log.tag_configure("info",     foreground=TEXT_LOG)
txt_log.tag_configure("aviso",    foreground=TEXT_WARN)
txt_log.tag_configure("erro",     foreground=TEXT_ERROR)
txt_log.tag_configure("ts",       foreground=TEXT_MUTED)
txt_log.tag_configure("destaque", foreground=TEXT_PRIMARY)
txt_log.tag_configure("sep",      foreground=ACCENT_DIM)

def registrar_log(msg, nivel="info"):
    """Thread-safe: agenda inserção no event loop do Tk."""
    def _inserir():
        ts = datetime.now().strftime("%H:%M:%S")
        txt_log.config(state=tk.NORMAL)
        txt_log.insert(tk.END, f"[{ts}]  ", "ts")
        txt_log.insert(tk.END, f"{msg}\n", nivel)
        txt_log.see(tk.END)
        txt_log.config(state=tk.DISABLED)
    root.after(0, _inserir)


# ══════════════════════════════════════════════════════════════════
#  NOTIFICAÇÃO WINDOWS (Toast / Balloon)
# ══════════════════════════════════════════════════════════════════
def notificar_windows(titulo: str, mensagem: str, arquivo: str = ""):
    """Exibe notificação nativa do Windows (Toast Win10/11 com fallback Balloon)."""
    try:
        t = str(titulo or "").replace("\r", " ").replace("\n", " ").strip()
        m = str(mensagem or "").replace("\r", " ").replace("\n", " ").strip()
        t_xml = html.escape(t, quote=False)
        m_xml = html.escape(m, quote=False)

        abrir = ""
        if arquivo:
            arquivo_ps = str(arquivo).replace("'", "''")
            abrir = f"""
try {{
    Start-Process explorer.exe -ArgumentList "/select,`"{arquivo_ps}`""
}} catch {{}}
"""

        ps = f"""$ErrorActionPreference = 'SilentlyContinue'
$t = @'
{t_xml}
'@
$m = @'
{m_xml}
'@
$ok = $false
try {{
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType=WindowsRuntime]
    $AppId = '{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe'
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
}} catch {{ $ok = $false }}
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
{abrir}
Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""
        fd, path = tempfile.mkstemp(suffix=".ps1", prefix="prodforge_")
        os.close(fd)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(ps)
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive",
             "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", path],
            creationflags=flags, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        registrar_log(f"Notificação: falha ({e})", "aviso")


# ══════════════════════════════════════════════════════════════════
#  FECHAMENTO CIRÚRGICO — ARQUIVOS SAP
# ══════════════════════════════════════════════════════════════════
def _fechar_prod_xlsx_definitivo(pasta_destino: str, nome_arquivo: str, timeout: int = 20) -> bool:
    """
    Localiza a instância exata do Excel que contém o arquivo (varrendo a ROT),
    fecha o arquivo cirurgicamente, trata estado de recuperação de documentos,
    e confirma que o SO liberou o bloqueio.
    """
    import win32process
    nome_base = Path(nome_arquivo).name.lower()
    caminho_full = Path(pasta_destino) / nome_arquivo

    # 1. Aguarda o arquivo existir fisicamente
    for _ in range(timeout * 2):
        if caminho_full.exists():
            break
        time.sleep(0.5)
    else:
        registrar_log(f"⚠ {nome_arquivo} nunca apareceu no disco.", "aviso")
        return False

    registrar_log(f"{nome_arquivo} detectado no disco.")
    registrar_log("Verificando instâncias Excel (ROT)…")

    _com_init = False
    try:
        pythoncom.CoInitialize()
        _com_init = True
    except Exception:
        pass

    sucesso = False
    try:
        wb_encontrado = None
        app_excel = None
        ctx = pythoncom.CreateBindCtx(0)

        # 2. Polling na ROT (Running Object Table)
        for _ in range(timeout * 2):
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot:
                try:
                    display = moniker.GetDisplayName(ctx, None)
                    if not display: continue
                    if Path(display).name.lower() == nome_base:
                        obj = rot.GetObject(moniker)
                        wb = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                        _ = wb.Sheets.Count  # Confirma que COM está pronto
                        wb_encontrado = wb
                        app_excel = wb.Application
                        break
                except Exception:
                    continue
            
            if wb_encontrado:
                break
            time.sleep(0.5)

        if not wb_encontrado:
            registrar_log(f"⚠ Motivo: workbook não ficou acessível via COM em {timeout}s.", "aviso")
            return False

        # 3. Tratamento e Fechamento
        try:
            hwnd = app_excel.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            registrar_log(f"{nome_arquivo} localizado na instância PID {pid}.")
        except Exception:
            pid = "Desconhecido"
            registrar_log(f"{nome_arquivo} localizado em instância Excel.")

        try:
            app_excel.DisplayAlerts = False
        except Exception:
            registrar_log("Recuperação de documentos detectada ou Excel bloqueado.", "aviso")
            registrar_log("Tratando estado de recuperação…")

        restantes = app_excel.Workbooks.Count
        registrar_log(f"Fechando {nome_arquivo}…")
        
        try:
            wb_encontrado.Close(SaveChanges=False)
            registrar_log(f"{nome_arquivo} fechado.")
        except Exception as e:
            registrar_log(f"⚠ Erro ao fechar via COM: {e}", "aviso")

        if restantes <= 1:
            try:
                # Se há recuperação de docs pendente, o Quit sem DisplayAlerts descarta.
                app_excel.Quit()
                registrar_log(f"Instância exclusiva (PID {pid}) finalizada.")
            except Exception:
                pass

        # 4. Confirmação COM
        time.sleep(1)
        ainda_aberto = False
        rot = pythoncom.GetRunningObjectTable()
        for moniker in rot:
            try:
                display = moniker.GetDisplayName(ctx, None)
                if display and Path(display).name.lower() == nome_base:
                    ainda_aberto = True
                    break
            except Exception:
                pass

        if ainda_aberto:
            registrar_log("⚠ Confirmado: workbook ainda está aberto na ROT.", "aviso")
        else:
            registrar_log("Confirmado: workbook não está mais aberto.")
            
        sucesso = not ainda_aberto

    except Exception as e:
        registrar_log(f"⚠ Erro no gerenciamento COM: {e}", "aviso")
    finally:
        if _com_init:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # 5. Confirmação de SO (bloqueio de arquivo)
    if sucesso:
        for _ in range(10):  # Aguarda até 5s
            if not _checar_arquivo_em_uso(str(caminho_full)):
                registrar_log("Confirmado: arquivo liberado.")
                return True
            time.sleep(0.5)
        registrar_log("⚠ Confirmado: arquivo continua bloqueado pelo SO.", "aviso")
        return False
        
    return sucesso



# ══════════════════════════════════════════════════════════════════
#  UTILITÁRIOS SAP / EXCEL
# ══════════════════════════════════════════════════════════════════
def enviar_para_clipboard(texto):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(texto, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()

def limpar_memoria_mb51(session):
    campos = [
        "wnd[0]/usr/ctxtMATNR-LOW","wnd[0]/usr/ctxtWERKS-LOW","wnd[0]/usr/ctxtLGORT-LOW",
        "wnd[0]/usr/ctxtCHARG-LOW","wnd[0]/usr/ctxtLIFNR-LOW","wnd[0]/usr/ctxtKUNNR-LOW",
        "wnd[0]/usr/ctxtBWART-LOW","wnd[0]/usr/ctxtSOBKZ-LOW","wnd[0]/usr/ctxtAUFNR-LOW",
        "wnd[0]/usr/txtMAT_KDAU-LOW","wnd[0]/usr/txtMAT_KDPO-LOW","wnd[0]/usr/ctxtBUDAT-LOW",
        "wnd[0]/usr/txtUSNAM-LOW","wnd[0]/usr/ctxtVGART-LOW","wnd[0]/usr/txtXBLNR-LOW",
        "wnd[0]/usr/ctxtALV_DEF","wnd[0]/usr/ctxtMATNR-HIGH","wnd[0]/usr/ctxtWERKS-HIGH",
        "wnd[0]/usr/ctxtLGORT-HIGH","wnd[0]/usr/ctxtBWART-HIGH","wnd[0]/usr/ctxtBUDAT-HIGH",
    ]
    for c in campos:
        try: session.findById(c).Text = ""
        except: pass

def extrair_dado_seguro(session, ids):
    for i in ids:
        try:
            v = session.findById(i).Text
            if v: return str(v).strip()
        except: pass
    return ""

# ══════════════════════════════════════════════════════════════════
#  ABRIR SAP
# ══════════════════════════════════════════════════════════════════

def conectar_sap_ou_abrir():
    """Realiza a verificação, abertura visível e conexão ao SAP."""
    registrar_log("Verificando conexão com o SAP...")
    SAPGui = None
    try:
        SAPGui = win32.GetObject("SAPGUI")
    except Exception:
        pass

    if SAPGui is None:
        registrar_log("SAP fechado. Abrindo o SAP Logon na tela...")
        try:
            caminho_sap = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            if not os.path.exists(caminho_sap):
                caminho_sap = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
            
            # Usa o shell ou garante que a janela seja exibida na frente
            os.startfile(caminho_sap)
        except Exception as e:
            try:
                subprocess.Popen([caminho_sap])
            except Exception as ex:
                raise Exception(f"Não foi possível iniciar o SAP Logon: {ex}")
        
        registrar_log("Aguardando o SAP Logon aparecer...")
        timeout_sap = 30
        for _ in range(timeout_sap):
            try:
                SAPGui = win32.GetObject("SAPGUI")
                if SAPGui is not None:
                    break
            except Exception:
                pass
            time.sleep(1)
        else:
            raise Exception("Tempo limite (30s) excedido aguardando o SAP Logon iniciar.")

    SAPApp = SAPGui.GetScriptingEngine
    
    # Garante que o aplicativo do SAP Logon traga as conexões visíveis
    if SAPApp.Children.Count == 0:
        registrar_log("Abrindo conexão com a instância SAP (SAP_CONNECTION_NAME)...")
        SAPCon = SAPApp.OpenConnection("SAP_CONNECTION_NAME", True)
    else:
        SAPCon = SAPApp.Children(0)
    
    time.sleep(3)
    
    # Tenta focar na janela principal e auto-confirmar pop-ups de conexão (wnd[1])
    try:
        session = SAPCon.Children(0)
        # Se houver pop-up de aviso/confirmação (wnd[1]), clica em Ok/Enter
        try:
            if session.Children.Count > 1 or session.ActiveWindow.Name == "wnd[1]":
                session.findById("wnd[1]").sendVKey(0)
                time.sleep(1)
        except Exception:
            pass
        session.findById("wnd[0]").maximize()
    except Exception:
        time.sleep(2)
        session = SAPCon.Children(0)
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1)
        except Exception:
            pass
        session.findById("wnd[0]").maximize()

    return session

# ══════════════════════════════════════════════════════════════════
#  COR3 — CONSULTA DE ORDEM
# ══════════════════════════════════════════════════════════════════
def consultar_ordem_cor3(session, numero_ordem):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCOR3"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = str(numero_ordem)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Fecha pop-up de aviso
        if session.Children.Count > 1:
            try:
                session.findById("wnd[1]").sendVKey(0)
                time.sleep(0.5)
            except: pass

        if session.findById("wnd[0]/sbar").MessageType == "E":
            return {"erro": session.findById("wnd[0]/sbar").Text}

        material_bruto = extrair_dado_seguro(session, [
            "wnd[0]/usr/subORD_HEADER:SAPLCOKO:5800/txtCAUFVD-MATNR",
            "wnd[0]/usr/subORD_HEADER:SAPLCOKO:5800/ctxtCAUFVD-MATNR",
            "wnd[0]/usr/txtCAUFVD-MATNR",
            "wnd[0]/usr/ctxtCAUFVD-MATNR",
        ])
        material = material_bruto.lstrip("0") if material_bruto else "Não encontrado"

        try:
            session.findById("wnd[0]/tbar[1]/btn[6]").Press()
            time.sleep(1)
        except: pass

        recurso_sap = extrair_dado_seguro(session, [
            "wnd[0]/usr/tblSAPLCOVGTCTRL_5100/ctxtAFVGD-ARBPL[4,0]",
            "wnd[0]/usr/tblSAPLCOVGTCTRL_5100/txtAFVGD-ARBPL[4,0]",
        ])
        if not recurso_sap:
            try:
                usr = session.findById("wnd[0]/usr")
                for i in range(usr.Children.Count):
                    if usr.Children(i).Type == "GuiTableControl":
                        recurso_sap = usr.Children(i).Rows(0).Cells(4).Text.strip()
                        break
            except: pass

        linha_nome = DEPARA_LINHAS.get(recurso_sap, recurso_sap) if recurso_sap else "Não encontrado"

        return {
            "ordem":         numero_ordem,
            "material":      material,
            "recurso_codigo": recurso_sap,
            "recurso_nome":  linha_nome,
            "erro":          None,
        }
    except Exception as e:
        return {"erro": f"Falha no SAP: {e}"}


# ══════════════════════════════════════════════════════════════════
#  COR3 — ATUALIZAÇÃO DA BASE ORDENS
# ══════════════════════════════════════════════════════════════════
def aplicar_cor3(wb_base, session, excel):
    marcar_etapa("Verificação COR3", "ativo")
    set_status("Verificando ordens não cadastradas…")

    # Abas
    try:
        ws_mb51 = wb_base.Sheets("MB51- PRODPLANTA")
    except Exception:
        registrar_log("Aba 'MB51- PRODPLANTA' não encontrada. COR3 ignorada.", "aviso")
        marcar_etapa("Verificação COR3", "aviso")
        return
    try:
        ws_bo = wb_base.Sheets("Base Ordens")
    except Exception:
        registrar_log("Aba 'Base Ordens' não encontrada. COR3 ignorada.", "aviso")
        marcar_etapa("Verificação COR3", "aviso")
        return

    # ── Coleta ordens com "NÃO ENCONTRADO" na coluna K (Leitura Rápida) ───────────
    # Coluna F (índice 6) = Ordem  |  Coluna K (índice 11) = Linha
    registrar_log("Varrendo coluna K por 'NÃO ENCONTRADO'…")
    ult = ws_mb51.Cells(ws_mb51.Rows.Count, "A").End(-4162).Row

    ordens_novas = []
    
    if ult >= 2:
        dados_bloco = ws_mb51.Range(f"F2:K{ult}").Value
        
        if isinstance(dados_bloco, tuple):
            if not isinstance(dados_bloco[0], tuple):
                dados_bloco = (dados_bloco,)
                
            for linha_dados in dados_bloco:
                val_k = str(linha_dados[5] or "").strip().upper()
                
                if "NÃO ENCONTRADO" in val_k or "NAO ENCONTRADO" in val_k:
                    val_f = str(linha_dados[0] or "").strip()
                    
                    if val_f.endswith(".0"):
                        val_f = val_f[:-2]
                        
                    if val_f and val_f not in ordens_novas:
                        ordens_novas.append(val_f)

    if not ordens_novas:
        registrar_log("Todas as ordens já estão cadastradas. Base Ordens OK.", "destaque")
        marcar_etapa("Verificação COR3", "concluido")
        marcar_etapa("Base Ordens", "concluido")
        return

    registrar_log(
        f"{len(ordens_novas)} ordem(ns) nova(s): {', '.join(ordens_novas)}", "destaque"
    )
    marcar_etapa("Verificação COR3", "concluido")
    marcar_etapa("Base Ordens", "ativo")
    set_status(f"Cadastrando {len(ordens_novas)} ordem(ns) via COR3…")

    ult_bo = ws_bo.Cells(ws_bo.Rows.Count, "A").End(-4162).Row
    linha_ref = ult_bo

    for i, ordem in enumerate(ordens_novas, 1):
        registrar_log(f"[{i}/{len(ordens_novas)}] COR3 → Ordem {ordem}…")
        set_status(f"COR3 — Ordem {ordem} ({i}/{len(ordens_novas)})…")

        res = consultar_ordem_cor3(session, ordem)
        if res.get("erro"):
            registrar_log(f"  ✗ {res['erro']}", "erro")
            continue

        nova = ult_bo + i

        ws_bo.Cells(nova, 1).Value = res["material"]
        ws_bo.Cells(nova, 2).Value = res["ordem"]
        ws_bo.Cells(nova, 3).Value = res["recurso_nome"]

        if linha_ref >= 2:
            try:
                ws_bo.Range(
                    ws_bo.Cells(linha_ref, 4), ws_bo.Cells(linha_ref, 6)
                ).Copy(Destination=ws_bo.Range(ws_bo.Cells(nova, 4), ws_bo.Cells(nova, 6)))
            except Exception as ex:
                registrar_log(f"  ⚠ Fórmulas D:F: {ex}", "aviso")

        if linha_ref >= 2:
            try:
                ws_bo.Range(
                    ws_bo.Cells(linha_ref, 1), ws_bo.Cells(linha_ref, 6)
                ).Copy()
                ws_bo.Range(
                    ws_bo.Cells(nova, 1), ws_bo.Cells(nova, 6)
                ).PasteSpecial(Paste=-4122)
                excel.CutCopyMode = False
            except Exception as ex:
                registrar_log(f"  ⚠ Formatação: {ex}", "aviso")

        registrar_log(
            f"  ✓ Material: {res['material']}  |  Linha: {res['recurso_nome']}", "destaque"
        )

    registrar_log("Ordenando Base Ordens pela coluna C (A→Z)...")
    nova_ult_bo = ws_bo.Cells(ws_bo.Rows.Count, "A").End(-4162).Row
    if nova_ult_bo >= 2:
        try:
            # Ordena a CurrentRegion a partir de A1. Garante que os dados inteiros das linhas
            # se movam juntos baseados na Coluna C, sem separar os dados de cada ordem.
            ws_bo.Range(f"A1:Z{nova_ult_bo}").Sort(
                Key1=ws_bo.Range("C1"), 
                Order1=1, # 1 = xlAscending
                Header=1, # 1 = xlYes
                Orientation=1 # 1 = xlSortColumns (sort by row)
            )
            registrar_log("Base Ordens organizada com sucesso.", "destaque")
        except Exception as e:
            registrar_log(f"⚠ Erro ao ordenar Base Ordens: {e}", "aviso")

    marcar_etapa("Base Ordens", "concluido")
    registrar_log("Base Ordens atualizada.", "destaque")

# ══════════════════════════════════════════════════════════════════
#  MM03 — BUSCA DE PESOS FALTANTES
# ══════════════════════════════════════════════════════════════════
def aplicar_mm03_pesos(wb_base, session, excel):
    marcar_etapa("Verificação MM03", "ativo")
    set_status("Verificando SKUs sem peso...")

    try:
        ws_bo = wb_base.Sheets("Base Ordens")
    except Exception:
        registrar_log("Aba 'Base Ordens' não encontrada. MM03 ignorada.", "aviso")
        marcar_etapa("Verificação MM03", "aviso")
        return
    try:
        ws_bl = wb_base.Sheets("Base Linha")
    except Exception:
        registrar_log("Aba 'Base Linha' não encontrada. MM03 ignorada.", "aviso")
        marcar_etapa("Verificação MM03", "aviso")
        return

    last_row_base_ordens = ws_bo.Cells(ws_bo.Rows.Count, "A").End(-4162).Row
    missing_weight_skus = {}
    
    if last_row_base_ordens >= 2:
        base_ordens_data = ws_bo.Range(f"A2:F{last_row_base_ordens}").Value
        if isinstance(base_ordens_data, tuple):
            if not isinstance(base_ordens_data[0], tuple):
                base_ordens_data = (base_ordens_data,)
            for row in base_ordens_data:
                base_peso_status = str(row[5] or "").strip().upper()
                if "NÃO ENCONTRADO" in base_peso_status or "NAO ENCONTRADO" in base_peso_status:
                    sku_code = str(row[0] or "").strip()
                    if sku_code.endswith(".0"):
                        sku_code = sku_code[:-2]
                    linha_name = str(row[2] or "").strip()
                    if sku_code and sku_code not in missing_weight_skus:
                        missing_weight_skus[sku_code] = linha_name

    if not missing_weight_skus:
        registrar_log("Todos os SKUs possuem peso. MM03 OK.", "destaque")
        marcar_etapa("Verificação MM03", "concluido")
        return

    registrar_log(f"{len(missing_weight_skus)} SKU(s) novo(s) detectado(s). Buscando pesos...", "destaque")
    set_status(f"Extraindo {len(missing_weight_skus)} peso(s) na MM03...")

    ult_bl = ws_bl.Cells(ws_bl.Rows.Count, "A").End(-4162).Row

    def limpar_numero(texto):
        try:
            return float(texto.replace(".", "").replace(",", "."))
        except:
            return 0.0

    tem_manual = False
    
    for i, (sku, linha_nome) in enumerate(missing_weight_skus.items(), 1):
        registrar_log(f"[{i}/{len(missing_weight_skus)}] Processando SKU {sku} (Linha: {linha_nome})")
        set_status(f"MM03 — SKU {sku} ({i}/{len(missing_weight_skus)})…")
        
        peso_liquido = 0.0
        base_64 = ""
        
        if linha_nome in ["MP04", "MP07", "MP08", "MP09"]:
            peso_liquido = 1.000
            base_64 = 1.000
            registrar_log(f"  ↳ Linha especial ({linha_nome}): Peso e Base64 = 1.000")
        else:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)

                session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = str(sku)
                session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(str(sku))
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.8)

                try:
                    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = True
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    time.sleep(1)
                except:
                    pass

                pl_id = "wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2007/txtMARA-NTGEW"
                try:
                    session.findById(pl_id).setFocus()
                    peso_liquido = limpar_numero(str(session.findById(pl_id).text).strip())
                except:
                    pass
                    
                base_64 = "CADASTRO MANUAL"
                tem_manual = True
                registrar_log(f"  ↳ Peso Líquido SAP: {peso_liquido}")
            except Exception as e_mm:
                registrar_log(f"  Aviso MM03 SKU {sku}: {e_mm}", "aviso")
                peso_liquido = 0.0
                base_64 = "ERRO SAP"
                tem_manual = True

        nova_linha = ult_bl + i
        ws_bl.Cells(nova_linha, 1).Value = sku
        ws_bl.Cells(nova_linha, 2).Value = peso_liquido
        ws_bl.Cells(nova_linha, 3).Value = base_64
        
        if ult_bl >= 2:
            try:
                ws_bl.Range(ws_bl.Cells(ult_bl, 1), ws_bl.Cells(ult_bl, 3)).Copy()
                ws_bl.Range(ws_bl.Cells(nova_linha, 1), ws_bl.Cells(nova_linha, 3)).PasteSpecial(Paste=-4122)
                excel.CutCopyMode = False
            except:
                pass

    if tem_manual:
        registrar_log("⚠ ALERTA: Há SKUs que exigem cadastro MANUAL da Base 64 na aba 'Base Linha'.", "aviso")
        
    marcar_etapa("Verificação MM03", "concluido")
    registrar_log("Base Linha (Pesos) atualizada.", "destaque")

# ══════════════════════════════════════════════════════════════════
#  FLUXO PRINCIPAL  (roda em thread separada)
# ══════════════════════════════════════════════════════════════════
_info_base = None   # preenchido na detecção inicial


# ══════════════════════════════════════════════════════════════════
#  CORREÇÃO DO PIVOTCACHE XML
#  Edita o atributo ref="B1:N{ult}" direto no ZIP do xlsx,
#  sem abrir o Excel — única forma de atualizar o range fixo
#  sem destruir segmentações, formatação ou layout da pivot.
# ══════════════════════════════════════════════════════════════════
def _corrigir_pivot_cache_xml(caminho_xlsx: str, nome_aba: str) -> str | None:
    """
    Lê o arquivo xlsx, encontra todos os pivotCacheDefinition que
    apontam para `nome_aba`, recalcula a última linha real de dados
    e substitui o atributo ref="..." pelo novo intervalo.

    Retorna a string do novo range se alterou algo, ou None se não
    encontrou nenhum cache para aquela aba.
    """
    import zipfile
    import re
    import shutil
    import tempfile
    import openpyxl

    # ── 1. Última linha via pandas (rápido) ──
    import pandas as pd
    try:
        # Lê apenas A:J para identificar o último registro real (ignora fórmulas a partir de K)
        df_data = pd.read_excel(caminho_xlsx, sheet_name=nome_aba, header=0, usecols="A:J")
    except Exception:
        return None

    if df_data.empty:
        return None

    # Encontra a última linha que efetivamente possui dados em A:J
    last_valid = df_data.last_valid_index()
    if last_valid is None:
        ult_linha = 1
    else:
        ult_linha = last_valid + 2  # +1 por causa do índice 0 e +1 pelo cabeçalho

    # ── 3. Edita o ZIP in-place ───────────────────────────────────
    tmp_path = caminho_xlsx + ".tmp_pivot"
    alterado = False
    novo_range_aplicado = None

    with zipfile.ZipFile(caminho_xlsx, "r") as zin, \
         zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            dados = zin.read(item.filename)

            if "pivotCacheDefinition" in item.filename and item.filename.endswith(".xml"):
                xml = dados.decode("utf-8")

                # Só altera caches que referenciem a aba correta
                if f'sheet="{nome_aba}"' in xml:
                    # Substitui ref="A1:N<qualquer_numero>" pelo novo range, preservando as colunas
                    def repl_range(match):
                        nonlocal novo_range_aplicado
                        inicio = match.group(1)
                        ref_old = match.group(2)
                        fim = match.group(3)
                        
                        if ":" in ref_old:
                            partes = ref_old.split(":")
                            # Extrai apenas as letras (ex: 'N70000' -> 'N')
                            col_esq = re.sub(r'\d+', '', partes[0])
                            col_dir = re.sub(r'\d+', '', partes[1])
                            novo_r = f"{col_esq}1:{col_dir}{ult_linha}"
                        else:
                            novo_r = f"A1:N{ult_linha}"
                            
                        novo_range_aplicado = novo_r
                        return f"{inicio}{novo_r}{fim}"

                    xml_novo = re.sub(
                        r'(<worksheetSource\s[^>]*ref=")([^"]+)(")',
                        repl_range,
                        xml,
                    )
                    if xml_novo != xml:
                        alterado = True
                    dados = xml_novo.encode("utf-8")

            zout.writestr(item, dados)

    if alterado:
        os.replace(tmp_path, caminho_xlsx)
    else:
        os.remove(tmp_path)
        return None

    return novo_range_aplicado


def executar_automacao():
    global _info_base

    excel = None
    wb_base = None

    try:
        info    = _info_base
        hoje    = datetime.today()
        # Usa o range calculado pela lógica inteligente de virada de mês
        sInicio = info["data_inicio"].strftime("%d.%m.%Y")
        sFim    = info["data_fim"].strftime("%d.%m.%Y")
        hoje_s  = hoje.strftime("%d.%m.%Y")
        dia_s   = hoje.strftime("%d.%m")

        registrar_log(f"Período MB51 configurado: {sInicio} → {sFim}", "destaque")

        # ── Gera caminhos de destino coerentes com o mês selecionado ──────────────
        dest = gerar_caminho_destino(info)
        pasta_destino = dest["pasta_destino"]
        nome_novo     = dest["nome_novo"]
        arquivo_novo  = dest["arquivo_novo"]

        # Cria a pasta de destino se não existir
        os.makedirs(pasta_destino, exist_ok=True)
        registrar_log(f"Pasta destino: {pasta_destino}")
        registrar_log(f"Arquivo destino: {dest['rel_destino']}")

        # ── Verifica se arquivo destino está em uso ────────────────
        if _checar_arquivo_em_uso(arquivo_novo):
            raise PermissionError(
                f"O arquivo de destino está aberto por outro usuário/processo:\n{arquivo_novo}\n\n"
                f"Peça para fechar o arquivo e execute novamente."
            )

        # ── Verifica se arquivo base (origem) está em uso ──────────
        if _checar_arquivo_em_uso(info["arquivo_base"]):
            registrar_log(
                f"⚠ Arquivo base está aberto: {os.path.basename(info['arquivo_base'])}. "
                "Continuando mesmo assim (somente leitura).", "aviso"
            )

        registrar_log(f"Copiando base de {info['dia_ref']} → {dest['rel_destino']}…")
        shutil.copy2(info["arquivo_base"], arquivo_novo)
        registrar_log(f"Base copiada: {nome_novo}", "destaque")

        # ── Pasta de exportação SAP (definida na interface) ──────────
        cfg_atual = _carregar_config()
        pasta_temp_cfg = cfg_atual.get("pasta_exportacao_sap", "")

        if not pasta_temp_cfg or not os.path.isdir(os.path.normpath(pasta_temp_cfg)):
            raise Exception("Pasta de exportação inválida ou não selecionada na interface. Cancele e selecione a pasta.")

        pasta_temp = os.path.normpath(pasta_temp_cfg)
        registrar_log(f"Pasta de exportação: {pasta_temp}", "destaque")

        os.makedirs(pasta_temp, exist_ok=True)
        caminho_prod = os.path.normpath(os.path.join(pasta_temp, "PROD.xlsx"))
        try:
            if os.path.exists(caminho_prod):
                os.remove(caminho_prod)
        except: pass

        # ════════════════════════════════════════════════════════
        # ETAPA 1 — CONEXÃO SAP
        # ════════════════════════════════════════════════════════
        marcar_etapa("Conexão SAP", "ativo")
        set_status("Conectando ao SAP…")
        
        # Chama a rotina robusta de conexão
        session = conectar_sap_ou_abrir()
        
        marcar_etapa("Conexão SAP", "concluido")
        registrar_log("SAP conectado com sucesso.", "destaque")

        # ════════════════════════════════════════════════════════
        # ETAPA 2 — EXTRAÇÃO MB51
        # ════════════════════════════════════════════════════════
        marcar_etapa("Extração MB51", "ativo")
        set_status("Executando transação MB51…")
        registrar_log("Executando transação MB51 — Produção (mov. 101 e 102)…")

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB51"
        session.findById("wnd[0]").sendVKey(0)

        limpar_memoria_mb51(session)

        session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1000"
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text  = "TS00"
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text  = sInicio
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = sFim
        session.findById("wnd[0]/usr/ctxtALV_DEF").Text    = "/PRODMDC"

        try:
            session.findById("wnd[0]/usr/radRFLAT_L").Select()
        except:
            pass

        enviar_para_clipboard("101\r\n102")
        session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").Press()
        session.findById("wnd[1]/tbar[0]/btn[16]").Press()
        session.findById("wnd[1]/tbar[0]/btn[24]").Press()
        session.findById("wnd[1]/tbar[0]/btn[8]").Press()

        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(16)

        session.findById("wnd[1]/tbar[0]/btn[20]").Press()
        # SAP exige caminho com barras invertidas e sem barra no final
        pasta_temp_sap = os.path.normpath(pasta_temp)
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text     = pasta_temp_sap
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PROD.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()

        set_status("Aguardando exportação do SAP…")
        registrar_log("Aguardando PROD.xlsx abrir pelo SAP…")

        # Chama a nova função robusta para aguardar e fechar a instância
        sucesso_fechamento = _fechar_prod_xlsx_definitivo(pasta_temp, "PROD.xlsx", timeout=30)
        if not sucesso_fechamento:
            registrar_log("⚠ O arquivo pode ainda estar bloqueado. A automação tentará continuar, mas poderá falhar na abertura.", "aviso")

        # ── Limpa arquivos antigos da pasta de exportação ─────────
        try:
            for arq in os.listdir(pasta_temp):
                if arq.upper().startswith("PROD") and arq.upper().endswith(".XLSX") and arq.upper() != "PROD.XLSX":
                    try:
                        os.remove(os.path.join(pasta_temp, arq))
                    except Exception:
                        pass
        except Exception:
            pass

        session.findById("wnd[0]").sendVKey(3)
        marcar_etapa("Extração MB51", "concluido")
        registrar_log("MB51 exportado com sucesso.", "destaque")

        # ════════════════════════════════════════════════════════
        # ETAPA 3 — TRATAMENTO EXCEL
        # ════════════════════════════════════════════════════════
        marcar_etapa("Tratamento Excel", "ativo")
        set_status("Abrindo Excel e tratando dados…")
        registrar_log("Conectando ao Excel via COM…")

        excel = _ensure_excel_dispatch()
        excel.Visible = True
        excel.DisplayAlerts = False

        # Abre o arquivo de relatório de produção (destino)
        wb_base      = excel.Workbooks.Open(arquivo_novo)
        ws_mb51_dest = wb_base.Sheets("MB51- PRODPLANTA")

        registrar_log("Limpando A2:J100000…")
        ws_mb51_dest.Range("A2:J100000").ClearContents()

        # Abre o PROD.xlsx exportado pelo SAP (origem)
        registrar_log("Abrindo PROD.xlsx exportado pelo SAP…")
        wb_prod  = excel.Workbooks.Open(caminho_prod)
        ws_prod  = wb_prod.Sheets(1)
        ult_prod = ws_prod.Cells(ws_prod.Rows.Count, "A").End(-4162).Row

        if ult_prod > 1:
            ws_prod.Range(f"A2:J{ult_prod}").Copy(
                Destination=ws_mb51_dest.Range("A2")
            )

            # ── Detecta e estende fórmulas K em diante ───────────
            registrar_log("Estendendo fórmulas e formatação (col K→)…")
            ultima_col = None
            for col in range(11, 30):
                cel = ws_mb51_dest.Cells(2, col)
                if cel.HasFormula or str(cel.Value or "").strip():
                    ultima_col = col
                else:
                    break

            if ultima_col:
                col_fim = chr(ord("A") + ultima_col - 1) if ultima_col <= 26 else "Z"
                rng_ref  = ws_mb51_dest.Range(f"K2:{col_fim}2")
                rng_dest = ws_mb51_dest.Range(f"K2:{col_fim}{ult_prod}")
                rng_ref.Copy()
                rng_dest.PasteSpecial(Paste=-4122)   # formatos
                rng_dest.PasteSpecial(Paste=-4123)   # fórmulas
                excel.CutCopyMode = False
                registrar_log(f"Fórmulas estendidas K:{col_fim} até linha {ult_prod}.")
            else:
                registrar_log("Nenhuma fórmula detectada além da col J.", "aviso")

            # ── TextToColumns B C F I ─────────────────────────────
            registrar_log("Texto para Colunas (B, C, F, I)…")
            for col in ["B", "C", "F", "I"]:
                try:
                    ws_mb51_dest.Columns(f"{col}:{col}").TextToColumns(
                        Destination=ws_mb51_dest.Range(f"{col}1"),
                        DataType=1, TextQualifier=1,
                        ConsecutiveDelimiter=False,
                        Tab=True, Semicolon=False, Comma=False,
                        Space=False, Other=False,
                        FieldInfo=((1, 1),),
                    )
                except Exception as ex:
                    registrar_log(f"  ⚠ Col {col}: {ex}", "aviso")

        wb_prod.Close(False)
        marcar_etapa("Tratamento Excel", "concluido")
        registrar_log("Tratamento Excel concluído.", "destaque")

        # ════════════════════════════════════════════════════════
        # ETAPAS 4 e 5 — COR3 + BASE ORDENS
        # ════════════════════════════════════════════════════════
        aplicar_cor3(wb_base, session, excel)

        # ════════════════════════════════════════════════════════
        # ETAPA EXTRA — MM03 + BASE LINHA
        # ════════════════════════════════════════════════════════
        aplicar_mm03_pesos(wb_base, session, excel)

        # ════════════════════════════════════════════════════════
        # ETAPA 6 — TABELA DINÂMICA
        # ════════════════════════════════════════════════════════
        marcar_etapa("Tabela Dinâmica", "ativo")
        set_status("Atualizando Tabelas Dinâmicas…")
        registrar_log("RefreshAll…")
        wb_base.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(4)
        marcar_etapa("Tabela Dinâmica", "concluido")
        registrar_log("Tabelas Dinâmicas atualizadas.", "destaque")

        # ════════════════════════════════════════════════════════
        # ETAPA 7 — SALVAMENTO
        # ════════════════════════════════════════════════════════
        marcar_etapa("Salvamento", "ativo")
        set_status("Salvando…")
        registrar_log(f"Salvando: {arquivo_novo}")
        wb_base.Save()
        excel.DisplayAlerts = True

        # Fecha o workbook antes de editar o ZIP
        wb_base.Close(False)
        wb_base = None

        try: os.remove(caminho_prod)
        except: pass

        # ── Corrige o range do PivotCache direto no XML ──────────
        # O PivotCache armazena um range fixo (ex: B1:N31209).
        # O RefreshAll via COM NÃO atualiza esse atributo no arquivo
        # salvo. A única forma segura de corrigi-lo sem destruir
        # segmentações é editar o XML dentro do ZIP do xlsx.
        set_status("Corrigindo range do PivotCache no XML…")
        registrar_log("Corrigindo range fixo do PivotCache no XML do xlsx…")

        _corrigido = _corrigir_pivot_cache_xml(arquivo_novo, "MB51- PRODPLANTA")
        if _corrigido:
            registrar_log(
                f"PivotCache atualizado: {_corrigido}", "destaque"
            )
        else:
            registrar_log(
                "PivotCache não encontrado no XML — verifique manualmente.", "aviso"
            )

        # ── Executa "Atualizar Tudo" após modificar o XML ─────────
        try:
            set_status("Executando Atualizar Tudo na guia Dados…")
            registrar_log("Reabrindo arquivo para aplicar a atualização final das tabelas…")
            
            wb_final = excel.Workbooks.Open(arquivo_novo)
            wb_final.RefreshAll()
            
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                time.sleep(3)
                
            # 2. SEGMENTAÇÃO DE DADOS “DATA LANÇAMENTO”
            try:
                slicer_cache = None
                for sc in wb_final.SlicerCaches:
                    if "data" in sc.Name.lower() and "lan" in sc.Name.lower():
                        slicer_cache = sc
                        break
                        
                if not slicer_cache:
                    for sc in wb_final.SlicerCaches:
                        for sl in sc.Slicers:
                            if "data lançamento" in sl.Caption.lower() or "data lancamento" in sl.Caption.lower():
                                slicer_cache = sc
                                break
                        if slicer_cache:
                            break
                            
                if slicer_cache:
                    registrar_log("OK: Segmentação \"Data Lançamento\" localizada.")
                    
                    items = slicer_cache.SlicerItems
                    qte_itens = items.Count
                    registrar_log(f"OK: {qte_itens} itens de data identificados.")
                    
                    if qte_itens >= 2:
                        idx_penultimo = qte_itens - 1
                        item_penultimo = items(idx_penultimo)
                        registrar_log(f"OK: Penúltimo item identificado: {item_penultimo.Name}")
                        
                        slicer_cache.ClearManualFilter()
                        items(idx_penultimo).Selected = True
                        
                        for i in range(1, qte_itens + 1):
                            if i != idx_penultimo:
                                items(i).Selected = False
                                
                        registrar_log("OK: Selecionando penúltimo item da segmentação \"Data Lançamento\".")
                    else:
                        registrar_log("ERRO: A segmentação \"Data Lançamento\" possui menos de 2 itens válidos. Não é possível selecionar o penúltimo item.", "erro")
                else:
                    registrar_log("ERRO: Segmentação \"Data Lançamento\" não localizada.", "erro")
            except Exception as e_slicer:
                registrar_log(f"ERRO: Falha ao manipular segmentação: {e_slicer}", "erro")
                
            wb_final.Save()
            wb_final.Close(False)
            wb_final = None
            registrar_log("✓ Atualizar Tudo executado com sucesso!", "destaque")
        except Exception as e_atualiza:
            registrar_log(f"⚠ Aviso na atualização final: {e_atualiza}", "aviso")

        marcar_etapa("Salvamento", "concluido")
        set_status("Concluído com sucesso!", cor=ACCENT)
        set_btn("done")
        registrar_log("━" * 60, "sep")
        registrar_log("PROCESSO CONCLUÍDO COM SUCESSO.", "destaque")
        registrar_log(f"Arquivo: {arquivo_novo}", "destaque")

        notificar_windows(
            "✅ Automação de Produção Concluída",
            f"Planilha salva em:\n{dest['rel_destino']}",
            arquivo=arquivo_novo,
        )

        if "--auto" in sys.argv:
            registrar_log("Encerrando aplicativo automaticamente em 5s…", "destaque")
            root.after(5000, root.destroy)
        else:
            root.after(0, lambda: messagebox.showinfo(
                "Concluído",
                f"Planilha salva em:\n\n{arquivo_novo}",
                parent=root,
            ))

    except Exception:
        tb = traceback.format_exc()
        registrar_log(f"ERRO CRÍTICO:\n{tb}", "erro")
        set_status("Erro — veja o terminal.", cor=TEXT_ERROR)
        set_btn("erro")

        # Marca etapa ativa como erro
        for nome in ETAPAS:
            dot, _ = step_widgets[nome]
            if dot.cget("fg") == ACCENT:
                marcar_etapa(nome, "erro")
                break

        notificar_windows(
            "❌ Automação de Produção — ERRO",
            "Ocorreu um erro crítico. Verifique o terminal do Prod Forge.",
        )

        if "--auto" not in sys.argv:
            root.after(0, lambda: messagebox.showerror(
                "Erro na Execução", tb[:1000], parent=root
            ))

    finally:
        # ── Cleanup: fecha arquivos abertos e mata instância zumbi do Excel ──
        try:
            if excel:
                excel.DisplayAlerts = False
                # Fecha todos os workbooks ainda abertos
                for _wb in list(excel.Workbooks):
                    try:
                        _wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                try:
                    excel.DisplayAlerts = True
                    excel.Quit()
                except Exception:
                    pass
                excel = None
        except Exception:
            pass
        # Garante que a referência COM seja liberada
        wb_base = None

def iniciar_execucao():
    if _info_base is None:
        messagebox.showerror("Erro", "Planilha base não detectada.", parent=root)
        return
    set_btn("running")
    t = threading.Thread(target=executar_automacao, daemon=True)
    t.start()


btn_executar.config(command=iniciar_execucao)


# ══════════════════════════════════════════════════════════════════
#  DETECÇÃO INICIAL (roda em thread para não bloquear UI)
# ══════════════════════════════════════════════════════════════════
def detectar_planilha():
    global _info_base
    try:
        info = localizar_planilha_base(_mes_selecionado, _ano_selecionado)
        _info_base = info
        root.after(0, lambda: atualizar_card_info(info))
        root.after(0, lambda: set_status(
            f"Base detectada: {os.path.basename(info['arquivo_base'])}"
        ))
        # Atualiza o label de range na UI
        root.after(0, lambda: lbl_range_datas.config(
            text=f"Range MB51: {info['data_inicio'].strftime('%d/%m/%Y')} → {info['data_fim'].strftime('%d/%m/%Y')}"
        ))
        registrar_log(f"Base detectada: {info['arquivo_base']}", "destaque")
        registrar_log(
            f"Referência: pasta {info['dia_ref']}  →  "
            f"Destino: {datetime.today().strftime('%d.%m')}  |  "
            f"Range MB51: {info['data_inicio'].strftime('%d/%m/%Y')} → {info['data_fim'].strftime('%d/%m/%Y')}"
        )

        if "--auto" in sys.argv:
            registrar_log("Modo Automático (--auto) ativo. Iniciando execução em 2 segundos…", "destaque")
            root.after(2000, iniciar_execucao)
    except FileNotFoundError as e:
        msg = str(e)
        _info_base = None
        root.after(0, lambda: card_erro(msg))
        root.after(0, lambda: set_status("Planilha não encontrada.", cor=TEXT_ERROR))
        root.after(0, lambda: set_btn("normal"))
        registrar_log(f"ERRO na detecção: {msg}", "erro")
    except Exception:
        tb = traceback.format_exc()
        _info_base = None
        root.after(0, lambda: set_status("Erro na detecção.", cor=TEXT_ERROR))
        registrar_log(f"ERRO na detecção:\n{tb}", "erro")


threading.Thread(target=detectar_planilha, daemon=True).start()

if __name__ == "__main__":
    root.mainloop()
