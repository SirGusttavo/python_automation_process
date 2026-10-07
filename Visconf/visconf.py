"""
visconf_engine_v4.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
VISCONF Engine v4.2.0

Correções sobre v2/v4.0:
  1.  moda_pallet(): empate ou todos diferentes → usa MAX (não mediana)
      Justificativa: pallet fechado → reservar pelo maior lote padrão
  2.  Reserva preventiva: quando pool_restante < 1 pallet após consumo
      → gera RESERVAR (1 pallet), mesmo que NL=0 (ponto 10)
  3.  Todo RESERVAR obrigatoriamente tem pallet calculado
      Se sem moda → pallet=None + alerta específico (ponto 11)
  4.  Distinção clara: material não encontrado ≠ material com qty=0 (ponto 12)
  5.  Data dd.mm.aaaa (ponto separador) em todas as saídas (ponto 6)
  6.  Nome do arquivo: "Reserva de Insumos - dd.mm.xlsx" (ponto 7)
  7.  Pipeline simplificado — sem etapas redundantes (ponto 2):
        Push Produção → Push MP01 → Desconto → Gravar MFR
        → Gravar NA → Relatório
  8.  Salvamento automático obrigatório do VISCONF após o término do pipeline
      Usa wb.save() ou wb.api.Save() automaticamente sem perguntar ao usuário.
  9.  Formatação numérica exata "0.00" nas colunas NB/Estoque/Consumo/NL (ponto 5)
  10. Und. Pallet = qtd_padrao real (moda tipo-50 da MP01), nunca placeholder
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import io, json, math, os, re, threading, warnings
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import multimode

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

CONFIG_FILE     = Path.home() / ".visconf_engine_v4.json"

TIPOS_IGNORADOS = {"PAPEL", "ENERGIA", "OUTROS"}

STATUS_OK           = "OK"
STATUS_RESERVAR     = "RESERVAR"
STATUS_ZERADO       = "SEM ESTOQUE"
STATUS_NAO_ENCONTRADO = "SEM ESTOQUE"

DEP_MP06 = "MP06"
DEP_MP05 = "MP05"
DEP_MP01 = "MP01"
MP01_TIPO_PALLET = {"50", "51"}   # tipo 51 = lotes de pallet físico na exportação LX02

# Depósito de recebimento por categoria
# [REVIEW] Category names (PROFESSIONAL, FORMATADOS, PH, DURAMAX) flagged for review
CATEGORIA_DEP = {
    "PROFESSIONAL": DEP_MP06,
    "FORMATADOS":   DEP_MP06,
    "PH":           DEP_MP05,
    "DURAMAX":      DEP_MP05,
}

CATEGORIA_LINHAS = {
    "PROFESSIONAL": ["PROF.01", "PROF.02", "PROF.03", "PROF.04"],
    "FORMATADOS":   ["FOR.01", "FOR.02", "FOR.03"],
    "PH":           ["PH.04", "PH.06", "PH.07","PH.08", "PH.12", "PH.13", "PH.14", "PH.15"],
    "DURAMAX":      ["DMX.05"],
}

SHEET_CONFIG = {
    "PROFESSIONAL": {"targets":{"PROF.01","PROF.02","PROF.03","PROF.04"}, "is_lcode":False, "categoria":"PROFESSIONAL"},
    "FORMATADOS":   {"targets":{"FOR.01","FOR.02","FOR.03"},               "is_lcode":False, "categoria":"FORMATADOS"},
    "PH":           {"targets":set(), "is_lcode":True,                     "categoria":"PH"},
    "DURAMAX":      {"targets":set(), "is_lcode":True,                     "categoria":"DURAMAX"},
}
SHEET_FALLBACK = {"row_header":8,"col_linha":2,"col_sku":4,"col_ordem":18}

MACHINE_MAP = {
    "PROF.01":"PROF.01", "PROF.02":"PROF.02", "PROF.03":"PROF.03", "PROF.04":"PROF.04",
    "FOR.01":"FOR.01",   "FOR.02":"FOR.02",   "FOR.03":"FOR.03",
    "PH.04":"PH.04",     "PH.06":"PH.06",     "PH.07":"PH.07",     "PH.08":"PH.08",
    "PH.12":"PH.12",     "PH.13":"PH.13",     "PH.14":"PH.14",     "PH.15":"PH.15",
    "DMX.05":"DMX.05"
}
STOP_WORDS  = {"Produzido","Projeção"}
SKIP_WORDS  = {"Total da linha >>","Parada prevista (h) >>",
               "Quantidade de setups >>",">>> TURNOS","Tipo de parada >>","Query","↓"}
MAX_HDR_R, MAX_HDR_C = 15, 30

LINE_ORDER_BASE = [
    "PROF.01", "PROF.02", "PROF.03", "PROF.04",
    "FOR.01", "FOR.02", "FOR.03",
    "PH.04", "PH.06", "PH.07", "PH.08", "PH.12", "PH.13", "PH.14", "PH.15",
    "DMX.05"
]

# ── CORES ─────────────────────────────────────────────────────────────────────
C_BG="1C1C1E";C_PANEL="2C2C2E";C_ACCENT="30D158";C_ACCENT2="0A84FF"
C_WARN="FF9F0A";C_ERR="FF453A";C_TEXT="F2F2F7";C_MUTED="8E8E93";C_BTN="3A3A3C"
def hx(c): return f"#{c}"

XL_HDR_RES="1E3A5F";XL_HDR_AV="4A1942";XL_HDR_SUM="1A3C2B"
XL_CRITICO="FDECEA";XL_ALERTA="FFF3CD";XL_OK_F="E9F7EF"
XL_ALT_B="EBF3FB";XL_ALT_P="F5EBF5";XL_WHITE="FFFFFF"


# ══════════════════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ══════════════════════════════════════════════════════════════════════════════

def normalizar_material(valor):
    if valor is None: return ""
    if isinstance(valor, (int, float)):
        return str(int(valor)) if int(valor) == valor else str(valor).upper().strip()
    s = str(valor).strip()
    try:
        return str(int(float(s)))
    except ValueError:
        return s.upper()

def normalize_linha(raw):
    s = str(raw).strip()
    if s in MACHINE_MAP: return MACHINE_MAP[s]
    m = re.fullmatch(r"L(\d+)", s, re.IGNORECASE)
    if m: return f"LINHA {int(m.group(1)):02d}"
    return s

def is_lcode(s): return bool(re.fullmatch(r"L\d+", str(s).strip(), re.IGNORECASE))

def linha_para_categoria(linha):
    for cat, linhas in CATEGORIA_LINHAS.items():
        if linha in linhas: return cat
    return "PROFESSIONAL"

def line_sort_key(l):
    if l in LINE_ORDER_BASE: return (0, LINE_ORDER_BASE.index(l), "")
    m = re.fullmatch(r"LINHA (\d+)", l)
    if m: return (1, int(m.group(1)), "")
    return (2, 0, l)

def fmt_data(dt) -> str:
    """Formata data como dd/mm/aaaa."""
    if dt is None: return ""
    if isinstance(dt, str): return dt
    if hasattr(dt, 'strftime'): return dt.strftime("%d/%m/%Y")
    return str(dt)

def safe_float(v, default=0.0):
    try: return float(v) if v is not None else default
    except: return default

def _thin(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(h): return PatternFill("solid", fgColor=h)
def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")
def _center(): return Alignment(horizontal="center", vertical="center")
def _right():  return Alignment(horizontal="right",  vertical="center")

def load_config():
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except: pass
    return {}

def save_config(data):
    try: CONFIG_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except: pass


# ══════════════════════════════════════════════════════════════════════════════
# MODA ESTATÍSTICA DE PALLET (corrigida)
# ══════════════════════════════════════════════════════════════════════════════

def moda_pallet(qtds: list) -> float:
    """
    Regra definitiva de quantidade padrão por pallet:
      1. Calcula a moda estatística dos lotes tipo-50 da MP01.
      2. Se moda única → usa esse valor.
      3. Se empate entre múltiplos valores → usa o MAIOR (objetivo: pallet fechado).
      4. Se todos diferentes (sem repetição) → usa o MAIOR valor.
         Justificativa: queremos reservar pallet fechado; o maior é o padrão logístico
         e os menores são pallets quebrados/fracionados.
    """
    if not qtds: return 0.0
    modos = multimode(qtds)
    if len(modos) == 1:
        return modos[0]          # moda clara
    return max(modos)             # empate → maior; todos diferentes → multimode = todos → maior


# ══════════════════════════════════════════════════════════════════════════════
# LEITURA DE ESTOQUE
# ══════════════════════════════════════════════════════════════════════════════

def ler_mp01(ws_mp01):
    """
    Retorna:
      estoque_mp01: {material: {"qty": float, "umb": str}}
      peso_mp01:    {material: {"qtd_padrao": float, "umb": str, "n_pallets": int}}

    Estrutura real da MP01 (exportação LX02):
      [0] (A) = Tipo de depósito
      [1] (B) = Material (pode estar vazio)
      [2] (C) = Texto breve material (pode conter o código)
      [3] (D) = Posição no depósito
      [4] (E) = Estoque disponível (qty)
      [5] (F) = UMB (pode estar embutido na qty)
      [6] (G) = Unidade de depósito
      [7] (H) = Data
    """

    lotes_t50: dict[str,list] = defaultdict(list)
    est_total: dict[str,dict] = defaultdict(lambda: {"qty":0.0,"umb":""})
    
    # ── Mapeamento FIXO de Colunas (Layout LX02 nunca muda) ──
    # A exportação do SAP pela LX02 possui a seguinte estrutura fixa:
    #   [0] (A) = Tipo de depósito
    #   [1] (B) = Material
    #   [2] (C) = Texto breve material (Descrição)
    #   [3] (D) = Posição no depósito
    #   [4] (E) = Estoque disponível (qty)
    #   [5] (F) = UME (unidade de medida)
    #   [6] (G) = Unidade de depósito (número gigante — NÃO usar como estoque!)
    #   [7] (H) = Data do vencimento
    c_tipo = 0
    c_mat  = 1
    c_desc = 2
    c_qty  = 4
    c_umb  = 5

    for row in _xw_iter_rows(ws_mp01, min_row=2):
        if not row or not any(row): continue
        
        val_tipo = row[c_tipo] if len(row) > c_tipo else None
        if val_tipo is None: continue
        tipo = str(val_tipo).strip()
        if tipo.endswith(".0"): tipo = tipo[:-2]
        if not tipo: continue
        
        mat_str = str(row[c_mat]).strip() if len(row) > c_mat and row[c_mat] else ""
        if not mat_str and len(row) > c_desc and row[c_desc]:
            mat_str = str(row[c_desc]).strip().split(" ")[0]
        mat = normalizar_material(mat_str)
        if not mat or mat == "NONE": continue
        
        qty_str = str(row[c_qty]).strip() if len(row) > c_qty and row[c_qty] else "0"
        qty_val = safe_float(qty_str.replace("KG", "").replace("UN", "").replace("PC", "").replace(",", ".").strip())
        umb = str(row[c_umb]).strip() if len(row) > c_umb and row[c_umb] else ""
        if not umb and " " in qty_str:
            umb = qty_str.split(" ")[-1]
            
        qty_val = abs(qty_val)
            
        if qty_val > 0:
            est_total[mat]["qty"] += qty_val
            est_total[mat]["umb"] = umb
        if tipo in MP01_TIPO_PALLET and qty_val > 0:
            lotes_t50[mat].append(qty_val)

    peso = {}
    for mat, qtds in lotes_t50.items():
        md = moda_pallet(qtds)
        if md > 0:
            peso[mat] = {"qtd_padrao":md, "umb":est_total[mat]["umb"], "n_pallets":len(qtds)}
    return dict(est_total), peso


def ler_estoque_deposito(ws, nome_aba) -> dict:
    """Aba de estoque de linha (MP06 ou MP05).
    MP06: Coluna J (index 9) = Total Estoque
    MP05: Coluna I (index 8) = Total Geral
    Sem deduzir RNC conforme baseline do VBA.
    """
    est = defaultdict(float)
    for row in _xw_iter_rows(ws, min_row=2):
        mat = normalizar_material(row[1])
        if not mat or mat=="None": continue
        
        # J=9 para MP06, I=8 para MP05
        if nome_aba == "MP06":
            qty = safe_float(row[9])
        elif nome_aba == "MP05":
            qty = safe_float(row[8])
        else:
            qty = 0.0
        
        # Valores negativos = insumos sem lote/HU → considerar positivo
        qty = abs(qty)
            
        if qty > 0:
            est[mat] += qty
    return dict(est)


def construir_estoques(wb_vis, categorias_ativas, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)
    r = {"MP06":{},"MP05":{},"MP01_total":{},"MP01_pallet":{}}

    sheetnames = [s.name for s in wb_vis.sheets]
    if "MP06" in sheetnames and ("PROFESSIONAL" in categorias_ativas or "FORMATADOS" in categorias_ativas):
        r["MP06"] = ler_estoque_deposito(wb_vis.sheets["MP06"], "MP06")
        _log(f"  MP06: {len(r['MP06'])} materiais")

    if "MP05" in sheetnames and ("PH" in categorias_ativas or "DURAMAX" in categorias_ativas):
        r["MP05"] = ler_estoque_deposito(wb_vis.sheets["MP05"], "MP05")
        _log(f"  MP05: {len(r['MP05'])} materiais")

    if "MP01" in sheetnames:
        r["MP01_total"], r["MP01_pallet"] = ler_mp01(wb_vis.sheets["MP01"])
        _log(f"  MP01: {len(r['MP01_total'])} materiais | {len(r['MP01_pallet'])} com padrão de pallet")

    return r


# ══════════════════════════════════════════════════════════════════════════════
# FASE 1 — PLAN READER
# ══════════════════════════════════════════════════════════════════════════════

def _detect_cols(rows, log_fn=None):
    found = {"row_header":None,"col_linha":None,"col_sku":None,"col_ordem":None}
    for ri in range(min(MAX_HDR_R,len(rows))):
        for ci in range(min(MAX_HDR_C,len(rows[ri]))):
            v = rows[ri][ci]
            if v is None: continue
            t = str(v).strip().lower()
            if t=="linha"  and found["col_linha"]  is None: found["col_linha"]=ci;  found["row_header"]=ri
            if t=="sku"    and found["col_sku"]    is None: found["col_sku"]=ci
            if t=="ordem"  and found["col_ordem"]  is None: found["col_ordem"]=ci
        if all(found[k] is not None for k in ("col_linha","col_sku","col_ordem")): break
    if any(found[k] is None for k in ("col_linha","col_sku","col_ordem")):
        if log_fn: log_fn("  ⚠️  Cabeçalho não detectado — usando fallback")
        return {**SHEET_FALLBACK,"auto_detected":False}
    return {**found,"auto_detected":True}

def _build_date_map(rows, row_hdr, col_start, dates_filter):
    dm = {}
    hdr = rows[row_hdr]
    for ci in range(col_start, len(hdr)):
        v = hdr[ci]
        if v is not None and hasattr(v,"year"):
            try:
                d = v.date()
                if d in dates_filter and d not in dm and 2024<=d.year<=2028:
                    dm[d] = ci
            except: continue
    return dm

def _extract_sheet(wb_do, sheet_name, cfg, dates_filter, linhas_ativas, log_fn=None):
    if sheet_name not in wb_do.sheetnames: return []
    ws   = wb_do[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    cols = _detect_cols(rows, log_fn)
    row_hdr = cols["row_header"]
    col_l, col_s, col_o = cols["col_linha"], cols["col_sku"], cols["col_ordem"]
    dm = _build_date_map(rows, row_hdr, col_o+1, dates_filter)
    if not dm:
        if log_fn: log_fn(f"  ⚠️  Sem datas em '{sheet_name}'")
        return []
    if log_fn: log_fn(f"  📅 {', '.join(d.strftime('%d/%m') for d in sorted(dm))}")

    targets   = cfg["targets"]
    use_lcode = cfg["is_lcode"]
    categoria = cfg["categoria"]
    records   = []

    for ri in range(row_hdr+1, len(rows)):
        row = rows[ri]
        if len(row)<=col_l or row[col_l] is None: continue
        s = str(row[col_l]).strip()
        if s in STOP_WORDS: break
        if s in SKIP_WORDS or s.replace(".","",1).isdigit() or s in ("0",""): continue
        if use_lcode:
            if not is_lcode(s): continue
        else:
            if s not in targets: continue
        linha_norm = normalize_linha(s)
        if linhas_ativas and linha_norm not in linhas_ativas: continue
        sku_raw = row[col_s] if len(row)>col_s else None
        if sku_raw is None: continue
        sku = normalizar_material(sku_raw)
        o_raw = row[col_o] if len(row)>col_o else None
        ordem = str(o_raw).split(".")[0].strip() if o_raw is not None else ""
        if not sku or sku in ("nan","None"): continue
        for d, ci in dm.items():
            raw = row[ci] if ci<len(row) else None
            try:   qty = float(raw) if raw is not None else 0.0
            except: qty = 0.0
            if qty>0:
                records.append({
                    "data":      datetime.combine(d, datetime.min.time()),
                    "linha":     linha_norm,
                    "sku":       sku,
                    "ordem":     ordem,
                    "qt":        int(qty),
                    "categoria": categoria,
                    "dep_recep": CATEGORIA_DEP[categoria],
                })
    return records

def fase1_plan_reader(plano_path, dates_filter, categorias_ativas, linhas_ativas, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)
    _log("  📖 Abrindo plano (data_only=True)...")
    if isinstance(plano_path,(bytes,bytearray)): bio = io.BytesIO(plano_path)
    elif isinstance(plano_path, io.BytesIO): plano_path.seek(0); bio=plano_path
    else: bio = str(plano_path)
    wb = openpyxl.load_workbook(bio if isinstance(bio,str) else bio, read_only=True, data_only=True)
    all_records = []
    for sheet_name, cfg in SHEET_CONFIG.items():
        if cfg["categoria"] not in categorias_ativas: continue
        _log(f"  📋 Aba '{sheet_name}' [{cfg['categoria']}]...")
        recs = _extract_sheet(wb, sheet_name, cfg, dates_filter, linhas_ativas, log_fn)
        _log(f"     ✅ {len(recs)} registro(s)")
        all_records.extend(recs)
    
    # IMPORTANTE: fechar o workbook para liberar o arquivo!
    wb.close()
    
    all_records.sort(key=lambda r: (r["data"], line_sort_key(r["linha"])))
    return all_records


# ══════════════════════════════════════════════════════════════════════════════
# FASE 2 — DESCONTO DE PRODUÇÃO
# ══════════════════════════════════════════════════════════════════════════════

def fase2_desconto(registros, ws_prod, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)
    
    prod = defaultdict(float)
    from datetime import datetime, date
    for row in _xw_iter_rows(ws_prod, min_row=2):
        mat = normalizar_material(row[2])
        qty = safe_float(row[4])
        dt_val = row[6]
        if not mat or qty <= 0: continue
        
        if hasattr(dt_val, "date"):
            d = dt_val.date()
        elif isinstance(dt_val, str):
            try:
                d = datetime.strptime(dt_val.strip(), "%d.%m.%Y").date()
            except:
                d = str(dt_val).strip()
        else:
            d = str(dt_val)
            
        prod[(d, mat)] += qty

    _log(f"  {len(prod)} chaves (Data+Material) com produção")
    
    for reg in registros:
        mat_plan = normalizar_material(reg["sku"])
        dt_plan = reg["data"]
        
        if hasattr(dt_plan, "date"):
            d_plan = dt_plan.date()
        else:
            d_plan = dt_plan
            
        chave = (d_plan, mat_plan)
        qt_orig = reg["qt"]
        
        produzido = prod.get(chave, 0.0)
        
        if produzido > 0:
            desconto = min(qt_orig, produzido)
            reg["qt_efetiva"] = max(0.0, qt_orig - desconto)
            reg["qt_produzida"] = desconto
            prod[chave] -= desconto  # Decrementa para a próxima ordem (se houver)
            _log(f"  Prod {mat_plan} | {reg['linha']:12s} | Plan:{qt_orig:.0f} → Prod_Disp:{produzido:.0f} → Efetivo:{reg['qt_efetiva']:.0f}")
        else:
            reg["qt_efetiva"] = float(qt_orig)
            reg["qt_produzida"] = 0.0
            
    return registros


# ══════════════════════════════════════════════════════════════════════════════
# FASE 3+4 — EXPLOSÃO + STOCK MANAGER + PALLET ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def _ler_bom(ws_bd):
    bom = defaultdict(list)
    for row in _xw_iter_rows(ws_bd, min_row=2):
        linha_bd  = str(row[0]).strip()  if row[0]  else ""
        sku_bd    = normalizar_material(row[1])
        insumo    = normalizar_material(row[13])
        descricao = str(row[14]).strip() if row[14] else ""
        tipo      = str(row[16]).strip().upper() if row[16] else ""
        um        = str(row[17]).strip() if row[17] else ""
        qtd_p1    = safe_float(row[19])
        if not sku_bd or not insumo or qtd_p1==0.0: continue
        if tipo in TIPOS_IGNORADOS: continue
        bom[(linha_bd,sku_bd)].append({"insumo":insumo,"descricao":descricao,
                                        "tipo":tipo,"um":um,"qtd_p1":qtd_p1})
    return dict(bom)


def _calcular_pallets(nl, insumo, peso_mp01, sem_pp_log, log_fn=None):
    """
    Calcula qtd de pallets para um material.
    *STAND BY* (pedido do usuário): não calcula a quantidade de pallets,
    retorna None. Apenas repassa a qtd_padrao para fins informativos.
    """
    if nl <= 0: return None, None, None
    pp = peso_mp01.get(insumo)
    if pp and pp["qtd_padrao"] > 0:
        return None, pp["qtd_padrao"], None
    # Sem padrão de pallet
    if insumo not in sem_pp_log:
        if log_fn: log_fn(f"  ⚠️  Sem padrão pallet (MP01 tipo-50): {insumo}")
        sem_pp_log.add(insumo)
    return None, None, "SEM_PP"


def fase3_4_explosao(registros, ws_bd, estoques, log_fn=None):
    """
    Explosão BOM + Stock Manager + Pallet Engine.

    Regras implementadas:
      • Pool de estoque por depósito (MP06/MP05), decrementada sequencialmente.
      • Todo RESERVAR deve ter pallet calculado (se sem moda → alerta específico).
      • Reserva PREVENTIVA: quando pool_restante < 1 pallet padrão após consumo
        e o status seria OK → muda para RESERVAR (1 pallet preventivo).
      • Distinção entre "não encontrado no depósito" e "quantidade zero".
    """
    def _log(m):
        if log_fn: log_fn(m)

    bom         = _ler_bom(ws_bd)
    peso_mp01   = estoques.get("MP01_pallet", {})

    pool = {
        DEP_MP06: dict(estoques.get("MP06", {})),
        DEP_MP05: dict(estoques.get("MP05", {})),
    }
    mp01_total = estoques.get("MP01_total", {})

    _log(f"  BOM: {sum(len(v) for v in bom.values())} componentes")
    _log(f"  MP06: {len(pool[DEP_MP06])} | MP05: {len(pool[DEP_MP05])} | MP01 pallet: {len(peso_mp01)}")

    resultados = []
    sem_bom    = set()
    sem_pp_log = set()

    # Acumuladores para avisos
    acc: dict = defaultdict(lambda: {
        "descricao":"","tipo":"","um":"","dep":"",
        "est_inicial":0.0,"nb_total":0.0,"consumo_total":0.0,"nl_total":0.0,
        "linhas_ok":[],"linhas_reservar":[],"linhas_sem_est":[],"linhas_nao_enc":[],
        "sem_pp":False,"nao_encontrado":False,"zero_estoque":False,
    })

    for reg in registros:
        qt = reg.get("qt_efetiva", reg["qt"])
        if qt < 0: continue

        linha    = reg["linha"]
        sku      = reg["sku"]
        data     = reg["data"]
        ordem    = reg["ordem"]
        dep_rec  = reg.get("dep_recep", DEP_MP06)

        componentes = bom.get((linha,sku)) or bom.get(("",sku)) or []
        if not componentes:
            if (linha,sku) not in sem_bom:
                _log(f"  ⚠️  BOM ausente: {linha} / {sku}")
                sem_bom.add((linha,sku))
            continue

        p = pool.setdefault(dep_rec, {})

        for comp in componentes:
            insumo    = comp["insumo"]
            descricao = comp["descricao"]
            tipo      = comp["tipo"]
            um        = comp["um"]
            qtd_p1    = comp["qtd_p1"]

            nb = qt * qtd_p1

            is_in_line_depot = insumo in p
            is_in_warehouse = insumo in mp01_total

            est_antes = p.get(insumo, 0.0)
            
            consumo = min(nb, est_antes)
            p[insumo] = est_antes - consumo
            nl = max(0.0, nb - consumo)

            mp01_qty = mp01_total.get(insumo, {}).get("qty", 0.0) if isinstance(mp01_total.get(insumo), dict) else mp01_total.get(insumo, 0.0)
            
            nao_encontrado_dep = False
            zero_estoque_dep = False
            estoque_insuficiente = False

            if nl == 0.0:
                # Se a necessidade líquida é zero, o estoque da linha ATENDEU tudo. 
                # Não importa o saldo da MP01, o status é OK.
                status = STATUS_OK
            else:
                # Se precisamos de material (nl > 0), verificamos se a MP01 consegue suprir
                status = STATUS_RESERVAR
                if not is_in_warehouse:
                    nao_encontrado_dep = True
                elif mp01_qty <= 0:
                    zero_estoque_dep = True
                elif mp01_qty < nl:
                    # Tem estoque na MP01, mas não é suficiente para cobrir toda a Necessidade Líquida (nl)
                    estoque_insuficiente = True

            # ── Qtd padrão de pallet ──────────────────────────────────────────
            pp_info = peso_mp01.get(insumo)
            qtd_padrao = pp_info["qtd_padrao"] if pp_info else None

            # ── Reserva PREVENTIVA (REMOVIDA a pedido do usuário) ───────────────
            preventiva = False

            # ── Cálculo de pallets agora é feito em agregar_pallets ───────────
            qtd_pallet = None
            alerta_pp  = None

            resultados.append({
                "data":            data,
                "ordem":           ordem,
                "linha":           linha,
                "sku":             sku,
                "insumo":          insumo,
                "descricao":       descricao,
                "tipo":            tipo,
                "nb":              round(nb, 2),
                "estoque_antes":   round(est_antes, 2),
                "consumo_estoque": round(consumo, 2),
                "nl":              round(nl, 2),
                "status":          status,
                "qtd_pallet":      qtd_pallet,
                "qtd_padrao":      qtd_padrao,
                "dep_recep":       dep_rec,
                "preventiva":      preventiva,
                "nao_encontrado":  nao_encontrado_dep,
                "zero_estoque":    zero_estoque_dep,
                "estoque_insuficiente": estoque_insuficiente,
                "sem_pp":          (alerta_pp == "SEM_PP"),
            })

            # ── Acumular avisos ───────────────────────────────────────────────
            a = acc[insumo]
            if not a["descricao"]:
                a["descricao"]=descricao; a["tipo"]=tipo; a["um"]=um; a["dep"]=dep_rec
                a["est_inicial"] = estoques.get(dep_rec,{}).get(insumo,0.0)
            a["nb_total"]      += nb
            a["consumo_total"] += consumo
            a["nl_total"]      += nl
            if nao_encontrado_dep and linha not in a["linhas_nao_enc"]:
                a["linhas_nao_enc"].append(linha); a["nao_encontrado"] = True
            elif zero_estoque_dep and linha not in a["linhas_sem_est"]:
                a["linhas_sem_est"].append(linha); a["zero_estoque"] = True
            elif status == STATUS_RESERVAR and linha not in a["linhas_reservar"]:
                a["linhas_reservar"].append(linha)
            elif status == STATUS_OK and linha not in a["linhas_ok"]:
                a["linhas_ok"].append(linha)
            if alerta_pp == "SEM_PP": a["sem_pp"] = True
            if estoque_insuficiente: a["estoque_insuficiente"] = True

    # ── Classificar avisos PCP ────────────────────────────────────────────────
    avisos = {}
    for ins, a in acc.items():
        est_final    = max(0.0, a["est_inicial"] - a["consumo_total"])
        cobertura    = (a["consumo_total"]/a["nb_total"]*100 if a["nb_total"]>0 else 100.0)
        pp_info      = peso_mp01.get(ins)

        if a["nao_encontrado"]:
            ta = "CRÍTICO - NÃO ENCONTRADO NA MP01"
        elif a["zero_estoque"]:
            ta = "CRÍTICO - SEM ESTOQUE DISPONÍVEL"
        elif a.get("estoque_insuficiente", False):
            ta = "CRÍTICO - ESTOQUE INSUFICIENTE"
        elif a["sem_pp"]:
            ta = "ALERTA - SEM PADRÃO DE PALLET"
        elif a["nl_total"] > 0:
            ta = "RESERVAR"
        else:
            ta = "OK"

        avisos[ins] = {
            **a,
            "est_final":      round(est_final, 2),
            "cobertura_pct":  round(cobertura, 1),
            "tipo_aviso":     ta,
            "qtd_padrao":     pp_info["qtd_padrao"] if pp_info else None,
            "n_pallets_mp01": pp_info["n_pallets"]  if pp_info else 0,
        }

    return resultados, dict(avisos)


# ══════════════════════════════════════════════════════════════════════════════
# FASE 4 — AGREGAÇÃO DE PALLETS POR (insumo, linha)
# ══════════════════════════════════════════════════════════════════════════════

def agregar_pallets(resultados: list, peso_mp01: dict, log_fn=None) -> tuple[list, dict]:
    """
    Calcula pallets POR LINHA (cada registro individualmente).

    Regras estritas:
      OK              → qtd_pallet = None, qtd_padrao = None
      RESERVAR        → qtd_pallet = ceil(nl / qtd_padrao), qtd_padrao = valor
                         Se item não encontrado no estoque → qtd_pallet = "SEM ESTOQUE"
                         Pallet limitado ao nº de pallets disponíveis em MP01
      SEM ESTOQUE     → qtd_pallet = "SEM ESTOQUE", qtd_padrao = None
    """
    def _log(m):
        if log_fn: log_fn(m)

    _log("\n  Calculando pallets por linha...")
    sem_pp = set()
    pallets_grupo = {}  # mantém compatibilidade de retorno
    pallets_usados_global = defaultdict(int)

    # Variáveis para regra de cálculo acumulado na corrida
    nl_acumulada = defaultdict(float)        # key: (insumo, linha)
    pallets_solicitados = defaultdict(int)   # key: (insumo, linha)

    for r in resultados:
        status = r["status"]
        nl     = r["nl"]
        insumo = r["insumo"]
        linha  = r["linha"]
        nao_enc = r.get("nao_encontrado", False)
        zero_est = r.get("zero_estoque", False)
        est_insuf = r.get("estoque_insuficiente", False)

        # ── OK = SEM PALLET + SEM QUANTIDADE DO PALLET ──
        if status == STATUS_OK:
            r["qtd_pallet"] = None
            r["qtd_padrao"] = None
            r["sem_pp"]     = False
            continue

        # ── ESTOQUE INSUFICIENTE (tem na MP01, mas não cobre a necessidade total) ──
        if est_insuf:
            r["status"]     = STATUS_RESERVAR
            r["qtd_pallet"] = "ESTOQUE INSUFICIENTE"
            r["qtd_padrao"] = "ESTOQUE INSUFICIENTE"
            r["sem_pp"]     = True
            _log(f"  {insumo} [{linha}]: ESTOQUE INSUFICIENTE")
            continue

        # ── SEM ESTOQUE (não encontrado ou zerado na MP01, e a linha precisa) ──
        if nao_enc or zero_est:
            r["status"]     = STATUS_RESERVAR  # "RESERVAR"
            r["qtd_pallet"] = "SEM ESTOQUE"
            r["qtd_padrao"] = "SEM ESTOQUE"
            r["sem_pp"]     = True
            _log(f"  {insumo} [{linha}]: SEM ESTOQUE")
            continue

        # ── RESERVAR com estoque ──
        pp = peso_mp01.get(insumo)

        if pp and pp["qtd_padrao"] > 0:
            qtd_padrao = pp["qtd_padrao"]
            n_disponiveis = pp.get("n_pallets", 999999)  # pallets disponíveis em MP01
            n_restantes = max(0, n_disponiveis - pallets_usados_global[insumo])

            if nl > 0:
                key_corrida = (insumo, linha)
                
                # 1. Somar nova necessidade à necessidade acumulada da corrida
                nl_acumulada[key_corrida] += nl
                
                # 2. Calcular pallets totais necessários para a quantidade acumulada
                pallets_totais = math.ceil(nl_acumulada[key_corrida] / qtd_padrao)
                
                # 3. Descobrir quantos pallets a mais precisamos pedir agora
                ja_solicitados = pallets_solicitados[key_corrida]
                pallets_adicionais = max(0, pallets_totais - ja_solicitados)
                
                # 4. Limitar ao nº de pallets físicos disponíveis no estoque
                if pallets_adicionais > n_restantes:
                    pallets_adicionais = n_restantes

                r["qtd_pallet"] = pallets_adicionais
                
                # 5. Atualizar os saldos de pallets usados
                pallets_solicitados[key_corrida] += pallets_adicionais
                pallets_usados_global[insumo] += pallets_adicionais
                
            else:
                r["qtd_pallet"] = None

            r["qtd_padrao"] = qtd_padrao
            r["sem_pp"]     = False
            _log(f"  {insumo} [{linha}]: NL={nl:.2f} (Acum: {nl_acumulada.get((insumo, linha), 0):.2f}) / Pallet={qtd_padrao} → {r['qtd_pallet']} pallet(s) adicionais")
        else:
            # Sem padrão de pallet
            r["qtd_pallet"] = "FALTA PADRÃO"
            r["qtd_padrao"] = "FALTA PADRÃO"
            r["sem_pp"]     = True
            if insumo not in sem_pp:
                _log(f"  ⚠️  Sem padrão pallet: {insumo} — {r.get('descricao','')[:30]}")
                sem_pp.add(insumo)

    return resultados, pallets_grupo


# ══════════════════════════════════════════════════════════════════════════════
# XLWINGS — ABRIR + SALVAR ROBUSTO
# ══════════════════════════════════════════════════════════════════════════════

def _xw_open(path, log_fn=None):
    import xlwings as xw
    def _log(m):
        if log_fn: log_fn(m)
    norm = os.path.normcase(os.path.abspath(path))
    name = Path(path).name

    # 1. Busca exata pelo caminho completo
    for ai in xw.apps:
        for book in ai.books:
            try: bpath = book.fullname
            except: continue
            if os.path.normcase(os.path.abspath(bpath)) == norm:
                _log("  📎 Conectando ao arquivo já aberto")
                ai.display_alerts = False   # suprime aviso de vínculos
                return book, None, False

    # 2. Busca pelo nome do arquivo (caso o caminho completo difira por encoding/OneDrive)
    candidatos = []
    for ai in xw.apps:
        for book in ai.books:
            try: bname = Path(book.fullname).name
            except: continue
            if bname.lower() == name.lower():
                candidatos.append((book, ai))
    if len(candidatos) >= 1:
        book, ai = candidatos[0]
        _log("  📎 Conectando pelo nome do arquivo")
        ai.display_alerts = False
        return book, None, False

    # 3. Se não encontrou, abre o arquivo no Excel via COM.
    _log(f"  🔓 Abrindo {name} via Excel COM...")
    
    app_viva = None
    # Procura uma instância viva do Excel para reaproveitar
    for a in xw.apps:
        try:
            # Toca na API pra garantir que o PID ainda existe e responde
            _ = a.version
            app_viva = a
            break
        except:
            pass
            
    criou_app = False
    if app_viva is None:
        app_viva = xw.App(visible=True, add_book=False)
        criou_app = True
        
    app_viva.display_alerts = False
    try:
        wb = app_viva.books.open(path, update_links=False)
    except Exception as e:
        if criou_app:
            try: app_viva.quit()
            except: pass
        raise Exception(f"Falha ao abrir '{name}': {e}")
        
    _log(f"  ✅ {name} aberto via COM.")
    return wb, (app_viva if criou_app else None), True


def _xw_save(wb, we_own, log_fn=None):
    """
    Salva as alterações obrigatoriamente (automático).
    Equivalente ao Ctrl+B / Ctrl+S do usuário — salva no mesmo arquivo, no mesmo local.
    """
    def _log(m):
        if log_fn: log_fn(m)
    try:
        wb.api.Save()
        _log("  💾 VISCONF salvo (automático)")
    except Exception as e:
        _log(f"  ⚠️  Não foi possível salvar automaticamente: {e}")
        _log("       Os dados foram gravados nas abas — salve manualmente (Ctrl+S).")


def _xw_iter_rows(ws, min_row=2):
    """
    Simula o comportamento de ws.iter_rows(min_row, values_only=True) do openpyxl para uma aba do xlwings.
    Utiliza used_range para garantir que 100% das linhas e colunas reais com dados sejam lidas.
    """
    ur = ws.used_range
    if not ur: return []
    
    # ur.row e ur.column dizem onde o used_range começa
    ur_first_row = ur.row
    ur_last_row = ur.row + ur.rows.count - 1
    
    if ur_last_row < min_row:
        return []
        
    ur_last_col = ur.column + ur.columns.count - 1
    
    # Sempre pega a partir da coluna 1 (A)
    dados = ws.range((min_row, 1), (ur_last_row, max(ur_last_col, 1))).value
    if dados is None: return []
    if not isinstance(dados, list): return [[dados]]
    if len(dados) > 0 and not isinstance(dados[0], list):
        if max(ur_last_col, 1) == 1: return [[v] for v in dados]
        else: return [dados]
    return dados


def _xw_gravar(wb, nome_aba, rows_data, num_cols, txt_cols_idx, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)
    if nome_aba not in [s.name for s in wb.sheets]:
        raise Exception(f"Aba '{nome_aba}' não encontrada.")
    ws  = wb.sheets[nome_aba]
    n   = len(rows_data)
    if n==0: _log(f"  ⚠️  Sem dados para '{nome_aba}'"); return
    lr  = 1+n
    last_col = get_column_letter(num_cols)
    ws.range(f"A2:{last_col}9999").clear_contents()

    # Pré-formatar colunas texto (evita conversão automática do Excel)
    for c in txt_cols_idx:
        ws.range(f"{get_column_letter(c)}2:{get_column_letter(c)}{lr}").number_format = "@"

    ws.range("A2").value = rows_data

    # Data col A → dd/mm/aaaa
    ws.range(f"A2:A{lr}").number_format = "DD/MM/AAAA"

    # Colunas numéricas NB/Estoque/Consumo/NL → exatamente 2 casas decimais
    # (índices passados como parâmetro)
    _log(f"  ✅ {n} linha(s) → '{nome_aba}'")


def gravar_planejamento_mfr(wb, registros, log_fn=None):
    ws = wb.sheets["PLANEJAMENTO MFR"]
    ws.range("A2:E9999").clear_contents()
    n = len(registros)
    if n == 0: return
    lr = 1 + n
    # Escreve dados PRIMEIRO (tipos Python corretos)
    rows = [[r["data"], r["linha"], r["sku"], r["ordem"], r["qt"]]
            for r in registros]
    ws.range("A2").value = rows
    # Formatos DEPOIS da escrita
    ws.range(f"A2:A{lr}").number_format = "DD/MM/AAAA"
    ws.range(f"C2:D{lr}").number_format = "@"   # SKU e Ordem como texto
    ws.range(f"E2:E{lr}").number_format = "0"
    if log_fn: log_fn(f"  ✅ {n} linha(s) → PLANEJAMENTO MFR")


def gravar_necessidade_automatica(wb, resultados, log_fn=None):
    """
    Grava NECESSIDADE AUTOMÁTICA via xlwings.

    Ordem correta: escreve os dados primeiro, aplica formatos depois.
    Isso garante que o Excel processe os tipos Python corretamente antes
    de aplicar qualquer número_format.

    Pallet: mostra o total agregado por (insumo, linha) em TODAS as linhas
    daquele par — o operador vê o total independente do dia consultado.
    """
    ws = wb.sheets["NECESSIDADE AUTOMÁTICA"]
    ws.range("A2:N9999").clear_contents()
    n = len(resultados)
    if n == 0: return
    lr = 1 + n

    rows = []
    for r in resultados:
        rows.append([
            r["data"],               # A — datetime (objeto Python real)
            str(r["ordem"]),         # B — texto
            str(r["linha"]),         # C — texto
            str(r["sku"]),           # D — texto
            str(r["insumo"]),        # E — texto
            str(r["descricao"]),     # F — texto
            str(r["tipo"]),          # G — texto
            float(r["nb"]),          # H — número real
            float(r["estoque_antes"]),   # I — número real
            float(r["consumo_estoque"]), # J — número real
            float(r["nl"]),          # K — número real
            str(r["status"]),        # L — texto
            r["qtd_pallet"],         # M — inteiro, string "SEM ESTOQUE"/"FALTA PADRÃO", ou None
            float(r["qtd_padrao"]) if isinstance(r["qtd_padrao"], (int, float)) else (r["qtd_padrao"] if r["qtd_padrao"] else None),  # N — número, string ou None
        ])

    # 1. ESCREVE PRIMEIRO (preserva tipos Python)
    ws.range("A2").value = rows

    # 2. FORMATOS DEPOIS DA ESCRITA
    # Data: dd/mm/aaaa com barra
    ws.range(f"A2:A{lr}").number_format = "DD/MM/AAAA"
    # Texto (evita auto-conversão pelo Excel) — M e N podem ter texto
    for col in ("B", "C", "D", "E", "F", "G", "L", "M", "N"):
        ws.range(f"{col}2:{col}{lr}").number_format = "@"
    # Numérico com exatamente 2 casas decimais
    for col in ("H", "I", "J", "K"):
        ws.range(f"{col}2:{col}{lr}").number_format = "#.##0,00"

    if log_fn: log_fn(f"  ✅ {n} linha(s) → NECESSIDADE AUTOMÁTICA")


# ══════════════════════════════════════════════════════════════════════════════
# GERAÇÃO DO RELATÓRIO EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def gerar_relatorio(resultados, avisos, registros, output_path, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)

    wb = openpyxl.Workbook()

    # Estatísticas
    total_ordens   = len({r["ordem"] for r in registros if r["ordem"]})
    total_comp     = len(resultados)
    n_reservar     = sum(1 for r in resultados if r["status"]==STATUS_RESERVAR)
    n_ok           = sum(1 for r in resultados if r["status"]==STATUS_OK)
    n_zerado       = sum(1 for r in resultados if r.get("zero_estoque"))
    n_nao_enc      = sum(1 for r in resultados if r.get("nao_encontrado"))
    n_insuf        = sum(1 for r in resultados if r.get("estoque_insuficiente"))
    n_preventivo   = sum(1 for r in resultados if r.get("preventiva"))
    n_sem_pp       = sum(1 for r in resultados if r.get("sem_pp") and r["status"]==STATUS_RESERVAR)
    total_pallets  = sum(r["qtd_pallet"] for r in resultados if r["status"]==STATUS_RESERVAR and isinstance(r.get("qtd_pallet"), (int, float)))

    criticos = [(ins,av) for ins,av in avisos.items() if "CRÍTICO" in av.get("tipo_aviso","")]
    sem_pp_av = [(ins,av) for ins,av in avisos.items() if "PADRÃO" in av.get("tipo_aviso","")]

    # ── ABA SUMÁRIO ───────────────────────────────────────────────────────────
    ws_s = wb.active; ws_s.title = "SUMÁRIO"
    ws_s.sheet_view.showGridLines = False

    def _row(r, label, valor, fl=XL_ALT_B, fv=XL_WHITE):
        c1 = ws_s.cell(r,1,label); c1.fill=_fill(fl); c1.font=_font(size=10)
        c1.border=_thin(); c1.alignment=Alignment(horizontal="left",indent=1)
        ws_s.merge_cells(f"A{r}:C{r}")
        c2 = ws_s.cell(r,4,valor); c2.fill=_fill(fv); c2.font=_font(bold=True,size=10)
        c2.border=_thin(); c2.alignment=_center()
        ws_s.merge_cells(f"D{r}:F{r}")
        ws_s.row_dimensions[r].height=18

    def _hdr(r, t, fill="1A3C2B"):
        c = ws_s.cell(r,1,t)
        c.fill=_fill(fill); c.font=_font(bold=True,color="FFFFFF",size=11)
        c.border=_thin(); c.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        ws_s.merge_cells(f"A{r}:F{r}"); ws_s.row_dimensions[r].height=22

    rn=1
    _hdr(rn, f"VISCONF — SUMÁRIO EXECUTIVO   {fmt_data(datetime.now())}"); rn+=2
    _hdr(rn, "PLANEJAMENTO","1A3C2B"); rn+=1
    _row(rn,"Ordens",total_ordens); rn+=1
    _row(rn,"Componentes analisados",total_comp); rn+=2

    _hdr(rn,"NECESSIDADE","1A3C2B"); rn+=1
    _row(rn,"RESERVAR",n_reservar,fl=XL_ALT_B); rn+=1
    _row(rn,"OK (estoque suficiente)",n_ok,fl=XL_OK_F); rn+=1
    _row(rn,"SEM ESTOQUE (qty=0)",n_zerado,fl=XL_CRITICO); rn+=1
    _row(rn,"NÃO ENCONTRADO na MP01",n_nao_enc,fl=XL_CRITICO); rn+=1
    _row(rn,"ESTOQUE INSUFICIENTE",n_insuf,fl=XL_ALERTA); rn+=1
    _row(rn,"Reservas preventivas (resíduo < 1 pallet)",n_preventivo,fl=XL_ALERTA); rn+=1
    _row(rn,"Total de pallets a reservar",total_pallets,fl=XL_ALT_B); rn+=2

    _hdr(rn,"ALERTAS","5C0000"); rn+=1
    _row(rn,"Materiais CRÍTICOS",len(criticos),fl=XL_CRITICO); rn+=1
    _row(rn,"Materiais sem padrão de pallet",len(sem_pp_av),fl=XL_ALERTA); rn+=2

    # Seção PCP
    _hdr(rn,"PCP — Ação imediata","7B241C"); rn+=1
    if criticos:
        for ins,av in criticos:
            _row(rn, f"  {ins} — {av['descricao'][:40]}", av["tipo_aviso"],
                 fl=XL_CRITICO,fv=XL_CRITICO); rn+=1
        if sem_pp_av:
            for ins,_ in sem_pp_av[:5]:
                _row(rn, f"  {ins} — Sem padrão logístico", "Cadastrar na MP01",
                     fl=XL_ALERTA,fv=XL_ALERTA); rn+=1
    else:
        _row(rn,"✅ Nenhuma ação crítica","—",fl=XL_OK_F,fv=XL_OK_F); rn+=1
    rn+=1

    # Seção Suprimentos
    _hdr(rn,"SUPRIMENTOS — Pallets a reservar","1E3A5F"); rn+=1
    dep_totais = defaultdict(int)
    for r in resultados:
        if r["status"]==STATUS_RESERVAR and isinstance(r.get("qtd_pallet"), (int, float)):
            dep_totais[r["dep_recep"]] += r["qtd_pallet"]
    for dep, n_pal in dep_totais.items():
        _row(rn, f"  {dep} ← {DEP_MP01}", f"{n_pal} pallets",fl=XL_ALT_B); rn+=1
    rn+=1

    # Seção Produção
    _hdr(rn,"PRODUÇÃO — Pallets por linha","1A3C2B"); rn+=1
    por_linha = defaultdict(int)
    for r in resultados:
        if r["status"]==STATUS_RESERVAR and isinstance(r.get("qtd_pallet"), (int, float)):
            por_linha[r["linha"]] += r["qtd_pallet"]
    for linha in sorted(por_linha, key=line_sort_key):
        _row(rn, f"  {linha}", f"{por_linha[linha]} pallets"); rn+=1

    for c in range(1,7):
        ws_s.column_dimensions[get_column_letter(c)].width=[28,12,12,14,12,10][c-1]

    # ── ABA RESERVA ───────────────────────────────────────────────────────────
    ws_r = wb.create_sheet("RESERVA"); ws_r.sheet_view.showGridLines=False
    hdr_r=["Data","Depósito Recep.","","Material","Qtd","Depósito","Recebedor"]
    for ci,h in enumerate(hdr_r,1):
        c=ws_r.cell(1,ci,h); c.fill=_fill(XL_HDR_RES)
        c.font=_font(bold=True,color="FFFFFF"); c.border=_thin(); c.alignment=_center()
    ws_r.row_dimensions[1].height=20

    # Gera uma entrada para CADA resultado RESERVAR com pallets.
    # Espelha exatamente o que está na planilha do VISCONF (todos os dias, todas as linhas).
    rr = 2
    for r in resultados:
        if r["status"] != STATUS_RESERVAR: continue
        n_pal = r["qtd_pallet"] if isinstance(r.get("qtd_pallet"), (int, float)) else 0
        if n_pal <= 0: continue
        fill  = _fill(XL_ALT_B if rr % 2 == 0 else XL_WHITE)
        data_val = r["data"]
        for _ in range(n_pal):
            vals = [data_val, r["dep_recep"], "", r["insumo"], 1, DEP_MP01, r["linha"]]
            for ci, v in enumerate(vals, 1):
                c = ws_r.cell(rr, ci, v)
                c.border = _thin(); c.fill = fill; c.font = _font()
                if ci == 1:
                    c.number_format = "DD/MM/AAAA"
                    c.alignment = _center()
                if ci in (2, 4, 6, 7): c.number_format = "@"
                if ci == 5: c.alignment = _center()
            rr += 1

    for i,w in enumerate([12,14,4,14,5,10,16],1):
        ws_r.column_dimensions[get_column_letter(i)].width=w
    ws_r.freeze_panes="A2"
    _log(f"  ✅ RESERVA: {rr-2} linha(s) de pallet")

    # ── ABA AVISOS ────────────────────────────────────────────────────────────
    ws_a = wb.create_sheet("AVISOS"); ws_a.sheet_view.showGridLines=False
    hdr_a=["Insumo","Descrição","Tipo","UM","Depósito",
            "Est. Inicial","NB Total","Consumo","NL Total","Est. Final",
            "Cobertura %","Qtd Padrão Pallet","N° Pallets MP01","Status",
            "Linhas OK","Linhas RESERVAR","Linhas SEM EST","Linhas NÃO ENC","Ação PCP"]
    for ci,h in enumerate(hdr_a,1):
        c=ws_a.cell(1,ci,h); c.fill=_fill(XL_HDR_AV)
        c.font=_font(bold=True,color="FFFFFF"); c.border=_thin()
        c.alignment=Alignment(horizontal="center",wrap_text=True)
    ws_a.row_dimensions[1].height=32

    ORDEM_AV={"CRÍTICO - NÃO ENCONTRADO NA MP01":0,"CRÍTICO - SEM ESTOQUE DISPONÍVEL":1,
              "CRÍTICO - ESTOQUE INSUFICIENTE":2,
              "ALERTA - SEM PADRÃO DE PALLET":3,"RESERVAR":4,"OK":5}
    ACOES={
        "CRÍTICO - NÃO ENCONTRADO NA MP01":  "⛔ COMUNICAR PCP - material não encontrado. Verificar estoque físico vs sistêmico",
        "CRÍTICO - SEM ESTOQUE DISPONÍVEL":  "⛔ COMUNICAR PCP IMEDIATAMENTE - material com qty=0 no depósito",
        "CRÍTICO - ESTOQUE INSUFICIENTE":    "⚠️ Estoque na MP01 insuficiente para cobrir a necessidade total",
        "ALERTA - SEM PADRÃO DE PALLET":     "📋 Cadastrar lotes tipo-50 na MP01 para habilitar cálculo de pallets",
        "RESERVAR":                           "✅ Reserva necessária",
        "OK":                                 "✅ Estoque suficiente",
    }
    FILLS_A={
        "CRÍTICO - NÃO ENCONTRADO NA MP01": XL_CRITICO,
        "CRÍTICO - SEM ESTOQUE DISPONÍVEL": XL_CRITICO,
        "CRÍTICO - ESTOQUE INSUFICIENTE":   XL_ALERTA,
        "ALERTA - SEM PADRÃO DE PALLET":    XL_ALERTA,
        "RESERVAR":                          XL_ALT_P,
        "OK":                                XL_OK_F,
    }

    avisos_ord = sorted(avisos.items(), key=lambda x:(ORDEM_AV.get(x[1]["tipo_aviso"],9),x[0]))
    ra=2
    for ins, av in avisos_ord:
        ta   = av["tipo_aviso"]
        fill = _fill(FILLS_A.get(ta,XL_WHITE))
        bold = "CRÍTICO" in ta or "ALERTA" in ta
        vals=[
            ins, av["descricao"], av["tipo"], av["um"], av["dep"],
            round(av["est_inicial"],2), round(av["nb_total"],2),
            round(av["consumo_total"],2), round(av["nl_total"],2), av["est_final"],
            av["cobertura_pct"],
            av.get("qtd_padrao") or "—",
            av.get("n_pallets_mp01") or 0,
            ta,
            ", ".join(av.get("linhas_ok",[])),
            ", ".join(av.get("linhas_reservar",[])),
            ", ".join(av.get("linhas_sem_est",[])),
            ", ".join(av.get("linhas_nao_enc",[])),
            ACOES.get(ta,""),
        ]
        for ci,v in enumerate(vals,1):
            c=ws_a.cell(ra,ci,v); c.border=_thin(); c.fill=fill; c.font=_font(bold=bold)
            if ci in (6,7,8,9,10): c.number_format="0.00"; c.alignment=_right()
            if ci==11: c.number_format="0.0"; c.alignment=_center()
            if ci in (1,3,4,5,14): c.alignment=_center()
            if ci in (1,2): c.number_format="@"
        ra+=1

    ws_a.freeze_panes="A2"
    for i,w in enumerate([14,32,12,6,8,12,12,12,12,12,10,14,12,30,20,20,20,20,44],1):
        ws_a.column_dimensions[get_column_letter(i)].width=w

    _log(f"  ✅ AVISOS: {ra-2} insumos analisados")

    wb.save(output_path)
    _log(f"  💾 Relatório: {Path(output_path).name}")
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE COMPLETO
# ══════════════════════════════════════════════════════════════════════════════

# ==============================================================
# FECHAMENTO CIRÚRGICO — ARQUIVOS SAP
# ==============================================================

def _aguardar_workbook_aberto(
    nome_arquivo: str,
    pasta_destino: str,
    log_fn=None,
    timeout: int = 20
):
    """
    Polling ativo até o Workbook estar acessível via COM.

    Corrige o problema de timing onde a exportação do SAP termina,
    mas o Excel ainda está carregando quando o código tenta acessar
    ou fechar o arquivo.
    """
    import time
    import win32com.client
    from pathlib import Path

    def _log(m):
        if log_fn: log_fn(m)

    nome_base = Path(nome_arquivo).name.lower()
    caminho_full = Path(pasta_destino) / nome_arquivo

    # Primeiro aguarda o arquivo aparecer fisicamente no disco
    for _ in range(timeout * 2):
        if caminho_full.exists():
            break
        time.sleep(0.5)
    else:
        _log(f"  [!] {nome_arquivo} nunca apareceu no disco.")
        return None

    # Depois aguarda o Workbook ficar realmente acessível via COM
    for _ in range(timeout * 2):
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")

            for i in range(1, xl.Workbooks.Count + 1):
                try:
                    wb = xl.Workbooks(i)

                    if wb.Name.lower() == nome_base:
                        # Confirma que o COM está totalmente inicializado
                        _ = wb.Sheets.Count
                        return wb

                except Exception:
                    continue

        except Exception:
            pass

        time.sleep(0.5)

    _log(f"  [!] {nome_arquivo} não ficou acessível via COM em {timeout}s.")
    return None


def fechar_arquivo_sap(
    pasta_destino: str,
    nome_arquivo: str,
    log_fn=None
):
    wb = _aguardar_workbook_aberto(
        nome_arquivo,
        pasta_destino,
        log_fn
    )

    if wb is None:
        if log_fn: log_fn(f"  [!] {nome_arquivo} não pôde ser fechado.")
        return

    try:
        app = wb.Application
        restantes = app.Workbooks.Count

        wb.Close(SaveChanges=False)

        if log_fn: log_fn(f"  🔒 {nome_arquivo} fechado.")

        if restantes <= 1:
            try:
                app.Quit()
                if log_fn: log_fn("  [OK] Instância Excel exclusiva finalizada.")
            except Exception:
                pass

    except Exception as e:
        if log_fn: log_fn(f"  [!] Erro ao fechar {nome_arquivo}: {type(e).__name__}")


def _get_sap_session(log_fn=None):
    import win32com.client
    import subprocess
    import time
    import os
    def _log(m):
        if log_fn: log_fn(m)
        
    try:
        sap = win32com.client.GetObject("SAPGUI")
        _log("  [SAP] SAP GUI detectado como aberto. Utilizando primeira sessão disponível...")
    except Exception:
        _log("  [SAP] SAP GUI não encontrado. Iniciando SAP Logon...")
        subprocess.Popen([r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"])
        time.sleep(5)  # espera inicial para o SAP Logon carregar
        _log("  [SAP] Aguardando SAP Logon inicializar (até 45s)...")
        for attempt in range(45):
            try:
                sap = win32com.client.GetObject("SAPGUI")
                _log(f"  [SAP] SAP GUI conectado após {attempt + 6}s")
                break
            except:
                time.sleep(1)
        else:
            raise Exception("Não foi possível conectar ao SAP GUI após 50 segundos.")

    app = sap.GetScriptingEngine
    if app.Children.Count == 0:
        _log("  [SAP] Conectando ao ambiente SAP_CONNECTION_NAME...")
        conn = app.OpenConnection("SAP_CONNECTION_NAME", True)
    else:
        conn = app.Children(0)
        
    time.sleep(2)
    session = conn.Children(0)
    return session


def _fechar_sap(session, log_fn=None):
    def _log(m):
        if log_fn: log_fn(m)
    try:
        if session:
            _log("  🔒 Fechando SAP...")
            session.findById("wnd[0]").Close()
            try:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").Press()
            except Exception:
                pass
            _log("  [OK] SAP fechado com sucesso.")
    except Exception as e:
        _log(f"  [!] Não foi possível fechar o SAP: {e}")


def _get_pasta_sap(log_fn=None):
    cfg = load_config()
    pasta = cfg.get("pasta_sap", "")
    import os
    if not pasta or not os.path.exists(pasta):
        if log_fn: log_fn("  ⚠️ Pasta SAP não configurada. Selecione a pasta base...")
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta = filedialog.askdirectory(title="Selecione a pasta raiz do SAP")
        root.destroy()
        if not pasta:
            raise Exception("Pasta SAP não selecionada. Operação cancelada.")
        cfg["pasta_sap"] = pasta
        save_config(cfg)
    return pasta


def exportar_producao_sap(session, pasta_base, log_fn=None):
    import os, time
    from datetime import datetime
    
    def _log(m):
        if log_fn: log_fn(m)
    _log("  🔄 Push Produção (SAP MB51)...")
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB51"
    session.findById("wnd[0]").sendVKey(0)
    
    _log("  ⏳ Preenchendo parâmetros MB51...")
    
    for field in ["MATNR-LOW", "WERKS-LOW", "LGORT-LOW", "CHARG-LOW", "LIFNR-LOW", 
                  "KUNNR-LOW", "BWART-LOW", "SOBKZ-LOW", "AUFNR-LOW", "BUDAT-LOW", 
                  "VGART-LOW", "ALV_DEF", "MATNR-HIGH", "WERKS-HIGH", "LGORT-HIGH", 
                  "BWART-HIGH", "BUDAT-HIGH"]:
        try: session.findById(f"wnd[0]/usr/ctxt{field}").Text = ""
        except: pass
        
    for field in ["MAT_KDAU-LOW", "MAT_KDPO-LOW", "USNAM-LOW", "XBLNR-LOW"]:
        try: session.findById(f"wnd[0]/usr/txt{field}").Text = ""
        except: pass

    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1000"
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "TS00"
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = datetime.now().strftime("%d.%m.%Y")
    session.findById("wnd[0]/usr/ctxtALV_DEF").Text = "/PRODMDC"
    
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append("101\n102")
        root.update()
        
        session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").Press()
        session.findById("wnd[1]/tbar[0]/btn[16]").Press()
        session.findById("wnd[1]/tbar[0]/btn[24]").Press()
        session.findById("wnd[1]/tbar[0]/btn[8]").Press()
        root.destroy()
    except Exception as e:
        _log(f"  ⚠️ Aviso clipboard: {e}")
    
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]").sendVKey(16)
    
    caminho_prod = os.path.join(pasta_base, "PROD.xlsx")
    if os.path.exists(caminho_prod):
        try: os.remove(caminho_prod)
        except: pass
        
    session.findById("wnd[1]/tbar[0]/btn[20]").Press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta_base
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PROD.xlsx"
    session.findById("wnd[1]/tbar[0]/btn[0]").Press()
    
    _log(f"  📤 Exportando MB51 para: {caminho_prod}")
    _log("  ⏳ Aguardando exportação da Produção no disco e acesso COM...")
    fechar_arquivo_sap(pasta_base, "PROD.xlsx", log_fn)
    _log("  ✅ Exportação da Produção concluída e arquivo fechado!")


def exportar_estoque_sap(session, pasta_base, log_fn=None):
    import os, time
    
    def _log(m):
        if log_fn: log_fn(m)
    _log("  🔄 Push Estoque (SAP LX02)...")
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nLX02"
    session.findById("wnd[0]").sendVKey(0)
    
    _log("  ⏳ Preenchendo parâmetros LX02...")
    
    for field in ["S1_LGNUM", "WERKS-LOW", "P_VARI"]:
        try: session.findById(f"wnd[0]/usr/ctxt{field}").Text = ""
        except: pass
        
    session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = "WH1"
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1000"
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/USER_LAYOUT"
    
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]").sendVKey(16)
    
    caminho_lx02 = os.path.join(pasta_base, "LX02.xlsx")
    if os.path.exists(caminho_lx02):
        try: os.remove(caminho_lx02)
        except: pass
        
    session.findById("wnd[1]/tbar[0]/btn[20]").Press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta_base
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "LX02.xlsx"
    session.findById("wnd[1]/tbar[0]/btn[0]").Press()
    
    try: session.findById("wnd[1]/tbar[0]/btn[11]").Press()
    except: pass
    
    _log(f"  📤 Exportando LX02 para: {caminho_lx02}")
    _log("  ⏳ Aguardando exportação do Estoque no disco e acesso COM...")
    fechar_arquivo_sap(pasta_base, "LX02.xlsx", log_fn)
    _log("  ✅ Exportação do Estoque (LX02) concluída e arquivo fechado!")


def injetar_dados_producao(caminho_prod, wb_vis, log_fn=None):
    import os, datetime as dt
    def _log(m):
        if log_fn: log_fn(m)

    wb_origem, app_origem, own_origem = _xw_open(caminho_prod, log_fn)
    _log("  ⏳ Injetando dados da Produção...")
    try:
        ws_origem = wb_origem.sheets[0]
        lr = ws_origem.range('A' + str(ws_origem.cells.last_cell.row)).end('up').row
        if lr > 1:
            dados = ws_origem.range(f'A2:J{lr}').value
            if not isinstance(dados[0], list): dados = [dados]
        else:
            dados = []
            
        dados_tratados = []
        for row in dados:
            mantem = True
            hora = row[7] # H = index 7
            if hora is not None:
                try:
                    if isinstance(hora, dt.datetime):
                        h_val = hora.time()
                    elif isinstance(hora, dt.time):
                        h_val = hora
                    else:
                        h_val = dt.datetime.strptime(str(hora).strip(), "%H:%M:%S").time()
                        
                    if h_val > dt.time(7, 0, 59):
                        mantem = False
                except:
                    pass
            if mantem:
                dados_tratados.append(row)
                
        ws_dest = wb_vis.sheets["Produção"]
        ws_dest.range("A2:J9999").clear_contents()
        if dados_tratados:
            ws_dest.range("A2").value = dados_tratados
            
    finally:
        try: 
            wb_origem.close()
            _log("  🔒 PROD.xlsx fechado.")
        except: 
            pass
        if own_origem and app_origem:
            try: 
                app_origem.quit()
                _log("  ✅ COM liberado.")
            except: 
                pass
        try: 
            os.remove(caminho_prod)
        except: 
            pass

    _log("  ✅ Dados da Produção injetados com sucesso!")


def injetar_dados_estoque(caminho_lx02, wb_vis, log_fn=None):
    import os
    def _log(m):
        if log_fn: log_fn(m)

    wb_origem, app_origem, own_origem = _xw_open(caminho_lx02, log_fn)
    _log("  ⏳ Injetando dados do Estoque...")
    try:
        ws_origem = wb_origem.sheets[0]
        lr = ws_origem.range('A' + str(ws_origem.cells.last_cell.row)).end('up').row
        lc = ws_origem.range('XFD4').end('left').column
        
        if lr >= 4:
            # Puxa a partir da linha 4 da exportação SAP (onde ficam os cabeçalhos)
            dados = ws_origem.range((4, 1), (lr, lc)).value
            if not isinstance(dados[0], list): dados = [dados]
        else:
            dados = []
            
        ws_dest = wb_vis.sheets["MP01"]
        ws_dest.range("A1:Z9999").clear_contents()
        
        if dados:
            # Cola a partir da célula A1 (sobrescrevendo o cabeçalho antigo do VISCONF com o do SAP)
            ws_dest.range("A1").value = dados
            
    finally:
        try: 
            wb_origem.close()
            _log("  🔒 LX02.xlsx fechado.")
        except: 
            pass
        if own_origem and app_origem:
            try: 
                app_origem.quit()
                _log("  ✅ COM liberado.")
            except: 
                pass
        try: 
            os.remove(caminho_lx02)
        except: 
            pass

    _log("  ✅ Dados do Estoque (MP01) injetados com sucesso!")


def push_producao_sap(visconf_path, log_fn=None):
    session = _get_sap_session(log_fn)
    pasta_sap = _get_pasta_sap(log_fn)
    pasta_base = os.path.join(pasta_sap, "Temp VSCF", "Temp PROD")
    os.makedirs(pasta_base, exist_ok=True)
    caminho_prod = os.path.join(pasta_base, "PROD.xlsx")
    exportar_producao_sap(session, pasta_base, log_fn)
    _fechar_sap(session, log_fn)
    wb_vis, _app, _own = _xw_open(visconf_path, log_fn)
    injetar_dados_producao(caminho_prod, wb_vis, log_fn)
    _xw_save(wb_vis, _own, log_fn)


def push_estoque_sap(visconf_path, log_fn=None):
    session = _get_sap_session(log_fn)
    pasta_sap = _get_pasta_sap(log_fn)
    pasta_base = os.path.join(pasta_sap, "Temp VSCF", "Temp LX02")
    os.makedirs(pasta_base, exist_ok=True)
    caminho_lx02 = os.path.join(pasta_base, "LX02.xlsx")
    exportar_estoque_sap(session, pasta_base, log_fn)
    _fechar_sap(session, log_fn)
    wb_vis, _app, _own = _xw_open(visconf_path, log_fn)
    injetar_dados_estoque(caminho_lx02, wb_vis, log_fn)
    _xw_save(wb_vis, _own, log_fn)


def executar_pipeline(visconf_path, plano_path, dates_filter, output_dir,
                      categorias_ativas, linhas_ativas, etapas_ativas,
                      log_fn=None) -> dict:
    """
    Pipeline completo VISCONF v4.2.0.
    """
    def _log(m):
        if log_fn: log_fn(m)

    stats      = {}
    registros  = []
    resultados = []
    avisos     = {}
    pallets_g  = {}

    # ══════════════════════════════════════════════════════════════════════════
    # CONECTAR AO VISCONF (xlwings) — UMA ÚNICA VEZ
    # ══════════════════════════════════════════════════════════════════════════
    _log("\n─── CONECTANDO AO VISCONF ───────────────────────────")
    wb_vis, _app_vis, _own_vis = _xw_open(visconf_path, log_fn)
    _log(f"  📎 Conectado: {wb_vis.name}")

    try:
        # ── 1. ETAPA DE EXTRAÇÃO SAP (PRODUÇÃO E/OU ESTOQUE) ─────────────────
        tem_producao = "push_producao" in etapas_ativas
        tem_estoque  = "push_estoque" in etapas_ativas

        caminho_prod = None
        caminho_lx02 = None

        if tem_producao or tem_estoque:
            _log("\n─── INICIANDO EXTRAÇÕES SAP (SESSÃO ÚNICA SEQUENCIAL) ─")
            session = _get_sap_session(log_fn)
            pasta_sap = _get_pasta_sap(log_fn)
            _log(f"📁 Pasta de exportação SAP: {pasta_sap}")

            # 1.1 PRODUÇÃO (SEMPRE VEM PRIMEIRO)
            if tem_producao:
                _log("\n─── PUXAR PRODUÇÃO (SAP MB51) ───────────────────────")
                caminho_prod = os.path.join(pasta_sap, "PROD.xlsx")
                try:
                    exportar_producao_sap(session, pasta_sap, log_fn)
                except Exception as e:
                    _log(f"  ❌ Erro ao exportar Produção SAP: {e}")

            # 1.2 ESTOQUE (SEGUNDO LUGAR)
            if tem_estoque:
                _log("\n─── PUXAR ESTOQUE (SAP LX02) ────────────────────────")
                caminho_lx02 = os.path.join(pasta_sap, "LX02.xlsx")
                try:
                    exportar_estoque_sap(session, pasta_sap, log_fn)
                except Exception as e:
                    _log(f"  ❌ Erro ao exportar Estoque SAP: {e}")

            # 1.3 FECHAR SAP APÓS AS EXTRAÇÕES HABILITADAS FINALIZAREM
            _fechar_sap(session, log_fn)

            # ── 2. INJETAR / TRATAR OS DADOS NO ARQUIVO PRINCIPAL ────────────
            _log("\n─── INJETANDO DADOS NO ARQUIVO PRINCIPAL (VISCONF) ──")
            if tem_producao and caminho_prod:
                _log(f"🔎 Procurando arquivo exportado em: {caminho_prod}")
                if os.path.exists(caminho_prod):
                    _log(f"✅ Arquivo encontrado: {caminho_prod}")
                    try:
                        injetar_dados_producao(caminho_prod, wb_vis, log_fn)
                    except Exception as e:
                        _log(f"  ❌ Erro ao injetar dados da Produção: {e}")
                else:
                    _log(f"❌ Arquivo não encontrado: {caminho_prod}")

            if tem_estoque and caminho_lx02:
                _log(f"🔎 Procurando arquivo exportado em: {caminho_lx02}")
                if os.path.exists(caminho_lx02):
                    _log(f"✅ Arquivo encontrado: {caminho_lx02}")
                    try:
                        injetar_dados_estoque(caminho_lx02, wb_vis, log_fn)
                    except Exception as e:
                        _log(f"  ❌ Erro ao injetar dados do Estoque: {e}")
                else:
                    _log(f"❌ Arquivo não encontrado: {caminho_lx02}")

        # ── 3. ATUALIZAR PLANEJAMENTO MFR (Plan Reader) ──────────────────────
        if "gravar_mfr" in etapas_ativas:
            _log("\n─── ATUALIZAR PLANEJAMENTO MFR ──────────────────────")
            registros = fase1_plan_reader(plano_path, dates_filter,
                                           categorias_ativas, linhas_ativas, log_fn)
            _log(f"  {len(registros)} registro(s) extraídos")
            stats["plan_registros"] = len(registros)
            gravar_planejamento_mfr(wb_vis, registros, log_fn)

        # ── 4. DESCONTAR PRODUÇÃO ────────────────────────────────────────────
        if "desconto" in etapas_ativas:
            _log("\n─── DESCONTAR PRODUÇÃO ──────────────────────────────")
            if not registros:
                for row in _xw_iter_rows(wb_vis.sheets["PLANEJAMENTO MFR"], min_row=2):
                    if not row[0]: continue
                    cat = linha_para_categoria(str(row[1]))
                    registros.append({
                        "data": row[0], "linha": str(row[1]),
                        "sku": str(row[2]).split('.')[0],
                        "ordem": str(row[3]).split('.')[0],
                        "qt": safe_float(row[4]),
                        "categoria": cat,
                        "dep_recep": CATEGORIA_DEP.get(cat, DEP_MP06),
                    })
            registros = fase2_desconto(registros, wb_vis.sheets["Produção"], log_fn)
            stats["efetivos"] = sum(1 for r in registros if r.get("qt_efetiva", r["qt"]) > 0)

        # ── 5. GERAR NECESSIDADE AUTOMÁTICA ──────────────────────────────────
        if "gravar_na" in etapas_ativas:
            _log("\n─── GERAR NECESSIDADE AUTOMÁTICA ────────────────────")
            estoques = construir_estoques(wb_vis, categorias_ativas, log_fn)

            # Se não há registros desta sessão, ler do VISCONF
            if not registros:
                for row in _xw_iter_rows(wb_vis.sheets["PLANEJAMENTO MFR"], min_row=2):
                    if not row[0]: continue
                    cat = linha_para_categoria(str(row[1]))
                    registros.append({
                        "data": row[0], "linha": str(row[1]),
                        "sku": str(row[2]).split('.')[0],
                        "ordem": str(row[3]).split('.')[0],
                        "qt": safe_float(row[4]),
                        "qt_efetiva": safe_float(row[4]),  # já descontado
                        "qt_produzida": 0.0,
                        "categoria": cat,
                        "dep_recep": CATEGORIA_DEP.get(cat, DEP_MP06),
                    })

            # Fase 3: Explosão BOM (calcula NL por linha, sem pallets)
            resultados, avisos = fase3_4_explosao(registros, wb_vis.sheets["BD"], estoques, log_fn)

            # Fase 4: Agregar pallets por (insumo, linha)
            _log("\n  Agregando pallets por (insumo × linha)...")
            resultados, pallets_g = agregar_pallets(
                resultados, estoques.get("MP01_pallet", {}), log_fn
            )

            sc = {}
            for r in resultados: sc[r["status"]] = sc.get(r["status"], 0) + 1
            n_prev = sum(1 for r in resultados if r.get("preventiva"))
            _log(f"  {len(resultados)} linhas | {sc} | {n_prev} preventivas")
            stats.update(sc)

            gravar_necessidade_automatica(wb_vis, resultados, log_fn)

        # ── 6. GERAR RELATÓRIO FINAL ─────────────────────────────────────────
        if "relatorio" in etapas_ativas and resultados:
            _log("\n─── GERAR RELATÓRIO FINAL ───────────────────────────")
            hoje     = datetime.now().strftime("%d.%m")
            rel_name = f"Reserva de Insumos - {hoje}.xlsx"
            rel_path = str(Path(output_dir) / rel_name)
            gerar_relatorio(resultados, avisos, registros, rel_path, log_fn)
            stats["relatorio"] = rel_path

        # ══════════════════════════════════════════════════════════════════════
        # SALVAR UMA ÚNICA VEZ NO FINAL (equivalente ao Ctrl+B do usuário)
        # ══════════════════════════════════════════════════════════════════════
        _log("\n─── SALVANDO VISCONF ────────────────────────────────")
        _xw_save(wb_vis, _own_vis, log_fn)

    except Exception:
        # Em caso de erro, tenta salvar o que já foi gravado
        try:
            _xw_save(wb_vis, _own_vis, log_fn)
        except Exception:
            pass
        raise

    return stats


# ══════════════════════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════════════════════

class VISCONFApp:
    PAD=10; PAD2=5

    def __init__(self, root):
        self.root    = root
        self.cfg     = load_config()
        self._mode   = tk.StringVar(value="local") # Mantido apenas para compatibilidade legada no dict, não muda mais
        self._plano  = tk.StringVar(value=self.cfg.get("plano",""))
        self._vis    = tk.StringVar(value=self.cfg.get("visconf",""))
        self._out    = tk.StringVar(value=self.cfg.get("output_dir",str(Path.home()/"Desktop")))
        self._pasta_sap = tk.StringVar(value=self.cfg.get("pasta_sap",""))

        self._dates: list[tuple[date,tk.BooleanVar]] = []
        self._cat_vars:   dict[str,tk.BooleanVar] = {}
        self._linha_vars: dict[str,tk.BooleanVar] = {}
        self._etapa_vars: dict[str,tk.BooleanVar] = {}
        self._running = False
        self._setup_window()
        self._build_ui()
        self._populate_dates()

    def _setup_window(self):
        self.root.title("Visconf Engine v4.2.0  ·  ACME Corp")
        self.root.geometry("860x1000")
        self.root.configure(bg=hx(C_BG))
        self.root.minsize(740,880)

    def _build_ui(self):
        r = self.root
        h = tk.Frame(r, bg=hx(C_BG)); h.pack(fill="x", padx=20, pady=(14,0))
        tk.Label(h, text="VISCONF ENGINE (v4.2.0)", font=("Helvetica Neue",12,"bold"),
                 bg=hx(C_BG), fg=hx(C_ACCENT)).pack(side="left")
        tk.Label(h, text="  Plan Reader · Desconto · Explosão · Pallets · Relatório",
                 font=("Helvetica Neue",9), bg=hx(C_BG), fg=hx(C_MUTED)).pack(side="left")
        ttk.Separator(r).pack(fill="x", padx=20, pady=(self.PAD,0))

        canvas = tk.Canvas(r, bg=hx(C_BG), highlightthickness=0)
        sb     = ttk.Scrollbar(r, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas, bg=hx(C_BG))
        cw   = canvas.create_window((0,0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))
        body.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        px=20

        # ── 1. FONTE DO PLANO ─────────────────────────────────────────────────
        self._sec(body,"1  FONTE DO PLANO"); sp=self._panel(body)
        self._f_local=self._subfr(sp)
        self._frow(self._f_local,"Arquivo do Plano:",self._plano,self._br_plano)
        self._f_local.pack(fill="x",padx=self.PAD,pady=self.PAD)

        # ── 2. DATAS ──────────────────────────────────────────────────────────
        self._sec(body,"2  DATAS DE INTERESSE"); dp=self._panel(body)
        af=tk.Frame(dp,bg=hx(C_PANEL)); af.pack(fill="x",padx=px,pady=(self.PAD,self.PAD2))
        for lbl,fn in [("Hoje + amanhã",self._sel_2d),("Sex → Seg",self._sel_fds),("Limpar",self._sel_none)]:
            self._btn(af,lbl,fn).pack(side="left",padx=(0,6))
        self._chk=tk.Frame(dp,bg=hx(C_PANEL)); self._chk.pack(fill="x",padx=px,pady=(0,self.PAD))

        # ── 3. CATEGORIAS E LINHAS ────────────────────────────────────────────
        self._sec(body,"3  CATEGORIAS E LINHAS")
        cat_panel=self._panel(body)
        for cat,linhas in CATEGORIA_LINHAS.items():
            cf=tk.Frame(cat_panel,bg="#383838"); cf.pack(fill="x",padx=px,pady=(0,self.PAD2))
            var_cat=tk.BooleanVar(value=True); self._cat_vars[cat]=var_cat
            dep_lbl=CATEGORIA_DEP.get(cat,"")
            tk.Checkbutton(cf,text=f"  {cat}  [{dep_lbl}]",variable=var_cat,
                           command=lambda c=cat:self._toggle_cat(c),
                           bg="#383838",fg=hx(C_ACCENT),selectcolor="#383838",
                           activebackground="#383838",activeforeground=hx(C_ACCENT),
                           font=("Helvetica Neue",10,"bold")).pack(side="left",padx=(0,8),pady=self.PAD2)
            lf=tk.Frame(cf,bg="#383838"); lf.pack(side="left",fill="x",expand=True)
            for i,ln in enumerate(linhas):
                var_l=tk.BooleanVar(value=True); self._linha_vars[ln]=var_l
                tk.Checkbutton(lf,text=ln,variable=var_l,bg="#383838",fg=hx(C_TEXT),
                               selectcolor="#383838",activebackground="#383838",
                               activeforeground=hx(C_ACCENT),font=("Helvetica Neue",9)
                               ).grid(row=i//6,column=i%6,sticky="w",padx=(0,10),pady=1)

        # ── 4. PIPELINE ───────────────────────────────────────────────────────
        self._sec(body,"4  PIPELINE")
        ep=self._panel(body)
        etapas=[
            ("push_producao", "🏭 Puxar Produção (SAP MB51)"),
            ("desconto",      "➖ Descontar Produção"),
            ("push_estoque",  "📦 Puxar Estoque (SAP LX02)"),
            ("gravar_mfr",    "📋 Atualizar Planejamento MFR"),
            ("gravar_na",     "💥 Gerar Necessidade Automática"),
            ("relatorio",     "📊 Gerar Relatório Final"),
        ]
        ef=tk.Frame(ep,bg=hx(C_PANEL)); ef.pack(fill="x",padx=px,pady=self.PAD)
        for i,(key,lbl) in enumerate(etapas):
            default_val = key not in ("push_producao","push_estoque")
            var=tk.BooleanVar(value=default_val); self._etapa_vars[key]=var
            tk.Checkbutton(ef,text=lbl,variable=var,bg=hx(C_PANEL),fg=hx(C_TEXT),
                           selectcolor=hx(C_PANEL),activebackground=hx(C_PANEL),
                           activeforeground=hx(C_ACCENT),font=("Helvetica Neue",9)
                           ).grid(row=i//2,column=i%2,sticky="w",padx=(0,20),pady=3)

        # Opção de salvar no VISCONF


        # ── 5. ARQUIVOS ───────────────────────────────────────────────────────
        self._sec(body,"5  ARQUIVOS (VISCONF / SAP)"); fp=self._panel(body)
        f5=tk.Frame(fp,bg=hx(C_PANEL)); f5.pack(fill="x",padx=px,pady=self.PAD)
        self._frow(f5,"VISCONF:",self._vis,self._br_vis,row=0)
        self._frow(f5,"Pasta SAP:",self._pasta_sap,self._br_sap,row=1)
        self._frow(f5,"Pasta de saída:",self._out,self._br_out,row=2)

        self._btn_run=tk.Button(body,text="▶  EXECUTAR PIPELINE",
            bg=hx(C_ACCENT),fg="#000000",activebackground="#25A244",activeforeground="#000000",
            font=("Helvetica Neue",12,"bold"),relief="flat",cursor="hand2",
            padx=self.PAD,pady=12,command=self._run)
        self._btn_run.pack(fill="x",padx=20,pady=(self.PAD,self.PAD2))

        style=ttk.Style(); style.theme_use("default")
        style.configure("G.Horizontal.TProgressbar",troughcolor=hx(C_PANEL),background=hx(C_ACCENT))
        self._prog=ttk.Progressbar(body,style="G.Horizontal.TProgressbar",mode="indeterminate")
        self._prog.pack(fill="x",padx=20,pady=(0,self.PAD2))

        self._sec(body,"LOG DE EXECUÇÃO"); lp=self._panel(body)
        self._log_w=tk.Text(lp,height=12,bg=hx(C_BG),fg=hx(C_TEXT),
                             insertbackground=hx(C_TEXT),relief="flat",
                             font=("Menlo",9),wrap="word",state="disabled")
        self._log_w.pack(fill="both",expand=True,padx=px,pady=self.PAD)
        for tag,fg,bold in [("ok",C_ACCENT,False),("err",C_ERR,False),
                              ("warn",C_WARN,False),("sec",C_ACCENT2,True),("dim",C_MUTED,False)]:
            kw={"foreground":hx(fg)}
            if bold: kw["font"]=("Menlo",9,"bold")
            self._log_w.tag_configure(tag,**kw)

    def _sec(self,p,t):
        tk.Label(p,text=t,bg=hx(C_BG),fg=hx(C_MUTED),
                 font=("Helvetica Neue",8,"bold"),anchor="w"
                 ).pack(fill="x",pady=(self.PAD,self.PAD2),padx=20)
    def _panel(self,p):
        f=tk.Frame(p,bg=hx(C_PANEL)); f.pack(fill="x",pady=(0,self.PAD),padx=20); return f
    def _subfr(self,parent):
        f=tk.Frame(parent,bg=hx(C_PANEL)); f.pack(fill="x",padx=self.PAD,pady=(0,self.PAD)); return f
    def _frow(self,parent,label,var,cmd,row=0):
        tk.Label(parent,text=label,bg=hx(C_PANEL),fg=hx(C_MUTED),
                 font=("Helvetica Neue",9),width=14,anchor="w"
                 ).grid(row=row,column=0,sticky="w",pady=2)
        tk.Entry(parent,textvariable=var,width=52,bg=hx(C_BG),fg=hx(C_TEXT),
                 insertbackground=hx(C_TEXT),relief="flat",font=("Helvetica Neue",9)
                 ).grid(row=row,column=1,padx=(8,4),sticky="ew",pady=2)
        self._btn(parent,"···",cmd).grid(row=row,column=2,pady=2)
        parent.columnconfigure(1,weight=1)
    def _btn(self,p,t,cmd):
        return tk.Button(p,text=t,command=cmd,bg=hx(C_BTN),fg=hx(C_TEXT),
                         activebackground="#4A4A4E",relief="flat",
                         cursor="hand2",font=("Helvetica Neue",9),padx=8,pady=4)
    def _toggle_cat(self,cat):
        val=self._cat_vars[cat].get()
        for ln in CATEGORIA_LINHAS.get(cat,[]): 
            if ln in self._linha_vars: self._linha_vars[ln].set(val)
    def _populate_dates(self):
        for w in self._chk.winfo_children(): w.destroy()
        self._dates.clear()
        today=date.today(); days=["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"]
        for i in range(14):
            d=today+timedelta(days=i); var=tk.BooleanVar(value=(i<2))
            self._dates.append((d,var))
            fg=hx(C_WARN) if d.weekday()==4 else hx(C_TEXT)
            tk.Checkbutton(self._chk,text=f"{days[d.weekday()]}  {d.strftime('%d/%m')}",
                           variable=var,bg=hx(C_PANEL),fg=fg,selectcolor=hx(C_PANEL),
                           activebackground=hx(C_PANEL),activeforeground=hx(C_ACCENT),
                           font=("Helvetica Neue",9)
                           ).grid(row=i//7,column=i%7,sticky="w",padx=(0,12),pady=2)
    def _sel_2d(self):  [var.set(i<2) for i,(_,var) in enumerate(self._dates)]
    def _sel_fds(self):
        t=set()
        for i in range(7):
            d=date.today()+timedelta(days=i)
            if d.weekday() in {4,5,6,0}: t.add(d)
            if d.weekday()==0: break
        [var.set(d in t) for d,var in self._dates]
    def _sel_none(self): [var.set(False) for _,var in self._dates]
    def _br_plano(self):
        p=filedialog.askopenfilename(title="Plano",filetypes=[("Excel","*.xlsx *.xlsm"),("Todos","*.*")])
        if p: self._plano.set(p)
    def _br_vis(self):
        p=filedialog.askopenfilename(title="VISCONF",filetypes=[("Excel Macro","*.xlsm"),("Todos","*.*")])
        if p: self._vis.set(p)
    def _br_out(self):
        p=filedialog.askdirectory(title="Pasta de saída")
        if p: self._out.set(p)

    def _br_sap(self):
        p=filedialog.askdirectory(title="Pasta de exportação do SAP")
        if p: self._pasta_sap.set(p)

    def _log(self,msg,tag=""):
        def _do():
            self._log_w.configure(state="normal")
            ts=datetime.now().strftime("%H:%M:%S")
            self._log_w.insert("end",f"[{ts}] {msg}\n",tag)
            self._log_w.see("end"); self._log_w.configure(state="disabled")
        self.root.after(0,_do)

    def _run(self):
        if self._running: return
        dates_sel=[d for d,var in self._dates if var.get()]
        if not dates_sel: messagebox.showwarning("Atenção","Selecione ao menos uma data."); return
        vis=self._vis.get().strip(); out=self._out.get().strip()
        if not vis or not os.path.exists(vis): messagebox.showerror("Erro","VISCONF não encontrado."); return
        if not out or not os.path.isdir(out): messagebox.showerror("Erro","Pasta de saída inválida."); return
        precisa_plano = self._etapa_vars.get("gravar_mfr", tk.BooleanVar(value=False)).get()
        if precisa_plano:
            plano=self._plano.get().strip()
            if not plano or not os.path.exists(plano): messagebox.showerror("Erro","Plano não encontrado."); return
        else:
            # MFR desabilitado — plano não é necessário
            plano=self._plano.get().strip() or None
            
        cats={cat for cat,var in self._cat_vars.items() if var.get()}
        lins={ln for ln,var in self._linha_vars.items() if var.get()}
        etps={ep for ep,var in self._etapa_vars.items() if var.get()}
        if not cats: messagebox.showwarning("Atenção","Selecione ao menos uma categoria."); return
        if not etps: messagebox.showwarning("Atenção","Selecione ao menos uma etapa."); return
        
        if ("push_producao" in etps or "push_estoque" in etps):
            if not self._pasta_sap.get().strip() or not os.path.exists(self._pasta_sap.get().strip()):
                messagebox.showerror("Erro","Pasta SAP inválida ou não configurada. Configure na seção 5 - Arquivos.")
                return
        
        existing_cfg = load_config()
        existing_cfg.update({
            "plano":self._plano.get(), 
            "visconf":vis, 
            "output_dir":out,
            "pasta_sap":self._pasta_sap.get()
        })
        save_config(existing_cfg)
        
        self._running=True
        self._btn_run.configure(state="disabled",text="  Executando...",bg=hx(C_BTN))
        self._prog.start(10)
        src=plano
        threading.Thread(target=self._thread,
            args=(src,vis,set(dates_sel),out,cats,lins,etps),
            daemon=True).start()

    def _thread(self,src,vis,dates_set,out,cats,lins,etps):
        try:
            self._log("═"*42,"sec"); self._log("  VISCONF ENGINE (v4.2.0)","sec"); self._log("═"*42,"sec")
            self._log(f"Datas: {', '.join(d.strftime('%d/%m') for d in sorted(dates_set))}","dim")
            self._log(f"Categorias: {', '.join(sorted(cats))}","dim")

            if "gravar_mfr" in etps:
                plano=src
            else:
                plano = None

            stats=executar_pipeline(vis,plano,dates_set,out,cats,lins,etps,
                                    log_fn=self._log)

            self._log("═"*42,"sec"); self._log("  ✅ CONCLUÍDO","ok"); self._log("═"*42,"sec")
            sc_str="\n".join(f"{k}: {v}" for k,v in stats.items() if k!="relatorio")
            rel=stats.get("relatorio","")
            msg=f"Pipeline concluído!\n\n{sc_str}"
            if rel: msg+=f"\n\nRelatório: {Path(rel).name}"
            self.root.after(0,lambda:messagebox.showinfo("Concluído",msg))
        except Exception as e:
            self._log(f"  ERRO: {e}","err")
            err=str(e); self.root.after(0,lambda:messagebox.showerror("Erro",err))
        finally:
            self.root.after(0,self._reset)

    def _reset(self):
        self._running=False; self._prog.stop()
        self._btn_run.configure(state="normal",text="▶  EXECUTAR PIPELINE",bg=hx(C_ACCENT))


def main():
    root=tk.Tk(); ttk.Style(root).theme_use("default"); VISCONFApp(root); root.mainloop()

if __name__=="__main__":
    main()
