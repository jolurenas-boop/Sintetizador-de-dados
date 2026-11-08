"""
================================================================================
PORTAL DE BIOENGENHARIA: FUSÃO MULTI-RELATÓRIO (CIVM + FADIGA sEMG)
UPLOAD EM LOTE, EXTRAÇÃO AUTOMÁTICA POR NOME DE ARQUIVO,
TABELA MESTRA, COMPARAÇÃO CROSSOVER (CONTROLE×INTERVENÇÃO / PRÉ×PÓS),
ANOVA DE MEDIDAS REPETIDAS & LAUDO WEASYPRINT
================================================================================

Convenção de nome de arquivo para o upload em lote (recomendada):
    <ID>_<CIVM|F50|FAD><C|I><PRE|POS>.pdf
    Exemplos: 001_CIVMCPRE.pdf, 001_CIVMIPOS.pdf, 001_FADCPRE.pdf, 001_FADIGAIPOS.pdf
    C = Controle | I = Intervenção | PRE = Pré-sessão | POS = Pós-sessão

Arquivos cujo nome não siga o padrão caem automaticamente em uma fila de
"pendências" na própria tela de upload, onde o usuário completa os metadados
manualmente antes de confirmar o processamento.
"""

from __future__ import annotations   # anotações lazy: o código não exige Python >= 3.10 para carregar

import os
import io
import re
import sys
import base64
import hashlib
import json
import inspect
import unicodedata
import itertools
import functools
import importlib
import importlib.util
import importlib.metadata

import streamlit as st


def _has_module(name: str) -> bool:
    """find_spec que nunca levanta exceção (ex.: ValueError quando __spec__ é None)."""
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DE INTERFACE STREAMLIT (precisa ser o 1º comando Streamlit)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="BioEng: Fusão CIVM & Fadiga sEMG",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# VERIFICAÇÃO DE DEPENDÊNCIAS
# Em vez de derrubar o app com "ModuleNotFoundError" (mensagem censurada pelo Streamlit
# Cloud), mostramos exatamente o que falta e como corrigir.
# -----------------------------------------------------------------------------
REQUIREMENTS_TXT = """streamlit
pandas
numpy
scipy
matplotlib
pdfplumber
openpyxl
pingouin
weasyprint"""

_REQUIRED_MODULES = {'numpy': 'numpy', 'pandas': 'pandas', 'matplotlib': 'matplotlib', 'pdfplumber': 'pdfplumber'}
_missing = [pkg for mod, pkg in _REQUIRED_MODULES.items() if not _has_module(mod)]
if _missing:
    st.error(f"❌ Dependência(s) obrigatória(s) não instalada(s) neste servidor: **{', '.join(_missing)}**")
    st.markdown("""
**Como corrigir no Streamlit Community Cloud**

1. Garanta que existe um arquivo **`requirements.txt`** na **raiz do repositório** (ou na mesma pasta do arquivo principal) com o conteúdo abaixo.
2. Faça *commit* e *push*. O Streamlit Cloud reinstala as dependências a cada alteração desse arquivo.
3. Em **Manage app → ⋮ → Reboot app** (se necessário, **Clear cache** antes).
4. Evite espaços e parênteses no nome do arquivo principal (ex.: `app_streamlit_master (4).py`). Renomeie para `app_streamlit_master.py` e confira o *Main file path* nas configurações do app.
5. *(Opcional, só para gerar o laudo em PDF)* crie também um `packages.txt` com as linhas `libpango-1.0-0` e `libpangoft2-1.0-0`.
""")
    st.code(REQUIREMENTS_TXT, language='text')
    st.caption(f"Python {sys.version.split()[0]}")
    st.stop()

import numpy as np
import pandas as pd

# Disponibilidade sem importar (find_spec é barato); o import pesado só acontece quando o recurso é usado.
PINGOUIN_AVAILABLE = _has_module('pingouin')
SCIPY_AVAILABLE = _has_module('scipy')
WEASYPRINT_INSTALLED = _has_module('weasyprint')
OPENPYXL_AVAILABLE = _has_module('openpyxl')


@functools.lru_cache(maxsize=None)
def _lazy(module_name: str):
    """Importa sob demanda (e uma única vez). Retorna None se faltar ou falhar (ex.: libs nativas do WeasyPrint)."""
    try:
        return importlib.import_module(module_name)
    except Exception:
        return None


@functools.lru_cache(maxsize=None)
def _plt():
    """pyplot com backend não interativo, importado só quando um gráfico é realmente gerado."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    return plt


# st.fragment (Streamlit >= 1.37) reexecuta só um trecho da página; sem ele, vira identidade.
_fragment = getattr(st, 'fragment', None) or getattr(st, 'experimental_fragment', None) or (lambda f: f)


def _stretch_for(fn) -> dict:
    """Largura total compatível com versões novas (width='stretch') e antigas (use_container_width=True)."""
    try:
        params = inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return {}
    if 'width' in params and 'stretch' in str(params['width'].annotation):
        return {'width': 'stretch'}
    if 'use_container_width' in params:
        return {'use_container_width': True}
    return {}


DF_STRETCH = _stretch_for(st.dataframe)
EDITOR_STRETCH = _stretch_for(st.data_editor)
DL_STRETCH = _stretch_for(st.download_button)


def dependency_status() -> list:
    spec = [('streamlit', True), ('pandas', True), ('numpy', True), ('matplotlib', True), ('pdfplumber', True),
            ('scipy', False), ('openpyxl', False), ('pingouin', False), ('weasyprint', False)]
    out = []
    for name, required in spec:
        ok = _has_module(name)
        ver = None
        if ok:
            try:
                ver = importlib.metadata.version(name)
            except Exception:
                ver = None
        out.append({'pacote': name, 'ok': ok, 'versao': ver, 'required': required})
    return out

st.markdown("""
<style>
    .main-title { font-size: 2.1rem; font-weight: 800; color: #1E3A8A; margin-bottom: 0.1rem; }
    .sub-title { font-size: 1.0rem; color: #4B5563; margin-bottom: 1.2rem; }
    .kpi-box { background-color: #F8FAFC; border-radius: 8px; padding: 1rem;
               border-left: 4px solid #2563EB; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
    .flag-box { background-color: #FEF2F2; border-left: 4px solid #DC2626; border-radius: 6px;
                padding: 0.8rem 1rem; margin-bottom: 0.6rem; font-size: 0.9rem; }
    .ok-box { background-color: #F0FDF4; border-left: 4px solid #16A34A; border-radius: 6px;
              padding: 0.8rem 1rem; margin-bottom: 0.6rem; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

# Inicialização de Session State
for key, default in [
    ('df_civm_master', pd.DataFrame()),       # formato largo: 1 linha por teste (CIVM)
    ('df_fatigue_master', pd.DataFrame()),    # formato longo: 1 linha por músculo por teste (Fadiga)
    ('df_merged_master', pd.DataFrame()),     # tabela mestra consolidada (o "modelo")
    ('pending_raw_files', {}),                # bytes dos arquivos não reconhecidos pelo nome
    ('pending_meta_table', pd.DataFrame()),   # tabela editável de metadados pendentes
    ('audit_store', {}),                       # {sha256: registro de auditoria} — texto/tabelas de origem, hash, estado
    ('audit_version', 0),                      # contador p/ memoizar o JSON de auditoria
    ('audit_json_cache', None),
    ('data_version', 0),                       # sobe a cada mudança nos bancos CIVM/Fadiga
    ('master_version', -1),                    # versão com a qual a Tabela Mestra foi construída
    ('laudo_pdf_key', None),                   # hash do HTML cujo PDF o usuário pediu para compilar
    ('last_chart_b64', None),
    ('last_stats_res', None),
    ('last_civm_col', None),
    ('last_fadiga_col', None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

def bump_data_version() -> None:
    """Sinaliza que os bancos mudaram: a Tabela Mestra será refeita sob demanda (get_master)."""
    st.session_state['data_version'] += 1


def get_master(force: bool = False) -> pd.DataFrame:
    """Tabela Mestra sempre atual, reconstruída apenas quando os bancos mudaram (ou se force=True)."""
    if force or st.session_state['master_version'] != st.session_state['data_version']:
        st.session_state['df_merged_master'] = build_master_table(
            st.session_state['df_civm_master'], st.session_state['df_fatigue_master'])
        st.session_state['master_version'] = st.session_state['data_version']
    return st.session_state['df_merged_master']


def clear_all_data() -> None:
    st.session_state['df_civm_master'] = pd.DataFrame()
    st.session_state['df_fatigue_master'] = pd.DataFrame()
    st.session_state['df_merged_master'] = pd.DataFrame()
    st.session_state['pending_raw_files'] = {}
    st.session_state['pending_meta_table'] = pd.DataFrame()
    st.session_state['audit_store'] = {}
    st.session_state['audit_version'] += 1
    st.session_state['audit_json_cache'] = None
    for k in ('last_chart_b64', 'last_stats_res', 'last_civm_col', 'last_fadiga_col', 'laudo_pdf_key'):
        st.session_state[k] = None
    bump_data_version()


MARKERS = ['o', 's', '^', 'D', 'v', 'P', 'X', '*', '<', '>']
COND_COLORS = {"Controle": "#6B8CAE", "Intervenção": "#D98E4A"}

# ==============================================================================
# MÓDULO 1: HELPERS GERAIS (NORMALIZAÇÃO, NÚMEROS, NOMES)
# ==============================================================================

@functools.lru_cache(maxsize=8192)
def _norm(s) -> str:
    """Remove acentuação e normaliza caixa/espacos para comparação robusta de texto (memoizada)."""
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).lower().strip()

_UNIT_RE = re.compile(r'(kgf/mV|kgf/s|uV/s|Hz/s|kgf|mV|uV|Hz|ms\b|%|\+)', re.IGNORECASE)

def clean_float(val) -> float:
    """
    Converte números no formato brasileiro (milhar com ponto, decimal com vírgula)
    e remove unidades comuns (kgf, mV, Hz, %, s, ms etc.) para float.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return np.nan
    s = str(val).strip()
    if s in ('', '-', '—', '–', 'nan', 'None', 'N/A'):
        return np.nan
    s = _UNIT_RE.sub('', s).strip()
    s = s.replace('\xa0', ' ').replace(' ', '')
    if ',' in s and s.count('.') <= 1:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except (ValueError, TypeError):
        return np.nan

def standardize_muscle_name(name: str) -> str:
    """Padroniza os nomes dos músculos/canais segundo a nomenclatura oficial do laboratório."""
    clean = _norm(name)
    is_left = any(w in clean for w in ['esquerdo', 'esquerda', 'pme', 'tbe', 'dle', 'cce']) or clean.endswith(' e') or '(e)' in clean or '_e' in clean
    if 'peitoral' in clean or re.search(r'\bpm\b', clean):
        return 'Peitoral maior esquerdo (PME)' if is_left else 'Peitoral maior direito (PMD)'
    elif 'triceps' in clean or re.search(r'\btb\b', clean):
        return 'Tríceps braquial esquerdo (TBE)' if is_left else 'Tríceps braquial direito (TBD)'
    elif 'deltoide' in clean or re.search(r'\bdl\b', clean):
        return 'Deltóide esquerdo (DLE)' if is_left else 'Deltóide direito (DLD)'
    elif any(k in clean for k in ['celula', 'carga']) or re.search(r'\bcc\b', clean):
        return 'Célula de carga esquerda (CCE)' if is_left else 'Célula de carga direita (CCD)'
    return str(name).strip()

def parse_filename_metadata(filename: str) -> dict:
    """
    Extrai ID do voluntário, tipo de relatório (CIVM/FADIGA), condição e momento
    a partir do nome do arquivo, no padrão <ID>_<CIVM|F50|FAD><C|I><PRE|POS>.
    Retorna matched=False se o padrão não for reconhecido (fica pendente de revisão manual).
    """
    stem = os.path.splitext(os.path.basename(filename))[0]
    clean = re.sub(r'[^A-Za-z0-9]', '', stem).upper()
    m = re.match(r'^(\d+)(CIVM|F50|FAD(?:IGA)?)([CI])(PRE|POS)$', clean)
    if not m:
        return {'matched': False, 'id': None, 'tipo': None, 'condicao': None, 'momento': None}
    vol_id, tipo, cond, mom = m.groups()
    return {
        'matched': True,
        'id': vol_id.zfill(3),
        'tipo': 'CIVM' if tipo == 'CIVM' else 'FADIGA',
        'condicao': 'Controle' if cond == 'C' else 'Intervenção',
        'momento': 'Pré' if mom == 'PRE' else 'Pós',
    }

def find_table_with_headers(tables, must_have):
    """Encontra uma tabela sem remover células vazias do cabeçalho (índices preservados)."""
    for t in tables:
        if not t or len(t) < 2:
            continue
        header = [_norm(c or '') for c in t[0]]
        if all(any(_norm(k) in h for h in header) for k in must_have):
            return t, header
    return None, None


@st.cache_data(show_spinner=False, max_entries=256)
def extract_pdf_content(file_bytes: bytes) -> dict:
    """
    Lê o PDF UMA única vez (texto + tabelas por página) e devolve o conteúdo literal.

    Antes, cada PDF era aberto e processado duas vezes (parser + auditoria) — a etapa mais
    cara do app. Agora parsers e auditoria compartilham este resultado, cacheado pelo
    conteúdo do arquivo: reprocessar o mesmo PDF é instantâneo.
    """
    sha = hashlib.sha256(file_bytes).hexdigest()
    pdfplumber = _lazy('pdfplumber')
    if pdfplumber is None:
        return {'pages': [], 'error': 'pdfplumber não está instalado neste ambiente.', 'sha256': sha}
    pages, error = [], ''
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_no, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ''
                tables = page.extract_tables() or []
                pages.append({
                    'pagina': page_no,
                    'texto_extraido': text,
                    'tabelas_extraidas': tables,
                    'texto_encontrado': bool(text.strip()),
                    'tabelas_encontradas': len(tables),
                })
                try:
                    page.flush_cache()          # libera objetos internos da página já processada
                except Exception:
                    pass
    except Exception as exc:
        error = str(exc)
    return {'pages': pages, 'error': error, 'sha256': sha}


def collect_tables(content: dict):
    """Achata o conteúdo extraído: (texto completo, lista de tabelas, {id(tabela): página})."""
    texts, all_tables, table_pages = [], [], {}
    for p in content['pages']:
        texts.append(p['texto_extraido'])
        for table in p['tabelas_extraidas']:
            all_tables.append(table)
            table_pages[id(table)] = p['pagina']
    return "\n".join(texts), all_tables, table_pages


def pdf_audit_record(content: dict, filename: str, meta: dict | None,
                     tipo: str, status: str, rows_extracted: int = 0,
                     error: str = '') -> dict:
    """Gera trilha de auditoria literal (sem reabrir o PDF), sem inferir conteúdo ausente."""
    pages = content['pages']
    text_pages = sum(bool(p['texto_encontrado']) for p in pages)
    table_count = sum(p['tabelas_encontradas'] for p in pages)
    return {
        'Arquivo_PDF': os.path.basename(filename),
        'SHA256_Arquivo': content['sha256'],
        'Tipo_Relatorio': tipo,
        'ID_Voluntario': (meta or {}).get('id'),
        'Condicao': (meta or {}).get('condicao'),
        'Momento': (meta or {}).get('momento'),
        'Status_Extracao': status,
        'Linhas_Extraidas': int(rows_extracted),
        'Paginas_PDF': len(pages),
        'Paginas_Com_Texto': text_pages,
        'Tabelas_PDF_Encontradas': table_count,
        'Aviso_OCR': ('PDF sem texto selecionável; OCR não foi aplicado. Conferir o original.' if pages and text_pages == 0 else ''),
        'Erro_Leitura': error or content['error'],
        'Paginas_Fonte': pages,      # texto/tabelas literais por página (exportados no JSON de auditoria)
    }


def update_file_audit(record: dict) -> None:
    """Uma entrada por hash (dicionário: O(1), sem recopiar DataFrames a cada arquivo)."""
    store = st.session_state['audit_store']
    sha = record.get('SHA256_Arquivo')
    store.pop(sha, None)             # mantém a semântica anterior: a versão nova vai para o fim
    store[sha] = record
    st.session_state['audit_version'] += 1


def audit_summary() -> pd.DataFrame:
    """Campos de controle de todos os PDFs (sem os payloads extensos de texto/tabelas)."""
    rows = [{k: v for k, v in rec.items() if k != 'Paginas_Fonte'}
            for rec in st.session_state['audit_store'].values()]
    return pd.DataFrame(rows)


def audit_json_bytes() -> bytes:
    """Empacota JSON de auditoria por arquivo, incluindo texto e tabelas por página."""
    def safe(value):
        if isinstance(value, dict):
            return {str(k): safe(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [safe(v) for v in value]
        if isinstance(value, np.generic):
            value = value.item()
        if value is None or (isinstance(value, float) and not np.isfinite(value)):
            return None
        return value
    output = []
    for rec in st.session_state['audit_store'].values():
        row = {k: v for k, v in rec.items() if k != 'Paginas_Fonte'}
        row['Paginas_Fonte'] = rec.get('Paginas_Fonte', [])
        output.append(safe(row))
    return json.dumps(output, ensure_ascii=False, indent=2, allow_nan=False, default=str).encode('utf-8')


def audit_json_bytes_cached() -> bytes:
    """JSON de auditoria memoizado por versão do armazém (não reserializa a cada interação)."""
    version = st.session_state['audit_version']
    cached = st.session_state.get('audit_json_cache')
    if cached and cached[0] == version:
        return cached[1]
    data = audit_json_bytes()
    st.session_state['audit_json_cache'] = (version, data)
    return data

def find_all_tables_with_headers(tables, table_pages, must_have):
    found = []
    for table in tables:
        if not table or len(table) < 2:
            continue
        header = [_norm(cell or '') for cell in table[0]]
        if all(any(_norm(term) in cell for cell in header) for term in must_have):
            found.append((table, header, table_pages.get(id(table))))
    return found

def col_index(header, keys):
    for idx, h in enumerate(header):
        if any(k in h for k in keys):
            return idx
    return None

def extract_identification_block(full_text: str) -> dict:
    """Extrai metadados comuns a ambos os tipos de relatório (gerados pela mesma suíte)."""
    out = {}
    m = re.search(r'Arquivo analisado\s*[:\n]?\s*([^\n]+)', full_text)
    out['arquivo_txt_original'] = m.group(1).strip() if m else None
    m = re.search(r'Dura[çc][ãa]o do registro\s*[:\n]?\s*([\d.,]+)\s*s', full_text)
    out['duracao_s'] = clean_float(m.group(1)) if m else np.nan
    m = re.search(r'([a-f0-9]{64})', full_text)
    out['sha256'] = m.group(1) if m else None
    return out

def extract_quality_alert(full_text: str) -> str:
    """Retorna literalmente o alerta; seção ausente não significa ausência de alertas."""
    m = re.search(r'4\.\s*Controle de qualidade\s*(.*?)\s*Auditoria espectral', full_text, re.DOTALL | re.IGNORECASE)
    if not m:
        return 'Seção de controle de qualidade não localizada — conferir PDF'
    txt = re.sub(r'\s+', ' ', m.group(1)).strip()
    if not txt:
        return 'Seção localizada, mas sem texto extraível — conferir PDF'
    if 'nenhum aviso' in _norm(txt):
        return 'Nenhum'
    return txt


MECH_PATTERNS = [
    (['forca', 'media', 'janela'], 'ForcaMedia'),
    (['pico', 'forca'], 'Pico'),
    (['instante', 'onset'], 'Onset'),
    (['instante', 'pico'], 'InstPico'),
    (['tempo', 'onset'], 'TempoOnsetPico'),
    (['rfd', '50'], 'RFD50'),
    (['rfd', '100'], 'RFD100'),
    (['rfd', '200'], 'RFD200'),
]

CIVM_LABELS = {
    'Pico_D': 'Pico de força D (kgf)', 'Pico_E': 'Pico de força E (kgf)', 'Assim_Pico_pct': 'Assimetria pico D-E (%)',
    'ForcaMedia_D': 'Força média na janela D (kgf)', 'ForcaMedia_E': 'Força média na janela E (kgf)',
    'Assim_ForcaMedia_pct': 'Assimetria força média D-E (%)',
    'Onset_D': 'Onset D (s)', 'Onset_E': 'Onset E (s)',
    'InstPico_D': 'Instante do pico D (s)', 'InstPico_E': 'Instante do pico E (s)',
    'TempoOnsetPico_D': 'Tempo onset-pico D (ms)', 'TempoOnsetPico_E': 'Tempo onset-pico E (ms)',
    'RFD50_D': 'RFD 0-50ms D (kgf/s)', 'RFD50_E': 'RFD 0-50ms E (kgf/s)', 'Assim_RFD50_pct': 'Assimetria RFD 0-50ms (%)',
    'RFD100_D': 'RFD 0-100ms D (kgf/s)', 'RFD100_E': 'RFD 0-100ms E (kgf/s)', 'Assim_RFD100_pct': 'Assimetria RFD 0-100ms (%)',
    'RFD200_D': 'RFD 0-200ms D (kgf/s)', 'RFD200_E': 'RFD 0-200ms E (kgf/s)', 'Assim_RFD200_pct': 'Assimetria RFD 0-200ms (%)',
}


def pretty_label(col: str) -> str:
    """Converte nomes técnicos em rótulos de interface legíveis."""
    if col in CIVM_LABELS:
        return CIVM_LABELS[col]
    value = col.replace('Fadiga_', '[Fadiga] ')
    value = re.sub(r'_D(_|$)', r' Direito ', value)
    value = re.sub(r'_E(_|$)', r' Esquerdo ', value)
    value = value.replace('_pct', ' (%)').replace('_mV', ' (mV)').replace('_dB', ' (dB)')
    value = value.replace('_media', ' (média)').replace('_pior', ' (pior caso)')
    value = value.replace('Assim_', 'Assimetria ').replace('_', ' ').strip()
    return value


def parse_civm_pdf(file_bytes: bytes, filename: str, meta: dict) -> dict | None:
    """
    Extrai um relatório de CIVM (força + EMG) em um único dicionário (1 linha, formato largo),
    replicando a estrutura real do relatório: mecânica (força), eletromiografia bilateral por
    músculo (RMS/ENM), auditoria espectral, alertas de qualidade e rastreabilidade (SHA-256).
    """
    content = extract_pdf_content(file_bytes)          # leitura única (cacheada)
    if content['error']:
        st.error(f"Falha ao ler '{filename}': {content['error']}")
        return None
    full_text, all_tables, table_pages = collect_tables(content)

    row = {
        'ID_Voluntario': meta['id'], 'Condicao': meta['condicao'], 'Momento': meta['momento'],
        'Arquivo_PDF': os.path.basename(filename),
        'SHA256_Arquivo': content['sha256'],
    }
    ident = extract_identification_block(full_text)
    row['Arquivo_txt_original'] = ident['arquivo_txt_original']
    row['Duracao_s'] = ident['duracao_s']
    row['SHA256'] = ident['sha256']
    row['Nome_Consistente'] = None
    if row['Arquivo_txt_original']:
        prefixo_esperado = 'CIVMC' if meta['condicao'] == 'Controle' else 'CIVMI'
        row['Nome_Consistente'] = row['Arquivo_txt_original'].upper().startswith(prefixo_esperado)

    m = re.search(r'Janela de CIVM\s*[:\n]?\s*([\d.,]+)\s*s\s*a\s*([\d.,]+)\s*s', full_text)
    row['Janela_Inicio_s'] = clean_float(m.group(1)) if m else np.nan
    row['Janela_Fim_s'] = clean_float(m.group(2)) if m else np.nan
    row['Alerta_Qualidade'] = extract_quality_alert(full_text)

    # --- tabela mecânica (força) ---
    mech_table, mech_header = find_table_with_headers(all_tables, ['vari', 'direito', 'esquerdo'])
    row['Tabela_Mecanica_Encontrada'] = mech_table is not None
    row['Pagina_Tabela_Mecanica'] = table_pages.get(id(mech_table)) if mech_table is not None else np.nan
    if mech_table:
        idx_var, idx_d, idx_e = col_index(mech_header, ['vari']), col_index(mech_header, ['direito']), col_index(mech_header, ['esquerdo'])
        idx_asym = col_index(mech_header, ['assimetria'])
        for r in mech_table[1:]:
            if not r or idx_var is None:
                continue
            label = _norm(r[idx_var])
            key = next((k for req, k in MECH_PATTERNS if all(s in label for s in req)), None)
            if key is None:
                continue
            row[f'{key}_D'] = clean_float(r[idx_d]) if idx_d is not None and idx_d < len(r) else np.nan
            row[f'{key}_E'] = clean_float(r[idx_e]) if idx_e is not None and idx_e < len(r) else np.nan
            if idx_asym is not None and idx_asym < len(r):
                row[f'Assim_{key}_pct'] = clean_float(r[idx_asym])

    # --- "Comparação bilateral por músculo" (RMS / ENM) ---
    bil_table, bil_header = find_table_with_headers(all_tables, ['musculo', 'rms dir', 'enm dir'])
    row['Tabela_Bilateral_Encontrada'] = bil_table is not None
    row['Pagina_Tabela_Bilateral'] = table_pages.get(id(bil_table)) if bil_table is not None else np.nan
    if bil_table:
        idx_m = col_index(bil_header, ['musculo'])
        idx_rd, idx_re = col_index(bil_header, ['rms dir']), col_index(bil_header, ['rms esq'])
        idx_asrms = col_index(bil_header, ['assimetria rms'])
        idx_ed, idx_ee = col_index(bil_header, ['enm dir']), col_index(bil_header, ['enm esq'])
        idx_asenm = col_index(bil_header, ['assimetria enm'])
        for r in bil_table[1:]:
            if not r or idx_m is None or not str(r[idx_m]).strip():
                continue
            mk = _norm(r[idx_m]).split()[0].capitalize()
            row[f'RMS_{mk}_D_mV'] = clean_float(r[idx_rd]) if idx_rd is not None else np.nan
            row[f'RMS_{mk}_E_mV'] = clean_float(r[idx_re]) if idx_re is not None else np.nan
            row[f'Assim_RMS_{mk}_pct'] = clean_float(r[idx_asrms]) if idx_asrms is not None else np.nan
            row[f'ENM_{mk}_D'] = clean_float(r[idx_ed]) if idx_ed is not None else np.nan
            row[f'ENM_{mk}_E'] = clean_float(r[idx_ee]) if idx_ee is not None else np.nan
            row[f'Assim_ENM_{mk}_pct'] = clean_float(r[idx_asenm]) if idx_asenm is not None else np.nan

    # --- auditoria espectral (atenuação em 60 Hz) ---
    aten_table, aten_header = find_table_with_headers(all_tables, ['canal', 'atenua'])
    row['Tabela_Atenuacao_Encontrada'] = aten_table is not None
    row['Pagina_Tabela_Atenuacao'] = table_pages.get(id(aten_table)) if aten_table is not None else np.nan
    if aten_table:
        idx_c, idx_v = col_index(aten_header, ['canal']), col_index(aten_header, ['atenua'])
        for r in aten_table[1:]:
            if not r or idx_c is None:
                continue
            canal = _norm(r[idx_c])
            if not canal:
                continue
            side = 'E' if 'esquerd' in canal else 'D'
            mk = canal.split()[0].capitalize()
            row[f'Atenuacao60Hz_{mk}_{side}_dB'] = clean_float(r[idx_v]) if idx_v is not None else np.nan

    identity = {'ID_Voluntario', 'Condicao', 'Momento', 'Arquivo_PDF', 'SHA256_Arquivo',
                'Arquivo_txt_original', 'Duracao_s', 'SHA256', 'Nome_Consistente',
                'Alerta_Qualidade', 'Tabela_Mecanica_Encontrada', 'Pagina_Tabela_Mecanica',
                'Tabela_Bilateral_Encontrada', 'Pagina_Tabela_Bilateral',
                'Tabela_Atenuacao_Encontrada', 'Pagina_Tabela_Atenuacao'}
    row['N_Parametros_Extraidos'] = sum(pd.notna(v) for k, v in row.items() if k not in identity)
    return row

# ==============================================================================
# MÓDULO 3: PARSER DO RELATÓRIO DE FADIGA (FORMATO LONGO: 1 LINHA / MÚSCULO)
# ==============================================================================

def parse_fadiga_pdf(file_bytes: bytes, filename: str, meta: dict) -> pd.DataFrame:
    """
    Extrai parâmetros de fadiga espectral por músculo: MDF inicial/final, MDF Slope,
    Queda de MDF (%), RMS Slope, R² e p-valor — a partir da tabela de resultados do
    relatório PDF de fadiga isométrica sustentada.
    """
    content = extract_pdf_content(file_bytes)          # leitura única (cacheada)
    if content['error']:
        st.error(f"Falha ao ler '{filename}': {content['error']}")
        return pd.DataFrame()
    full_text, all_tables, table_pages = collect_tables(content)
    file_sha = content['sha256']

    ident = extract_identification_block(full_text)
    alerta = extract_quality_alert(full_text)

    records = []
    matched_tables = find_all_tables_with_headers(all_tables, table_pages, ['mdf', 'musculo'])
    if not matched_tables:
        matched_tables = find_all_tables_with_headers(all_tables, table_pages, ['mdf', 'canal'])
    for table, header, page_no in matched_tables:
        col_map = {
            'musculo': col_index(header, ['musculo', 'canal']),
            'mdf_inicial': col_index(header, ['mdf inicial']),
            'mdf_final': col_index(header, ['mdf final']),
            'mdf_slope': col_index(header, ['mdf slope', 'inclin']),
            'queda_mdf': col_index(header, ['queda', 'delta']),
            'rms_slope': col_index(header, ['rms slope']),
            'r2': col_index(header, ['r2', 'r²']),
            'p_val': col_index(header, ['p-valor', 'p valor', 'p-val']),
        }
        # Constantes por tabela calculadas UMA vez (antes eram recalculadas a cada linha/músculo).
        cols_json = json.dumps({k: header[v] if v is not None and v < len(header) else None for k, v in col_map.items()},
                               ensure_ascii=False)
        metric_idx = {k: idx for k, idx in col_map.items() if k != 'musculo'}
        for r in table[1:]:
            if not r or col_map['musculo'] is None:
                continue
            musc_raw = str(r[col_map['musculo']] or '').strip() if col_map['musculo'] < len(r) else ''
            if not musc_raw or _norm(musc_raw).startswith('musculo'):
                continue
            vals = {k: (clean_float(r[idx]) if idx is not None and idx < len(r) else np.nan)
                    for k, idx in metric_idx.items()}
            records.append({
                'ID_Voluntario': meta['id'], 'Condicao': meta['condicao'], 'Momento': meta['momento'],
                'Arquivo_PDF': os.path.basename(filename), 'SHA256_Arquivo': file_sha,
                'Pagina_Extracao': page_no,
                'Arquivo_txt_original': ident['arquivo_txt_original'],
                'Duracao_s': ident['duracao_s'], 'SHA256_Relatorio': ident['sha256'],
                'Alerta_Qualidade': alerta,
                'Musculo': standardize_muscle_name(musc_raw),
                'MDF_Inicial_Hz': vals['mdf_inicial'], 'MDF_Final_Hz': vals['mdf_final'],
                'MDF_Slope_Hz_s': vals['mdf_slope'], 'Queda_MDF_pct': vals['queda_mdf'],
                'RMS_Slope_uV_s': vals['rms_slope'], 'MDF_R2': vals['r2'], 'MDF_p_valor': vals['p_val'],
                'Colunas_Encontradas_JSON': cols_json,
                'N_Parametros_Extraidos': sum(pd.notna(v) for v in vals.values()),
            })
    return pd.DataFrame(records)

def process_pdf_upload(file_bytes: bytes, filename: str, meta: dict, tipo: str):
    """Executa parser e registra auditoria por arquivo; nenhum erro é ignorado."""
    civm_record, fadiga_frame, error = None, pd.DataFrame(), ''
    try:
        if tipo == 'CIVM':
            civm_record = parse_civm_pdf(file_bytes, filename, meta)
            rows = int(civm_record is not None)
            if civm_record is None:
                status = 'SEM_DADOS_COMPATIVEIS'
            elif int(civm_record.get('N_Parametros_Extraidos', 0) or 0) == 0:
                status = 'PARCIAL — nenhuma métrica numérica reconhecida'
            elif not civm_record.get('Tabela_Mecanica_Encontrada') or not civm_record.get('Tabela_Bilateral_Encontrada'):
                status = 'PARCIAL — tabela(s) esperada(s) não localizada(s)'
            else:
                status = 'EXTRAIDO — revisar auditoria'
        elif tipo == 'FADIGA':
            fadiga_frame = parse_fadiga_pdf(file_bytes, filename, meta)
            rows = len(fadiga_frame)
            if not rows:
                status = 'SEM_DADOS_COMPATIVEIS'
            elif int(fadiga_frame.get('N_Parametros_Extraidos', pd.Series(dtype=float)).sum()) == 0:
                status = 'PARCIAL — identificadores extraídos, nenhuma métrica numérica reconhecida'
            else:
                status = 'EXTRAIDO — revisar auditoria'
        else:
            rows, status = 0, 'TIPO_NAO_IDENTIFICADO'
    except Exception as exc:
        rows, status, error = 0, 'ERRO_NO_PARSER', str(exc)
        civm_record, fadiga_frame = None, pd.DataFrame()
    audit = pdf_audit_record(extract_pdf_content(file_bytes), filename, meta, tipo, status, rows, error)
    update_file_audit(audit)
    return civm_record, fadiga_frame, audit


def known_processed_hash(file_bytes: bytes, digest: str | None = None) -> bool:
    """Impede reimportação acidental do mesmo conteúdo binário (consulta O(1) no armazém de auditoria)."""
    digest = digest or hashlib.sha256(file_bytes).hexdigest()
    rec = st.session_state['audit_store'].get(digest)
    if rec is None:
        return False
    status = str(rec.get('Status_Extracao', ''))
    return status.startswith('EXTRAIDO') or status.startswith('PARCIAL')


def aggregate_fadiga_to_wide(df_fadiga_long: pd.DataFrame) -> pd.DataFrame:
    """Resume a fadiga (1 linha/músculo) em 1 linha por teste (média entre os canais)."""
    if df_fadiga_long.empty:
        return pd.DataFrame(columns=['ID_Voluntario', 'Condicao', 'Momento'])
    source = df_fadiga_long.copy()
    for optional in ['Arquivo_PDF', 'SHA256_Arquivo']:
        if optional not in source.columns:
            source[optional] = np.nan
    agg = source.groupby(['ID_Voluntario', 'Condicao', 'Momento'], as_index=False).agg(
        Fadiga_MDF_Inicial_Hz_media=('MDF_Inicial_Hz', 'mean'),
        Fadiga_MDF_Final_Hz_media=('MDF_Final_Hz', 'mean'),
        Fadiga_MDF_Slope_Hz_s_media=('MDF_Slope_Hz_s', 'mean'),
        Fadiga_MDF_Slope_Hz_s_pior=('MDF_Slope_Hz_s', 'min'),
        Fadiga_Queda_MDF_pct_media=('Queda_MDF_pct', 'mean'),
        Fadiga_RMS_Slope_uV_s_media=('RMS_Slope_uV_s', 'mean'),
        Fadiga_N_Musculos=('Musculo', 'nunique'),
        Fadiga_Arquivos_Origem=('Arquivo_PDF', lambda values: ' | '.join(sorted(set(str(v) for v in values.dropna())))),
        Fadiga_Hashes_Origem=('SHA256_Arquivo', lambda values: ' | '.join(sorted(set(str(v) for v in values.dropna())))),
    )
    return agg

# ==============================================================================
# MÓDULO 4: FUSÃO (TABELA MESTRA) E COMPARAÇÃO CROSSOVER
# ==============================================================================

def build_master_table(df_civm: pd.DataFrame, df_fadiga_long: pd.DataFrame) -> pd.DataFrame:
    """Combina apenas chaves únicas; CIVM duplicado permanece visível, sem pareamento inventado."""
    keys = ['ID_Voluntario', 'Condicao', 'Momento']
    fad = aggregate_fadiga_to_wide(df_fadiga_long)
    if df_civm.empty and fad.empty:
        return pd.DataFrame()
    civm = df_civm.copy()
    ambiguous = pd.DataFrame()
    if not civm.empty:
        duplicate_mask = civm.duplicated(keys, keep=False)
        if duplicate_mask.any():
            ambiguous = civm.loc[duplicate_mask].copy()
            ambiguous['Fonte_Dados'] = 'CIVM duplicado — fadiga não pareada automaticamente'
            ambiguous['Flag_Revisao'] = 'Mais de um PDF CIVM na mesma chave de voluntário/condição/momento'
            ambiguous_keys = set(map(tuple, ambiguous[keys].itertuples(index=False, name=None)))
            civm = civm.loc[~duplicate_mask].copy()
            if not fad.empty:
                fad_keys = fad[keys].apply(tuple, axis=1)
                fad = fad.loc[~fad_keys.isin(ambiguous_keys)].copy()
            # Garante alinhamento de colunas ao concatenar e conserva todas as linhas repetidas.
            for col in fad.columns:
                if col not in ambiguous.columns:
                    ambiguous[col] = np.nan
    if civm.empty:
        merged = fad.copy()
        if not merged.empty:
            merged['Fonte_Dados'] = 'Somente Fadiga'
    elif fad.empty:
        merged = civm.copy()
        merged['Fonte_Dados'] = 'Somente CIVM'
    else:
        merged = pd.merge(civm, fad, on=keys, how='outer', indicator=True, validate='one_to_one')
        merged['Fonte_Dados'] = merged['_merge'].map({'both': 'CIVM + Fadiga', 'left_only': 'Somente CIVM', 'right_only': 'Somente Fadiga'})
        merged.drop(columns=['_merge'], inplace=True)
    if not ambiguous.empty:
        merged = pd.concat([merged, ambiguous], ignore_index=True, sort=False)
    front = keys + ['Fonte_Dados', 'Arquivo_PDF', 'Fadiga_Arquivos_Origem', 'Arquivo_txt_original',
                    'Nome_Consistente', 'Alerta_Qualidade', 'Flag_Revisao']
    front = [c for c in front if c in merged.columns]
    rest = [c for c in merged.columns if c not in front]
    return merged[front + rest].sort_values(keys).reset_index(drop=True)

def numeric_metric_columns(df_master: pd.DataFrame) -> list:
    """Lista as colunas numéricas de métrica (exclui IDs, flags e metadados)."""
    exclude = {'Duracao_s', 'Janela_Inicio_s', 'Janela_Fim_s', 'Nome_Consistente', 'Fadiga_N_Musculos'}
    cols = []
    for c in df_master.columns:
        if c in exclude or c in ('ID_Voluntario', 'Condicao', 'Momento'):
            continue
        if c.startswith(('Pagina_', 'N_Parametros_', 'Tabela_')) or pd.api.types.is_bool_dtype(df_master[c]):
            continue
        if pd.api.types.is_numeric_dtype(df_master[c]):
            cols.append(c)
    return cols

def build_comparison_table(df_master: pd.DataFrame, metric_cols: list,
                            subject_col: str = 'ID_Voluntario') -> pd.DataFrame:
    """
    Para cada métrica, calcula média±DP em Controle-Pré/Pós e Intervenção-Pré/Pós,
    e testa (t pareado + Wilcoxon) se a variação (Pós−Pré) difere entre as condições,
    pareando pelo voluntário (desenho crossover). Equivalente à "Tabela 1" do laudo.
    """
    rows = []
    scipy_stats = _lazy('scipy.stats') if SCIPY_AVAILABLE else None   # import preguiçoso
    key_cols = [subject_col, 'Condicao', 'Momento']
    if df_master.duplicated(key_cols, keep=False).any():
        return pd.DataFrame([{'Aviso': 'Comparação bloqueada: duplicatas na chave voluntário×condição×momento; revise os arquivos.'}])
    try:
        wide = df_master.pivot_table(index=subject_col, columns=['Condicao', 'Momento'], values=metric_cols)
    except Exception:
        return pd.DataFrame()

    for var in metric_cols:
        try:
            cp, co = wide[(var, 'Controle', 'Pré')], wide[(var, 'Controle', 'Pós')]
            ip, io_ = wide[(var, 'Intervenção', 'Pré')], wide[(var, 'Intervenção', 'Pós')]
        except KeyError:
            continue
        paired = pd.concat([cp, co, ip, io_], axis=1).dropna()
        n_pares = len(paired)
        p_t = p_w = np.nan
        if n_pares >= 3:
            d_c = (co - cp).reindex(paired.index)
            d_i = (io_ - ip).reindex(paired.index)
            try:
                _, p_t = scipy_stats.ttest_rel(d_i, d_c)
            except Exception:
                pass
            try:
                if not np.allclose(d_i, d_c):
                    _, p_w = scipy_stats.wilcoxon(d_i, d_c)
            except Exception:
                pass
        rows.append({
            'Variável': pretty_label(var), 'coluna': var, 'N pares': n_pares,
            'Controle - Pré': f"{cp.mean():.2f} ± {cp.std():.2f}" if cp.notna().any() else "—",
            'Controle - Pós': f"{co.mean():.2f} ± {co.std():.2f}" if co.notna().any() else "—",
            'Intervenção - Pré': f"{ip.mean():.2f} ± {ip.std():.2f}" if ip.notna().any() else "—",
            'Intervenção - Pós': f"{io_.mean():.2f} ± {io_.std():.2f}" if io_.notna().any() else "—",
            'p (t pareado)': f"{p_t:.3f}" if pd.notna(p_t) else "n.s.i.",
            'p (Wilcoxon)': f"{p_w:.3f}" if pd.notna(p_w) else "—",
        })
    return pd.DataFrame(rows)

@st.cache_data(show_spinner=False, max_entries=16)
def _cached_comparison(df_num: pd.DataFrame, metric_cols: tuple) -> pd.DataFrame:
    return build_comparison_table(df_num, list(metric_cols))

def cached_comparison_table(df_master: pd.DataFrame):
    """Tabela comparativa cacheada (só colunas numéricas entram na chave). Retorna (tabela, colunas)."""
    metric_cols = numeric_metric_columns(df_master)
    keys = ['ID_Voluntario', 'Condicao', 'Momento']
    return _cached_comparison(df_master[keys + metric_cols], tuple(metric_cols)), metric_cols

# ==============================================================================
# MÓDULO 5: VISUALIZAÇÃO — BARRAS + TRAJETÓRIAS INDIVIDUAIS ("ESPAGUETE")
# ==============================================================================

def _render_png(fig, dpi: int = 160) -> bytes:
    """Serializa a figura em PNG e SEMPRE a fecha (evita vazamento de memória entre reruns)."""
    plt = _plt()
    buf = io.BytesIO()
    try:
        fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight')
    finally:
        plt.close(fig)
    return buf.getvalue()

def png_to_base64(png: bytes) -> str:
    return base64.b64encode(png).decode('utf-8')

def _mean_ci95(values) -> tuple:
    """Média e meia-largura do IC95% (t de Student, analítico — determinístico e sem bootstrap)."""
    x = pd.Series(values).dropna().to_numpy(dtype=float)
    n = len(x)
    if n == 0:
        return np.nan, np.nan
    mean = float(x.mean())
    if n < 2:
        return mean, 0.0
    sem = float(x.std(ddof=1) / np.sqrt(n))
    sp = _lazy('scipy.stats')
    tcrit = float(sp.t.ppf(0.975, n - 1)) if sp is not None else 1.96
    return mean, tcrit * sem

@st.cache_data(show_spinner=False, max_entries=64)
def render_bar_spaghetti_png(df_plot: pd.DataFrame, value_col: str, subject_col: str = 'ID_Voluntario') -> bytes:
    """Barras (média±DP) + trajetórias individuais Pré→Pós, painéis Controle | Intervenção. Retorna PNG (cacheado)."""
    plt = _plt()
    vols = sorted(df_plot[subject_col].dropna().unique())
    marker_cycle = dict(zip(vols, itertools.cycle(MARKERS)))
    with plt.rc_context({'axes.spines.top': False, 'axes.spines.right': False, 'axes.grid': False}):
        fig, axes = plt.subplots(1, 2, figsize=(9, 4.6), sharey=True)
        for ax, cond in zip(axes, ["Controle", "Intervenção"]):
            sub = df_plot[df_plot['Condicao'] == cond]
            pre = sub[sub['Momento'] == 'Pré'].set_index(subject_col)[value_col].reindex(vols)
            pos = sub[sub['Momento'] == 'Pós'].set_index(subject_col)[value_col].reindex(vols)
            means = [pre.mean(), pos.mean()]
            sds = [pre.std(), pos.std()]
            ax.bar([0, 1], means, yerr=sds, capsize=5, color=COND_COLORS[cond], alpha=0.35, width=0.5, zorder=1)
            for v in vols:
                if pd.notna(pre.get(v)) and pd.notna(pos.get(v)):
                    ax.plot([0, 1], [pre[v], pos[v]], color="gray", marker=marker_cycle[v], markersize=6,
                            linewidth=1.1, alpha=0.9, zorder=2, label=f"Vol. {v}" if cond == "Controle" else None)
            ax.set_xticks([0, 1]); ax.set_xticklabels(["Pré", "Pós"])
            ax.set_title(cond, fontsize=12, fontweight='bold')
            ax.set_xlim(-0.5, 1.5)
        axes[0].set_ylabel(pretty_label(value_col))
        handles, labels = axes[0].get_legend_handles_labels()
        if handles:
            fig.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, 1.08),
                       ncol=min(len(vols), 8), frameon=False, fontsize=8)
        fig.tight_layout()
        return _render_png(fig)

# ==============================================================================
# MÓDULO 6: ANÁLISE ESTATÍSTICA (PINGOUIN ANOVA TWO-WAY RM)
# ==============================================================================

def run_repeated_measures_statistics(df_master: pd.DataFrame, civm_col: str, fadiga_col: str) -> dict:
    """
    ANOVA de Medidas Repetidas (Condição × Momento, sujeito = ID_Voluntario) para a
    métrica de CIVM e para a métrica de Fadiga selecionadas, + correlação entre ambas.
    """
    stats_out = {}
    pg = _lazy('pingouin') if PINGOUIN_AVAILABLE else None      # import pesado só quando necessário
    if pg is None:
        stats_out['analysis_blocked'] = 'Pingouin não está instalado (ou falhou ao importar); a ANOVA não foi executada.'
        stats_out['processed_data'] = df_master.copy()
        return stats_out
    keys = ['ID_Voluntario', 'Condicao', 'Momento']
    if df_master.duplicated(keys, keep=False).any():
        stats_out['analysis_blocked'] = 'Há mais de uma observação na mesma chave voluntário×condição×momento. Corrija/revise antes da ANOVA; nenhum valor foi agregado automaticamente.'
        stats_out['processed_data'] = df_master.copy()
        return stats_out
    cols_needed = ['ID_Voluntario', 'Condicao', 'Momento']
    sub = df_master[cols_needed + [c for c in dict.fromkeys([civm_col, fadiga_col]) if c in df_master.columns]].copy()

    if civm_col in sub.columns:
        try:
            aov = pg.rm_anova(data=sub.dropna(subset=[civm_col]), dv=civm_col,
                               within=['Condicao', 'Momento'], subject='ID_Voluntario', detailed=True)
            stats_out['aov_civm'] = aov
        except Exception as e:
            stats_out['aov_civm_err'] = str(e)

    if fadiga_col in sub.columns:
        try:
            aov = pg.rm_anova(data=sub.dropna(subset=[fadiga_col]), dv=fadiga_col,
                               within=['Condicao', 'Momento'], subject='ID_Voluntario', detailed=True)
            stats_out['aov_fatigue'] = aov
        except Exception as e:
            stats_out['aov_fatigue_err'] = str(e)

    try:
        both = sub.dropna(subset=[civm_col, fadiga_col])
        corr_res = pg.corr(both[civm_col], both[fadiga_col], method='pearson')
        stats_out['correlation'] = corr_res
    except Exception:
        pass

    stats_out['processed_data'] = sub
    return stats_out

@st.cache_data(show_spinner=False, max_entries=16)
def _cached_rm_stats(sub_df: pd.DataFrame, civm_col: str, fadiga_col: str) -> dict:
    return run_repeated_measures_statistics(sub_df, civm_col, fadiga_col)

def cached_rm_stats(df_master: pd.DataFrame, civm_col: str, fadiga_col: str) -> dict:
    """ANOVA cacheada: só recalcula quando os dados ou as variáveis escolhidas mudam."""
    cols = ['ID_Voluntario', 'Condicao', 'Momento'] + [c for c in dict.fromkeys([civm_col, fadiga_col]) if c in df_master.columns]
    return _cached_rm_stats(df_master[cols], civm_col, fadiga_col)

@st.cache_data(show_spinner=False, max_entries=32)
def render_statistical_plots_png(sub_df: pd.DataFrame, civm_col: str, fadiga_col: str) -> bytes:
    """Interação (CIVM e Fadiga) + dispersão/correlação. IC95% analítico (t), sem seaborn/bootstrap."""
    plt = _plt()
    with plt.rc_context({'axes.grid': True, 'grid.alpha': 0.3,
                         'axes.spines.top': False, 'axes.spines.right': False}):
        fig, axes = plt.subplots(1, 3, figsize=(17, 4.6))
        panels = ((axes[0], civm_col, "Capacidade de força (CIVM)"),
                  (axes[1], fadiga_col, "Resistência à fadiga"))
        for ax, col, title in panels:
            for j, (cond, marker, ls) in enumerate((('Controle', 'o', '-'), ('Intervenção', 's', '--'))):
                xs, ms, cis = [], [], []
                for i, mom in enumerate(['Pré', 'Pós']):
                    vals = sub_df[(sub_df['Condicao'] == cond) & (sub_df['Momento'] == mom)][col]
                    m, ci = _mean_ci95(vals)
                    xs.append(i + (j - 0.5) * 0.08); ms.append(m); cis.append(ci)
                ax.errorbar(xs, ms, yerr=cis, marker=marker, linestyle=ls, capsize=4, linewidth=1.5,
                            color=COND_COLORS[cond], label=cond)
            ax.set_xticks([0, 1]); ax.set_xticklabels(['Pré', 'Pós']); ax.set_xlim(-0.4, 1.4)
            ax.set_title(f"{title}\n{pretty_label(col)}", fontsize=10, fontweight='bold')
            ax.set_ylabel(pretty_label(col)); ax.set_xlabel("Momento")
            ax.legend(title='Condição', fontsize=8, frameon=False)

        ax3 = axes[2]
        both = sub_df.dropna(subset=[civm_col, fadiga_col])
        if len(both) >= 2:
            x = both[civm_col].to_numpy(dtype=float)
            y = both[fadiga_col].to_numpy(dtype=float)
            ax3.scatter(x, y, alpha=0.7, s=50, color='#2563EB')
            if np.ptp(x) > 0:
                slope, intercept = np.polyfit(x, y, 1)
                xx = np.linspace(x.min(), x.max(), 50)
                ax3.plot(xx, slope * xx + intercept, color='darkred', linewidth=2)
            corr_val = np.corrcoef(x, y)[0, 1]
            ax3.set_title(f"Correlação força × fadiga\n(r = {corr_val:.3f})", fontsize=10, fontweight='bold')
        else:
            ax3.text(0.5, 0.5, "Dados insuficientes\npara correlação", ha='center', va='center')
        ax3.set_xlabel(pretty_label(civm_col)); ax3.set_ylabel(pretty_label(fadiga_col))
        fig.tight_layout()
        return _render_png(fig)

@st.cache_data(show_spinner=False, max_entries=4)
def compile_pdf_bytes(html_doc: str) -> bytes:
    """Compila o laudo em PDF (etapa mais lenta do módulo 5) — cacheado pelo conteúdo do HTML."""
    weasy = _lazy('weasyprint')
    if weasy is None:
        raise RuntimeError("WeasyPrint indisponível neste ambiente.")
    return weasy.HTML(string=html_doc).write_pdf()

@st.cache_data(show_spinner=False, max_entries=8)
def build_master_xlsx(df_tidy: pd.DataFrame, df_civm: pd.DataFrame, df_fad: pd.DataFrame, df_audit: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_tidy.to_excel(writer, index=False, sheet_name='Tabela_Mestra')
        df_civm.to_excel(writer, index=False, sheet_name='CIVM_Original')
        df_fad.to_excel(writer, index=False, sheet_name='Fadiga_por_musculo')
        df_audit.to_excel(writer, index=False, sheet_name='Auditoria_Arquivos')
    return buf.getvalue()

@st.cache_data(show_spinner=False, max_entries=8)
def build_comparison_xlsx(df_comp: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_comp.to_excel(writer, index=False, sheet_name='Comparacao')
    return buf.getvalue()

@_fragment
def comparison_chart_section(df_tidy: pd.DataFrame, metric_cols: list):
    """Seletor + gráfico. Como fragmento, trocar a variável NÃO reexecuta a página inteira."""
    if not metric_cols:
        st.info("Nenhuma métrica numérica disponível para o gráfico.")
        return
    var_options = {pretty_label(c): c for c in metric_cols}
    var_label = st.selectbox("Escolha a variável para visualizar:", sorted(var_options.keys()), key='cmp_var')
    chosen_col = var_options[var_label]
    df_plot = df_tidy[['ID_Voluntario', 'Condicao', 'Momento', chosen_col]]
    if df_plot.duplicated(['ID_Voluntario', 'Condicao', 'Momento']).any():
        st.warning("Gráfico bloqueado: há mais de uma observação na mesma chave voluntário×condição×momento (revise os PDFs duplicados).")
        return
    png = render_bar_spaghetti_png(df_plot, chosen_col)
    st.image(png)
    st.session_state['last_chart_b64'] = png_to_base64(png)

# ==============================================================================
# MÓDULO 7: LAUDO ACADÊMICO WEASYPRINT (HTML + CSS PRINT)
# ==============================================================================

def build_weasyprint_html(df_master: pd.DataFrame, comparison_df: pd.DataFrame, stats_results: dict,
                           b64_chart: str, civm_col: str, fadiga_col: str) -> str:
    aov_civm_html = ""
    if stats_results and 'aov_civm' in stats_results:
        dfc = stats_results['aov_civm'][['Source', 'SS', 'ddof1', 'ddof2', 'F', 'p-unc' if 'p-unc' in stats_results['aov_civm'].columns else 'p_unc', 'ng2']].copy()
        dfc.columns = ['Fonte de Variação', 'SQ', 'GL1', 'GL2', 'F', 'p-valor', 'Eta²g']
        aov_civm_html = dfc.to_html(index=False, classes='styled-table', float_format="%.4f")
    aov_fat_html = ""
    if stats_results and 'aov_fatigue' in stats_results:
        dff = stats_results['aov_fatigue'][['Source', 'SS', 'ddof1', 'ddof2', 'F', 'p-unc' if 'p-unc' in stats_results['aov_fatigue'].columns else 'p_unc', 'ng2']].copy()
        dff.columns = ['Fonte de Variação', 'SQ', 'GL1', 'GL2', 'F', 'p-valor', 'Eta²g']
        aov_fat_html = dff.to_html(index=False, classes='styled-table', float_format="%.4f")

    comp_html = comparison_df.drop(columns=['coluna'], errors='ignore').to_html(index=False, classes='styled-table')

    n_vol = df_master['ID_Voluntario'].nunique()
    n_testes = len(df_master)
    flags = df_master[(df_master.get('Nome_Consistente') == False) | (df_master.get('Alerta_Qualidade', 'Seção não localizada') != 'Nenhum')] \
        if 'Nome_Consistente' in df_master.columns else pd.DataFrame()
    flags_html = ""
    if not flags.empty:
        cols_show = [c for c in ['ID_Voluntario', 'Condicao', 'Momento', 'Arquivo_PDF', 'Nome_Consistente', 'Alerta_Qualidade'] if c in flags.columns]
        flags_html = flags[cols_show].to_html(index=False, classes='styled-table')

    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="utf-8">
        <title>Laudo Biomecânico: Crossover CIVM e Fadiga sEMG</title>
        <style>
            @page {{ size: A4 portrait; margin: 1.5cm;
                @bottom-right {{ content: "Página " counter(page) " de " counter(pages); font-size: 8pt; color: #64748B; }}
                @bottom-left {{ content: "Laboratório de Biomecânica e Eletromiografia"; font-size: 8pt; color: #64748B; }} }}
            body {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #1E293B; line-height: 1.4; font-size: 10pt; }}
            .header-bar {{ border-bottom: 3px solid #2563EB; padding-bottom: 8px; margin-bottom: 16px; }}
            h1 {{ color: #1E3A8A; font-size: 18pt; margin: 0; font-weight: bold; }}
            .subtitle {{ color: #64748B; font-size: 9.5pt; margin-top: 4px; }}
            h2 {{ color: #1E3A8A; font-size: 12pt; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px; margin-top: 16px; margin-bottom: 8px; }}
            .styled-table {{ width: 100%; border-collapse: collapse; margin: 8px 0 14px 0; font-size: 7.6pt; }}
            .styled-table th {{ background-color: #2563EB; color: white; text-align: center; padding: 5px; font-weight: bold; }}
            .styled-table td {{ border: 1px solid #E2E8F0; padding: 4px 6px; text-align: center; }}
            .styled-table tr:nth-child(even) {{ background-color: #F8FAFC; }}
            .chart-img {{ width: 100%; max-width: 100%; height: auto; border: 1px solid #CBD5E1; border-radius: 4px; margin-top: 8px; }}
            .info-card {{ background-color: #F8FAFC; border-left: 4px solid #3B82F6; padding: 8px 12px; font-size: 8.5pt; margin-bottom: 10px; }}
            .warn-card {{ background-color: #FEF2F2; border-left: 4px solid #DC2626; padding: 8px 12px; font-size: 8pt; margin-bottom: 10px; }}
        </style>
    </head>
    <body>
        <div class="header-bar">
            <h1>LAUDO DE ANÁLISE BIOMECÂNICA CROSSOVER</h1>
            <div class="subtitle">Fusão de relatórios de CIVM (força + sEMG) e Fadiga Isométrica Sustentada — Mobilização Miofascial</div>
        </div>
        <div class="info-card">
            <b>Delineamento:</b> Crossover de Medidas Repetidas (Condição: Controle × Intervenção; Momento: Pré × Pós).<br>
            <b>Voluntários:</b> {n_vol} | <b>Testes consolidados na Tabela Mestra:</b> {n_testes} |
            <b>Variável CIVM em foco:</b> {pretty_label(civm_col)} | <b>Variável de Fadiga em foco:</b> {pretty_label(fadiga_col)}
        </div>
        {f'<div class="warn-card"><b>⚠ Atenção — {len(flags)} teste(s) com inconsistência de nome de arquivo ou alerta automático de controle de qualidade:</b>{flags_html}</div>' if not flags.empty else ''}

        <h2>1. Tabela Comparativa — Controle × Intervenção, Pré × Pós (média ± DP)</h2>
        {comp_html}
        <p style="font-size:7.5pt; color:#64748B;">p calculado pelo teste t pareado (e Wilcoxon) comparando a variação Pós−Pré entre Intervenção e Controle, pareado por voluntário. "n.s.i." = n insuficiente (&lt; 3 pares completos).</p>

        <h2>2. ANOVA de Medidas Repetidas (Pingouin)</h2>
        <p><b>A) {pretty_label(civm_col)}:</b></p>{aov_civm_html}
        <p><b>B) {pretty_label(fadiga_col)}:</b></p>{aov_fat_html}

        <h2>3. Interação e Correlação Força × Fadiga</h2>
        <img class="chart-img" src="data:image/png;base64,{b64_chart}" alt="Gráficos de interação e dispersão" />

        <div style="page-break-before: always;"></div>
        <h2>4. Síntese Interpretativa</h2>
        <p style="text-align: justify; font-size: 8.5pt;">
            A análise integrada dos dados de Contração Isométrica Voluntária Máxima (CIVM) e da taxa de fadiga espectral
            permite avaliar se a mobilização miofascial alterou a capacidade de produção de força e/ou a resistência à
            fadiga neuromuscular entre os momentos Pré e Pós, em comparação à condição Controle. Interações estatisticamente
            significativas (Condição × Momento, p &lt; 0,05) indicam associação compatível com diferenças entre condições ao longo do tempo; isoladamente, não comprovam causalidade. Na ausência de significância ou com amostra pequena, os resultados são exploratórios.
            Testes com inconsistência de nomenclatura ou alerta automático de qualidade (listados acima, quando houver)
            devem ser conferidos manualmente contra os PDFs e a auditoria de origem antes de qualquer interpretação ou divulgação científica.
        </p>
    </body>
    </html>
    """
    return html_content

# ==============================================================================
# MÓDULO 8: DADOS DEMONSTRATIVOS (PARA TESTE RÁPIDO DA INTERFACE)
# ==============================================================================

def generate_mock_master_study(n_vol: int = 5) -> tuple:
    """Gera uma Tabela Mestra sintética (CIVM + Fadiga) já no formato final, para demonstração."""
    rng = np.random.default_rng(42)
    civm_rows, fad_rows = [], []
    muscles = ['Peitoral maior direito (PMD)', 'Peitoral maior esquerdo (PME)',
               'Tríceps braquial direito (TBD)', 'Tríceps braquial esquerdo (TBE)',
               'Deltóide direito (DLD)', 'Deltóide esquerdo (DLE)']
    for sub_id in range(1, n_vol + 1):
        vol = f"{sub_id:03d}"
        for cond in ['Controle', 'Intervenção']:
            for mom in ['Pré', 'Pós']:
                boost = 2.0 if (cond == 'Intervenção' and mom == 'Pós') else 0.0
                pico_d = 45 + boost + rng.normal(0, 8)
                pico_e = 44 + boost + rng.normal(0, 8)
                civm_rows.append({
                    'ID_Voluntario': vol, 'Condicao': cond, 'Momento': mom,
                    'Arquivo_PDF': f"{vol}_CIVM{'C' if cond=='Controle' else 'I'}{'PRE' if mom=='Pré' else 'POS'}.pdf (demo)",
                    'Arquivo_txt_original': 'demo.txt', 'Nome_Consistente': True, 'Alerta_Qualidade': 'Nenhum',
                    'Pico_D': round(pico_d, 2), 'Pico_E': round(pico_e, 2),
                    'Assim_Pico_pct': round(100 * (pico_d - pico_e) / max(abs(pico_d), abs(pico_e)), 1),
                    'ForcaMedia_D': round(pico_d * 0.95, 2), 'ForcaMedia_E': round(pico_e * 0.95, 2),
                    'RFD200_D': round(60 + boost * 5 + rng.normal(0, 15), 2),
                    'RFD200_E': round(58 + boost * 5 + rng.normal(0, 15), 2),
                })
                atten = 0.15 if (cond == 'Intervenção' and mom == 'Pós') else 0.0
                for musc in muscles:
                    slope = -0.5 + atten + rng.normal(0, 0.08)
                    fad_rows.append({
                        'ID_Voluntario': vol, 'Condicao': cond, 'Momento': mom,
                        'Arquivo_PDF': f"{vol}_FAD{'C' if cond=='Controle' else 'I'}{'PRE' if mom=='Pré' else 'POS'}.pdf (demo)",
                        'Musculo': musc,
                        'MDF_Inicial_Hz': round(120 + rng.normal(0, 5), 1),
                        'MDF_Final_Hz': round(120 + slope * 28 + rng.normal(0, 5), 1),
                        'MDF_Slope_Hz_s': round(slope, 3),
                        'Queda_MDF_pct': round(slope * 28 / 120 * 100, 1),
                        'RMS_Slope_uV_s': round(0.3 - atten * 0.3 + rng.normal(0, 0.05), 3),
                        'MDF_R2': round(float(rng.uniform(0.7, 0.95)), 2),
                        'MDF_p_valor': round(float(rng.uniform(0.001, 0.045)), 4),
                    })
    return pd.DataFrame(civm_rows), pd.DataFrame(fad_rows)

# ==============================================================================
# INTERFACE DO USUÁRIO STREAMLIT
# ==============================================================================

st.markdown('<p class="main-title">🧬 Plataforma de Bioengenharia: Fusão CIVM & Fadiga sEMG</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Upload em lote com reconhecimento automático por nome de arquivo, tabela mestra consolidada, comparação Controle×Intervenção / Pré×Pós e laudo WeasyPrint</p>', unsafe_allow_html=True)

st.sidebar.header("🧭 Módulos do Sistema")
nav_choice = st.sidebar.radio(
    "Navegação:",
    [
        "1. Upload & Extração",
        "2. Tabela Mestra (Modelo Tidy)",
        "3. Comparação Controle × Intervenção",
        "4. Análise Estatística Crossover",
        "5. Emissão de Laudo PDF (WeasyPrint)",
    ]
)

st.sidebar.markdown("---")
st.sidebar.subheader("⚡ Acesso Rápido de Teste")
n_vol_demo = st.sidebar.number_input("Nº de voluntários (demo):", min_value=2, max_value=20, value=5)
if st.sidebar.button("Carregar Estudo Demonstrativo"):
    civm_demo, fad_demo = generate_mock_master_study(n_vol_demo)
    st.session_state['df_civm_master'] = civm_demo
    st.session_state['df_fatigue_master'] = fad_demo
    bump_data_version()          # a Tabela Mestra é reconstruída sob demanda (get_master)
    st.sidebar.success(f"Estudo demonstrativo carregado: {len(civm_demo)} testes de CIVM, {fad_demo['ID_Voluntario'].nunique()} voluntários.")

st.sidebar.markdown("---")
if st.sidebar.button("🗑️ Limpar todos os dados carregados"):
    clear_all_data()
    st.sidebar.success("Dados limpos.")

with st.sidebar.expander("🔧 Diagnóstico do ambiente", expanded=False):
    st.caption(f"Python {sys.version.split()[0]}")
    for _row in dependency_status():
        st.write(f"{'✅' if _row['ok'] else ('❌' if _row['required'] else '⚪')} `{_row['pacote']}` {_row['versao'] or ''}"
                 f"{'' if _row['ok'] else ('  — **obrigatório**' if _row['required'] else '  — opcional')}")
    if st.button("♻️ Limpar cache de processamento"):
        st.cache_data.clear()
        st.success("Cache limpo.")

# -----------------------------------------------------------------------------
# PÁGINA 1: UPLOAD & EXTRAÇÃO
# -----------------------------------------------------------------------------
if nav_choice == "1. Upload & Extração":
    st.header("📂 Módulo 1: Upload e Extração de Relatórios")
    st.warning("A extração é literal e limitada às tabelas/cabeçalhos reconhecidos. Valores ausentes não são imputados. Use a auditoria por página para confrontar cada PDF antes de interpretar ou publicar resultados.")

    tab_lote, tab_manual = st.tabs(["📦 Upload em Lote (automático)", "🛠️ Upload Manual (avançado)"])

    # ---------------- TAB: UPLOAD EM LOTE ----------------
    with tab_lote:
        st.markdown("""
        Arraste **todos os PDFs de uma vez** (CIVM e Fadiga, de todos os voluntários, condições e momentos).
        O sistema identifica automaticamente **ID do voluntário**, **tipo de relatório**, **Condição**
        (Controle/Intervenção) e **Momento** (Pré/Pós) pelo nome do arquivo, no padrão:
        """)
        st.code("<ID>_<CIVM|F50|FAD><C|I><PRE|POS>.pdf   —  ex.: 002 CIVMCPRE.pdf, 003 F50IPOS.pdf", language="text")

        batch_files = st.file_uploader(
            "Selecione todos os PDFs do estudo (CIVM + Fadiga, todos os voluntários):",
            type=["pdf"], accept_multiple_files=True, key="up_batch"
        )

        if st.button("🚀 Processar Lote", type="primary", disabled=not batch_files):
            civm_new, fad_new, unresolved = [], [], {}
            batch_status, seen_hashes = [], set()
            progress = st.progress(0.0, text="Processando arquivos...")
            for i, f in enumerate(batch_files):
                data = f.getvalue()
                digest = hashlib.sha256(data).hexdigest()
                if digest in seen_hashes or known_processed_hash(data, digest):
                    batch_status.append({'Arquivo': f.name, 'Estado': 'Duplicado — já importado; não foi reimportado', 'Linhas': 0})
                    progress.progress((i + 1) / len(batch_files), text=f"Ignorando duplicata: {f.name}")
                    continue
                seen_hashes.add(digest)
                meta = parse_filename_metadata(f.name)
                if meta['matched']:
                    rec, frame, audit = process_pdf_upload(data, f.name, meta, meta['tipo'])
                    if rec is not None:
                        civm_new.append(rec)
                    if not frame.empty:
                        fad_new.append(frame)
                    batch_status.append({'Arquivo': f.name, 'Estado': audit['Status_Extracao'], 'Linhas': audit['Linhas_Extraidas']})
                else:
                    unresolved[f.name] = data
                    audit = pdf_audit_record(extract_pdf_content(data), f.name, None, 'Não identificado', 'PENDENTE — metadados do nome não reconhecidos')
                    update_file_audit(audit)
                    batch_status.append({'Arquivo': f.name, 'Estado': 'Pendente — complete os metadados abaixo', 'Linhas': 0})
                progress.progress((i + 1) / len(batch_files), text=f"Processado: {f.name}")
            progress.empty()

            # Arquivos/linhas repetidos não são apagados silenciosamente: duplicatas idênticas
            # são barradas pelo SHA-256 e reportadas no resultado acima.
            if civm_new:
                st.session_state['df_civm_master'] = pd.concat(
                    [st.session_state['df_civm_master'], pd.DataFrame(civm_new)], ignore_index=True
                )
            if fad_new:
                st.session_state['df_fatigue_master'] = pd.concat(
                    [st.session_state['df_fatigue_master']] + fad_new, ignore_index=True
                )
            st.session_state['pending_raw_files'].update(unresolved)
            if civm_new or fad_new:
                bump_data_version()
            if batch_status:
                st.markdown("#### Relatório deste lote")
                st.dataframe(pd.DataFrame(batch_status), hide_index=True, **DF_STRETCH)
            st.success(f"Lote encerrado: {len(civm_new)} relatório(s) CIVM e {sum(len(d) for d in fad_new)} linha(s) de Fadiga extraídos. Consulte pendências/alertas; extração não equivale a validação científica.")

        # --- fila de pendências (nomes não reconhecidos) ---
        if st.session_state['pending_raw_files']:
            st.markdown("### 🔧 Arquivos pendentes — complete os metadados manualmente")
            pend_names = list(st.session_state['pending_raw_files'].keys())
            default_tbl = pd.DataFrame({
                'Arquivo': pend_names,
                'ID_Voluntario': ['' for _ in pend_names],
                'Tipo': ['CIVM' for _ in pend_names],
                'Condicao': ['Controle' for _ in pend_names],
                'Momento': ['Pré' for _ in pend_names],
            })
            edited = st.data_editor(
                default_tbl, hide_index=True, **EDITOR_STRETCH, key="pending_editor",
                column_config={
                    "Tipo": st.column_config.SelectboxColumn(options=["CIVM", "FADIGA"]),
                    "Condicao": st.column_config.SelectboxColumn(options=["Controle", "Intervenção"]),
                    "Momento": st.column_config.SelectboxColumn(options=["Pré", "Pós"]),
                }
            )
            if st.button("✅ Confirmar e Processar Pendências"):
                civm_new, fad_new, processed, failures = [], [], [], []
                for _, r in edited.iterrows():
                    vol = str(r['ID_Voluntario']).strip()
                    if not vol:
                        continue
                    meta = {'id': vol.zfill(3), 'condicao': r['Condicao'], 'momento': r['Momento']}
                    data = st.session_state['pending_raw_files'][r['Arquivo']]
                    rec, frame, audit = process_pdf_upload(data, r['Arquivo'], meta, r['Tipo'])
                    if rec is not None:
                        civm_new.append(rec); processed.append(r['Arquivo'])
                    elif not frame.empty:
                        fad_new.append(frame); processed.append(r['Arquivo'])
                    else:
                        failures.append({'Arquivo': r['Arquivo'], 'Estado': audit['Status_Extracao']})
                for name in processed:
                    st.session_state['pending_raw_files'].pop(name, None)
                if civm_new:
                    st.session_state['df_civm_master'] = pd.concat([st.session_state['df_civm_master'], pd.DataFrame(civm_new)], ignore_index=True)
                if fad_new:
                    st.session_state['df_fatigue_master'] = pd.concat([st.session_state['df_fatigue_master']] + fad_new, ignore_index=True)
                if failures:
                    st.warning("Alguns PDFs continuam pendentes porque não produziram registros; eles foram mantidos para revisão.")
                    st.dataframe(pd.DataFrame(failures), hide_index=True, **DF_STRETCH)
                if civm_new or fad_new:
                    bump_data_version()
                st.success(f"{len(processed)} arquivo(s) processado(s) com metadados informados.")
                st.rerun()

    # ---------------- TAB: UPLOAD MANUAL ----------------
    with tab_manual:
        st.markdown("Upload avançado: defina os metadados de uma sessão e envie os arquivos correspondentes (útil para nomes de arquivo fora do padrão).")
        col_up1, col_up2 = st.columns(2)

        with col_up1:
            st.subheader("📊 Relatórios de CIVM (PDF)")
            files_civm = st.file_uploader("Selecione um ou mais PDFs de CIVM", type=["pdf"], accept_multiple_files=True, key="up_civm_manual")
            meta_civm_id = st.text_input("ID do Voluntário:", value="001", key="mc_id")
            meta_civm_cond = st.selectbox("Condição:", ["Controle", "Intervenção"], key="mc_cond")
            meta_civm_mom = st.selectbox("Momento:", ["Pré", "Pós"], key="mc_mom")
            if st.button("Extrair e Adicionar CIVM ao Banco", type="primary", key="btn_civm_manual"):
                if files_civm and meta_civm_id.strip().isdigit():
                    meta = {'id': meta_civm_id.strip().zfill(3), 'condicao': meta_civm_cond, 'momento': meta_civm_mom}
                    new, skipped = [], 0
                    for f in files_civm:
                        data = f.getvalue()
                        if known_processed_hash(data):
                            skipped += 1; continue
                        rec, _, _ = process_pdf_upload(data, f.name, meta, 'CIVM')
                        if rec is not None:
                            new.append(rec)
                    if new:
                        st.session_state['df_civm_master'] = pd.concat([st.session_state['df_civm_master'], pd.DataFrame(new)], ignore_index=True)
                    if new:
                        bump_data_version()
                    st.success(f"{len(new)} relatório(s) CIVM importado(s); {skipped} duplicata(s) já processada(s) ignorada(s).")
                elif not meta_civm_id.strip().isdigit():
                    st.error("Informe um ID numérico de voluntário.")
                else:
                    st.warning("Selecione ao menos um arquivo.")

        with col_up2:
            st.subheader("📑 Relatórios de Fadiga (PDF)")
            files_fad = st.file_uploader("Selecione um ou mais PDFs de Fadiga", type=["pdf"], accept_multiple_files=True, key="up_fad_manual")
            meta_fad_id = st.text_input("ID do Voluntário:", value="001", key="mf_id")
            meta_fad_cond = st.selectbox("Condição:", ["Controle", "Intervenção"], key="mf_cond")
            meta_fad_mom = st.selectbox("Momento:", ["Pré", "Pós"], key="mf_mom")
            if st.button("Extrair e Adicionar Fadiga ao Banco", type="primary", key="btn_fad_manual"):
                if files_fad and meta_fad_id.strip().isdigit():
                    meta = {'id': meta_fad_id.strip().zfill(3), 'condicao': meta_fad_cond, 'momento': meta_fad_mom}
                    frames, skipped = [], 0
                    for f in files_fad:
                        data = f.getvalue()
                        if known_processed_hash(data):
                            skipped += 1; continue
                        _, frame, _ = process_pdf_upload(data, f.name, meta, 'FADIGA')
                        if not frame.empty:
                            frames.append(frame)
                    if frames:
                        st.session_state['df_fatigue_master'] = pd.concat([st.session_state['df_fatigue_master']] + frames, ignore_index=True)
                    if frames:
                        bump_data_version()
                    st.success(f"{sum(len(d) for d in frames)} linha(s) de Fadiga importadas; {skipped} duplicata(s) ignorada(s).")
                elif not meta_fad_id.strip().isdigit():
                    st.error("Informe um ID numérico de voluntário.")
                else:
                    st.warning("Selecione ao menos um arquivo.")

    # --- prévia dos bancos brutos ---
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Banco de CIVM:** ({len(st.session_state['df_civm_master'])} testes)")
        if not st.session_state['df_civm_master'].empty:
            st.dataframe(st.session_state['df_civm_master'][['ID_Voluntario', 'Condicao', 'Momento', 'Arquivo_PDF', 'Nome_Consistente', 'Alerta_Qualidade']], **DF_STRETCH, height=220)
    with c2:
        st.markdown(f"**Banco de Fadiga:** ({len(st.session_state['df_fatigue_master'])} linhas músculo)")
        if not st.session_state['df_fatigue_master'].empty:
            st.dataframe(st.session_state['df_fatigue_master'][['ID_Voluntario', 'Condicao', 'Momento', 'Musculo', 'MDF_Slope_Hz_s']], **DF_STRETCH, height=220)

    if st.session_state['audit_store']:
        with st.expander("🧾 Auditoria dos PDFs (hash, status, texto e tabelas por página)", expanded=False):
            st.dataframe(audit_summary(), **DF_STRETCH, hide_index=True)
            st.download_button(
                "Baixar auditoria integral (JSON — texto/tabelas extraídos por página)",
                data=audit_json_bytes_cached(),
                file_name="auditoria_extracao_pdf.json", mime="application/json"
            )
            if any(rec.get('Aviso_OCR') for rec in st.session_state['audit_store'].values()):
                st.warning("Um ou mais PDFs parecem ser digitalizações sem texto selecionável. Eles foram sinalizados; não foi aplicado OCR automático.")

    if st.button("🔗 (Re)construir Tabela Mestra agora"):
        get_master(force=True)
        st.success(f"Tabela mestra reconstruída: {len(st.session_state['df_merged_master'])} linhas.")

# -----------------------------------------------------------------------------
# PÁGINA 2: TABELA MESTRA
# -----------------------------------------------------------------------------
elif nav_choice == "2. Tabela Mestra (Modelo Tidy)":
    st.header("🔗 Módulo 2: Tabela Mestra Consolidada (CIVM + Fadiga)")
    st.markdown("Uma linha por **voluntário × condição × momento**. A fadiga nesta tabela é um resumo agregado; os valores extraídos por músculo permanecem no banco longo exportável, sem substituição dos valores originais.")

    if st.button("🔗 Atualizar Tabela Mestra", type="primary"):
        get_master(force=True)

    df_tidy = get_master()

    if not df_tidy.empty:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Voluntários", df_tidy['ID_Voluntario'].nunique())
        k2.metric("Testes na Tabela Mestra", len(df_tidy))
        k3.metric("Com CIVM + Fadiga", int((df_tidy.get('Fonte_Dados', pd.Series(dtype=str)) == 'CIVM + Fadiga').sum()))
        flag_mask = pd.Series(False, index=df_tidy.index)
        if 'Nome_Consistente' in df_tidy.columns:
            flag_mask |= df_tidy['Nome_Consistente'].ne(True)
        if 'Alerta_Qualidade' in df_tidy.columns:
            flag_mask |= df_tidy['Alerta_Qualidade'].fillna('Seção não localizada').ne('Nenhum')
        if 'Flag_Revisao' in df_tidy.columns:
            flag_mask |= df_tidy['Flag_Revisao'].fillna('').astype(str).ne('')
        flags = df_tidy.loc[flag_mask]
        k4.metric("Testes sinalizados (QC)", len(flags))

        if not flags.empty:
            st.markdown(f"""<div class="flag-box">⚠️ <b>{len(flags)} teste(s)</b> com nome de arquivo inconsistente com o conteúdo interno do relatório, e/ou alerta automático de controle de qualidade. Revise antes da análise estatística.</div>""", unsafe_allow_html=True)
            qc_cols = [c for c in ['ID_Voluntario', 'Condicao', 'Momento', 'Arquivo_PDF', 'Arquivo_txt_original', 'Nome_Consistente', 'Alerta_Qualidade', 'Flag_Revisao'] if c in flags.columns]
            st.dataframe(flags[qc_cols], **DF_STRETCH)

        st.markdown("### 📋 Tabela Mestra Completa")
        st.dataframe(df_tidy, **DF_STRETCH)

        st.markdown("### 📥 Exportar Tabela Mestra")
        d1, d2 = st.columns(2)
        if OPENPYXL_AVAILABLE:
            xlsx_bytes = build_master_xlsx(df_tidy, st.session_state['df_civm_master'],
                                           st.session_state['df_fatigue_master'], audit_summary())
            d1.download_button("📊 Baixar em Excel (.xlsx)", data=xlsx_bytes, file_name="tabela_mestra_civm_fadiga.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **DL_STRETCH)
        else:
            d1.caption("Excel indisponível: instale `openpyxl` (requirements.txt). Use o CSV ao lado.")
        csv_bytes = df_tidy.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
        d2.download_button("📄 Baixar em CSV (SPSS/R)", data=csv_bytes, file_name="tabela_mestra_civm_fadiga.csv",
                            mime="text/csv", **DL_STRETCH)
    else:
        st.info("Nenhum dado carregado. Vá ao Módulo 1 (Upload) ou carregue o Estudo Demonstrativo na barra lateral.")

# -----------------------------------------------------------------------------
# PÁGINA 3: COMPARAÇÃO CONTROLE × INTERVENÇÃO
# -----------------------------------------------------------------------------
elif nav_choice == "3. Comparação Controle × Intervenção":
    st.header("⚖️ Módulo 3: Comparação Crossover — Controle × Intervenção, Pré × Pós")

    df_tidy = get_master()
    if not df_tidy.empty:
        n_vol = df_tidy['ID_Voluntario'].nunique()
        if n_vol < 6:
            st.info(f"ℹ️ Amostra atual: **{n_vol} voluntário(s)**. Com n pequeno, trate os valores de p como exploratórios (baixo poder estatístico).")

        comp_df, metric_cols = cached_comparison_table(df_tidy)
        comp_view = comp_df.drop(columns=['coluna'], errors='ignore')

        st.markdown("### 📊 Tabela Comparativa (média ± DP por condição e momento)")
        st.dataframe(comp_view, **DF_STRETCH, height=420)

        if OPENPYXL_AVAILABLE:
            st.download_button("📊 Baixar Tabela Comparativa (.xlsx)", data=build_comparison_xlsx(comp_view),
                                file_name="comparacao_controle_intervencao.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.caption("Exportação em Excel indisponível: instale `openpyxl` (requirements.txt).")

        st.markdown("---")
        st.markdown("### 📈 Gráfico por variável (barras + trajetória individual)")
        comparison_chart_section(df_tidy, metric_cols)
    else:
        st.info("Tabela mestra vazia. Faça o upload no Módulo 1 ou carregue o Estudo Demonstrativo.")

# -----------------------------------------------------------------------------
# PÁGINA 4: ANÁLISE ESTATÍSTICA (ANOVA)
# -----------------------------------------------------------------------------
elif nav_choice == "4. Análise Estatística Crossover":
    st.header("📊 Módulo 4: Análise Estatística Crossover (Pingouin)")
    if not PINGOUIN_AVAILABLE:
        st.error("Pingouin não está instalado. A extração continua disponível; instale as dependências de requirements.txt para executar ANOVA.")
    st.markdown("""
    Modelo de **Medidas Repetidas (Within-Subject)**: Fator 1 = Condição (Controle vs. Intervenção),
    Fator 2 = Momento (Pré vs. Pós), unidade amostral = `ID_Voluntario`.
    """)

    df_tidy = get_master()
    if not df_tidy.empty:
        metric_cols = numeric_metric_columns(df_tidy)
        civm_cols = [c for c in metric_cols if not c.startswith('Fadiga_')]
        fadiga_cols = [c for c in metric_cols if c.startswith('Fadiga_')]

        civm_opts = sorted([pretty_label(c) for c in civm_cols]) or ["(sem dados de CIVM)"]
        fadiga_opts = sorted([pretty_label(c) for c in fadiga_cols]) or ["(sem dados de Fadiga)"]
        c1, c2 = st.columns(2)
        default_idx = civm_opts.index(pretty_label('Pico_D')) if 'Pico_D' in civm_cols else 0
        civm_label = c1.selectbox("Variável de CIVM:", civm_opts, index=default_idx)
        fadiga_label = c2.selectbox("Variável de Fadiga:", fadiga_opts)

        civm_col = {pretty_label(c): c for c in civm_cols}.get(civm_label)
        fadiga_col = {pretty_label(c): c for c in fadiga_cols}.get(fadiga_label)

        if civm_col and fadiga_col:
            stats_res = cached_rm_stats(df_tidy, civm_col, fadiga_col)

            st.markdown("---")
            st.subheader(f"1. ANOVA — {pretty_label(civm_col)}")
            if 'aov_civm' in stats_res:
                dfc = stats_res['aov_civm']
                pcol = 'p-unc' if 'p-unc' in dfc.columns else 'p_unc'
                dfc = dfc[['Source', 'SS', 'ddof1', 'ddof2', 'MS', 'F', pcol, 'ng2']].copy()
                dfc.columns = ['Fonte de Variação', 'SQ', 'GL1', 'GL2', 'QM', 'F', 'p-valor', 'Eta²g']
                st.dataframe(dfc.style.format({'SQ': '{:.2f}', 'QM': '{:.2f}', 'F': '{:.3f}', 'p-valor': '{:.4f}', 'Eta²g': '{:.3f}'}), **DF_STRETCH)
            elif 'aov_civm_err' in stats_res:
                st.warning(f"Aviso: {stats_res['aov_civm_err']}")

            st.subheader(f"2. ANOVA — {pretty_label(fadiga_col)}")
            if 'aov_fatigue' in stats_res:
                dff = stats_res['aov_fatigue']
                pcol = 'p-unc' if 'p-unc' in dff.columns else 'p_unc'
                dff = dff[['Source', 'SS', 'ddof1', 'ddof2', 'MS', 'F', pcol, 'ng2']].copy()
                dff.columns = ['Fonte de Variação', 'SQ', 'GL1', 'GL2', 'QM', 'F', 'p-valor', 'Eta²g']
                st.dataframe(dff.style.format({'SQ': '{:.4f}', 'QM': '{:.4f}', 'F': '{:.3f}', 'p-valor': '{:.4f}', 'Eta²g': '{:.3f}'}), **DF_STRETCH)
            elif 'aov_fatigue_err' in stats_res:
                st.warning(f"Aviso: {stats_res['aov_fatigue_err']}")

            if 'analysis_blocked' in stats_res:
                st.warning(stats_res['analysis_blocked'])
            else:
                st.markdown("---")
                st.subheader("3. Gráficos de Interação e Correlação Força × Fadiga")
                png = render_statistical_plots_png(stats_res['processed_data'], civm_col, fadiga_col)
                st.image(png)
                st.session_state['last_chart_b64'] = png_to_base64(png)
                st.session_state['last_civm_col'] = civm_col
                st.session_state['last_fadiga_col'] = fadiga_col
        else:
            st.info("Carregue dados de Fadiga para habilitar esta análise (ou use o Estudo Demonstrativo).")
    else:
        st.info("Tabela mestra vazia. Faça o upload no Módulo 1 ou carregue o Estudo Demonstrativo.")

# -----------------------------------------------------------------------------
# PÁGINA 5: LAUDO PDF
# -----------------------------------------------------------------------------
elif nav_choice == "5. Emissão de Laudo PDF (WeasyPrint)":
    st.header("📄 Módulo 5: Compilação de Laudo Automatizado (WeasyPrint)")

    df_tidy = get_master()
    if not df_tidy.empty:
        metric_all = numeric_metric_columns(df_tidy)
        civm_col = st.session_state.get('last_civm_col') or next((c for c in metric_all if not c.startswith('Fadiga_')), None)
        fadiga_col = st.session_state.get('last_fadiga_col') or next((c for c in metric_all if c.startswith('Fadiga_')), None)
        if civm_col not in df_tidy.columns:
            civm_col = next((c for c in metric_all if not c.startswith('Fadiga_')), None)
        if fadiga_col not in df_tidy.columns:
            fadiga_col = next((c for c in metric_all if c.startswith('Fadiga_')), None)

        stats_res, b64_chart = None, None
        if civm_col and fadiga_col:
            stats_res = cached_rm_stats(df_tidy, civm_col, fadiga_col)
            if 'analysis_blocked' not in stats_res:
                b64_chart = png_to_base64(render_statistical_plots_png(stats_res['processed_data'], civm_col, fadiga_col))
            else:
                st.warning(stats_res['analysis_blocked'])

        comp_df, _ = cached_comparison_table(df_tidy)

        if civm_col and fadiga_col and b64_chart and stats_res and 'analysis_blocked' not in stats_res:
            html_doc = build_weasyprint_html(df_tidy, comp_df, stats_res, b64_chart, civm_col, fadiga_col)
            html_key = hashlib.md5(html_doc.encode('utf-8')).hexdigest()

            # Prévia recolhida por padrão: o iframe só pesa na página quando o usuário pede.
            with st.expander("👁️ Visualizar Prévia do Laudo (HTML)", expanded=False):
                st.components.v1.html(html_doc, height=550, scrolling=True)

            st.markdown("### 📥 Download do Laudo PDF")
            weasy = _lazy('weasyprint') if WEASYPRINT_INSTALLED else None
            if weasy is not None:
                st.caption("A compilação do PDF leva alguns segundos e só é feita quando você pede (o resultado fica em cache).")
                if st.button("⚙️ Compilar laudo em PDF", type="primary"):
                    st.session_state['laudo_pdf_key'] = html_key
                if st.session_state.get('laudo_pdf_key') == html_key:
                    try:
                        pdf_bytes = compile_pdf_bytes(html_doc)
                        st.download_button("📑 Baixar Laudo Científico (PDF)", data=pdf_bytes,
                                            file_name="laudo_crossover_civm_fadiga.pdf", mime="application/pdf",
                                            type="primary", **DL_STRETCH)
                    except Exception as e:
                        st.error(f"Erro ao compilar PDF com WeasyPrint: {e}")
            else:
                st.warning("⚠️ WeasyPrint indisponível neste ambiente (biblioteca ou dependências nativas Pango/Cairo ausentes). "
                           "Baixe o HTML e use Imprimir → Salvar como PDF no navegador.")
                st.download_button("🌐 Baixar Laudo (.html — imprimível em PDF pelo navegador)",
                                    data=html_doc.encode('utf-8'), file_name="laudo_crossover_civm_fadiga.html",
                                    mime="text/html", **DL_STRETCH)
        else:
            st.info("É necessário ter pelo menos uma variável de CIVM e uma de Fadiga na Tabela Mestra. Visite o Módulo 4 para selecioná-las.")
    else:
        st.info("Tabela mestra vazia. Execute os módulos anteriores ou carregue o Estudo Demonstrativo.")
