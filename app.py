#!/usr/bin/env python3
"""
Gerenciador de PDFs — Renomeação por Planilha
Flask web app para corrigir nomes de PDFs usando planilhas de referência.
"""
from flask import Flask, render_template, request, jsonify
import os, re, unicodedata, datetime, json
from difflib import SequenceMatcher
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import llm_matcher
import confinicial
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import math

app = Flask(__name__)
app.secret_key = 'gerenciador_pdfs_2024'

# ─── Estado global (app local, single-user) ────────────────────────────
_state = {
    'pasta_raiz': '',
    'pastas_ref': [],
    'nomes_ref': [],      # nomes normalizados de todas as pastas de referência
    'analise': None,
    'needs_update': False,
    'planilhas': {}       # { '01-2008': {'caminho': '...', 'df': DataFrame, 'nome_original': '...'} }
}

HISTORICO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historico.json')

def load_historico():
    if os.path.exists(HISTORICO_FILE):
        try:
            with open(HISTORICO_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []

def save_historico(h):
    try:
        with open(HISTORICO_FILE, 'w', encoding='utf-8') as f:
            json.dump(h, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Erro salvando histórico: {e}")

def add_historico(acao, arquivo_original, novo_nome, pasta, status, msg=""):
    h = load_historico()
    h.append({
        'data': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'acao': acao,
        'arquivo_original': arquivo_original,
        'novo_nome': novo_nome,
        'pasta': pasta,
        'status': status,
        'msg': msg
    })
    # Keep last 1000
    if len(h) > 1000: h = h[-1000:]
    save_historico(h)

class PDFWatcher(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.is_directory: return
        p1 = getattr(event, 'src_path', '').lower()
        p2 = getattr(event, 'dest_path', '').lower()
        if p1.endswith('.pdf') or p2.endswith('.pdf'):
            _state['needs_update'] = True

observer = None
def setup_watchdog(path):
    global observer
    if observer:
        observer.stop()
        observer.join()
    observer = Observer()
    observer.schedule(PDFWatcher(), path, recursive=True)
    observer.start()

# ─── Constantes ────────────────────────────────────────────────────────
EXCLUIR_PADROES = [
    'ARQUIVO SEFIP', 'SEFIP', 'GRRF', 'DEPOSITADO', 'COMPROVANTE',
    'RECIBO', 'EXTRATO', 'FOLHA', 'RELATORIO', 'GUIA', 'GPS', 'GFIP',
    'PROTOCOLO', 'COMPENSACAO', 'DECLARACAO',
]
CONECTORES = {'DE', 'DA', 'DO', 'DOS', 'DAS'}
RE_VARIANTE = re.compile(r'(REC\s*\.?\s*\d+)', re.IGNORECASE)

# ─── Utilitárias ───────────────────────────────────────────────────────
def normalizar(t):
    if not isinstance(t, str): return ''
    s = unicodedata.normalize('NFKD', t)
    s = ''.join(c for c in s if not unicodedata.combining(c)).upper()
    return ' '.join(s.replace('.', '').replace('-', ' ').replace('_', ' ').split())

def similaridade(a, b):
    return SequenceMatcher(None, a, b).ratio()

def levenshtein(a, b):
    m, n = len(a), len(b)
    dp = list(range(n + 1))
    for i in range(1, m + 1):
        prev = dp[0]; dp[0] = i
        for j in range(1, n + 1):
            temp = dp[j]
            dp[j] = min(dp[j] + 1, dp[j - 1] + 1, prev + (0 if a[i - 1] == b[j - 1] else 1))
            prev = temp
    return dp[n]

def extrair_variante(nome):
    m = RE_VARIANTE.search(nome)
    if m:
        tag = re.sub(r'[\s\.]', '', m.group(1)).upper()
        return nome[:m.start()].strip(), tag
    return nome, None

def eh_pdf_valido(f):
    if not f.lower().endswith('.pdf'): return False
    n = os.path.splitext(f)[0].upper()
    return not any(p in n for p in EXCLUIR_PADROES)

def eh_homonimo(nome_sem_ext):
    return bool(re.search(r'\s\d+$', nome_sem_ext.strip()))

def verificar_abreviacao(a, b):
    ta = [t for t in a.split() if t not in CONECTORES]
    tb = [t for t in b.split() if not t.isdigit() and t not in CONECTORES]
    if abs(len(ta) - len(tb)) > 2: return False
    i = j = mc = 0
    while i < len(ta) and j < len(tb):
        pa, pb = ta[i], tb[j]
        if pa == pb: mc += 1; i += 1; j += 1; continue
        if (len(pa) <= 4 and pb.startswith(pa)) or (len(pb) <= 4 and pa.startswith(pb)):
            mc += 1; i += 1; j += 1; continue
        if len(pa) >= 4 and len(pb) >= 4 and len(ta) == len(tb):
            if levenshtein(pa, pb) <= max(1, min(len(pa), len(pb)) // 3):
                mc += 1; i += 1; j += 1; continue
            return False
        if len(pa) >= 4 and len(pb) >= 4: i += 1; j += 1
        elif len(ta) > len(tb): i += 1
        elif len(tb) > len(ta): j += 1
        else: i += 1; j += 1
    return mc >= max(len(ta), len(tb)) - 1

def ler_nomes_pasta_ref(pasta):
    nomes = []
    if not pasta or not os.path.isdir(pasta): return nomes
    for f in os.listdir(pasta):
        if not eh_pdf_valido(f): continue
        stem, _ = extrair_variante(os.path.splitext(f)[0])
        if eh_homonimo(stem): continue
        n = normalizar(stem)
        if n and n not in nomes:
            nomes.append(n)
    return nomes

def buscar_todos_pdfs(pasta_raiz):
    result = []
    if not pasta_raiz or not os.path.isdir(pasta_raiz): return result
    for d in sorted(os.listdir(pasta_raiz)):
        cp = os.path.join(pasta_raiz, d)
        if not os.path.isdir(cp): continue
        for f in sorted(os.listdir(cp)):
            if not eh_pdf_valido(f): continue
            stem, var = extrair_variante(os.path.splitext(f)[0])
            result.append({
                'pasta': d,
                'pasta_path': cp,
                'arquivo': f,
                'stem': stem,
                'nome_norm': normalizar(stem),
                'variante': var,
            })
    return result

def encontrar_match_ref(nome_norm, nome_stem, nomes_ref, ignorar_acentos=False):
    """Retorna (nome_ref, tipo, razao). Tipos: correto|rename_auto|duvida|sem_match"""
    # 1. Exata
    for ref in nomes_ref:
        if normalizar(ref) == nome_norm:
            if ignorar_acentos or ref == nome_stem:
                return ref, 'correto', 1.0
            return ref, 'rename_auto', 1.0

    # 2. Abreviação
    melhor = None; best_diff = 9999
    for ref in nomes_ref:
        rn = normalizar(ref)
        if verificar_abreviacao(nome_norm, rn):
            diff = abs(len(rn) - len(nome_norm))
            if diff < best_diff:
                best_diff = diff; melhor = ref
    if melhor:
        r = similaridade(nome_norm, normalizar(melhor))
        return melhor, ('rename_auto' if r >= 0.82 else 'duvida'), round(r, 3)

    # 3. Similaridade pura
    melhor = None; best_r = 0
    for ref in nomes_ref:
        r = similaridade(nome_norm, normalizar(ref))
        if r > best_r:
            best_r = r; melhor = ref
    if melhor and best_r >= 0.60:
        return melhor, 'duvida', round(best_r, 3)

    return None, 'sem_match', 0

def abrir_pasta():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); root.update()
        p = filedialog.askdirectory(title='Selecione uma pasta')
        root.destroy()
        return p or ''
    except Exception:
        return ''

def abrir_arquivos():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); root.update()
        files = filedialog.askopenfilenames(
            title='Selecione planilhas',
            filetypes=[('Planilhas', '*.xls *.xlsx *.ods')]
        )
        root.destroy()
        return list(files)
    except Exception:
        return []

# ─── Rotas ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/browse-folder', methods=['POST'])
def browse_folder():
    return jsonify({'path': abrir_pasta()})

@app.route('/api/set-pasta-raiz', methods=['POST'])
def set_pasta_raiz():
    p = request.json.get('path', '')
    if not p or not os.path.isdir(p):
        return jsonify({'ok': False, 'erro': 'Pasta inválida ou não encontrada'})
    _state['pasta_raiz'] = p
    subs = [d for d in os.listdir(p) if os.path.isdir(os.path.join(p, d))]
    total = sum(len([f for f in os.listdir(os.path.join(p, d)) if eh_pdf_valido(f)]) for d in subs)
    
    setup_watchdog(p)
    _state['needs_update'] = False
    
    return jsonify({'ok': True, 'pasta': p, 'subpastas': len(subs), 'total_pdfs': total})

@app.route('/api/add-pasta-ref', methods=['POST'])
def add_pasta_ref():
    p = request.json.get('path', '')
    if not p or not os.path.isdir(p):
        return jsonify({'ok': False, 'erro': 'Pasta de referência inválida ou não encontrada'})
    if p in _state['pastas_ref']:
        return jsonify({'ok': False, 'erro': 'Esta pasta já foi adicionada'})
        
    nomes = ler_nomes_pasta_ref(p)
    if not nomes:
        return jsonify({'ok': False, 'erro': 'Nenhum nome válido (PDF) encontrado na pasta'})
    
    _state['pastas_ref'].append(p)
    _state['nomes_ref'].extend(n for n in nomes if n not in _state['nomes_ref'])
    
    return jsonify({
        'ok': True,
        'pastas': _state['pastas_ref'],
        'total': len(_state['nomes_ref']),
    })

@app.route('/api/clear-pastas-ref', methods=['POST'])
def clear_pastas_ref():
    _state['pastas_ref'] = []
    _state['nomes_ref'] = []
    return jsonify({'ok': True})

@app.route('/api/nomes-ref', methods=['GET'])
def get_nomes_ref():
    return jsonify({'nomes': _state['nomes_ref']})

@app.route('/api/check-update', methods=['GET'])
def check_update():
    if _state['needs_update']:
        _state['needs_update'] = False
        return jsonify({'update': True})
    return jsonify({'update': False})

@app.route('/api/historico', methods=['GET'])
def get_historico():
    h = load_historico()
    h.reverse()
    return jsonify({'historico': h})

@app.route('/api/shutdown', methods=['POST'])
def shutdown():
    global observer
    if observer:
        try:
            observer.stop()
        except: pass
    
    def kill_server():
        import time
        time.sleep(0.5)
        os._exit(0)
    
    import threading
    threading.Thread(target=kill_server).start()
    return jsonify({'ok': True})

@app.route('/api/analisar', methods=['POST'])
def analisar():
    if not _state['pasta_raiz']:
        return jsonify({'ok': False, 'erro': 'Selecione a pasta raiz primeiro'})
    if not _state['nomes_ref']:
        return jsonify({'ok': False, 'erro': 'Carregue a pasta de referência primeiro'})

    ignorar_acentos = request.json.get('ignorar_acentos', False)
    pdfs = buscar_todos_pdfs(_state['pasta_raiz'])
    nomes_ref = _state['nomes_ref']
    corretos = []; rename_auto = []; duvidas = []; sem_match = []

    for pdf in pdfs:
        sem_ext = os.path.splitext(pdf['arquivo'])[0]
        if eh_homonimo(sem_ext):
            corretos.append({**pdf, 'nome_ref': pdf['stem'], 'tipo': 'homonimo', 'razao': 1.0})
            continue
        ref, tipo, razao = encontrar_match_ref(pdf['nome_norm'], pdf['stem'], nomes_ref, ignorar_acentos)
        entry = {**pdf, 'nome_ref': ref, 'tipo': tipo, 'razao': razao}
        if tipo == 'correto':
            corretos.append(entry)
        elif tipo == 'rename_auto':
            rename_auto.append(entry)
        elif tipo == 'duvida':
            duvidas.append(entry)
        else:
            sem_match.append(entry)

    _state['analise'] = {
        'corretos': corretos,
        'rename_auto': rename_auto,
        'duvidas': duvidas,
        'sem_match': sem_match,
    }

    return jsonify({
        'ok': True,
        'corretos': len(corretos),
        'rename_auto': len(rename_auto),
        'duvidas': len(duvidas),
        'sem_match': len(sem_match),
        'rename_auto_data': rename_auto,
        'duvidas_data': duvidas,
        'sem_match_data': sem_match,
    })

@app.route('/api/executar', methods=['POST'])
def executar():
    items = request.json.get('items', [])
    resultados = []
    for item in items:
        pasta_path = item['pasta_path']
        arquivo = item['arquivo']
        nome_ref = item.get('nome_ref', '')
        if not nome_ref:
            resultados.append({'arquivo': arquivo, 'status': 'erro', 'msg': 'Nome de referência não informado'})
            continue
        origem = os.path.join(pasta_path, arquivo)
        if not os.path.exists(origem):
            resultados.append({'arquivo': arquivo, 'status': 'erro', 'msg': 'Arquivo não encontrado'})
            continue
        stem, var = extrair_variante(os.path.splitext(arquivo)[0])
        novo_nome = nome_ref + (' ' + var if var else '') + '.pdf'
        destino = os.path.join(pasta_path, novo_nome)
        pasta_nome = os.path.basename(pasta_path)
        if arquivo == novo_nome:
            resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'igual', 'msg': 'Já estava correto'})
            continue
        if os.path.exists(destino):
            if os.path.getsize(origem) == os.path.getsize(destino):
                try:
                    os.remove(origem)
                    resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'excluido',
                                       'msg': f'Duplicado removido — "{novo_nome}" já existia'})
                    add_historico('Lote (Excluído)', arquivo, '-', pasta_nome, 'Sucesso', 'Duplicado removido')
                except Exception as e:
                    resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'erro', 'msg': str(e)})
                    add_historico('Lote (Excluído)', arquivo, '-', pasta_nome, 'Erro', str(e))
            else:
                resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'conflito',
                                   'msg': f'Conflito: "{novo_nome}" existe com conteúdo diferente'})
            continue
        try:
            os.rename(origem, destino)
            resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'ok', 'msg': f'→ {novo_nome}'})
            add_historico('Lote', arquivo, novo_nome, pasta_nome, 'Sucesso')
        except Exception as e:
            resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'erro', 'msg': str(e)})
            add_historico('Lote', arquivo, novo_nome, pasta_nome, 'Erro', str(e))

    return jsonify({'ok': True, 'resultados': resultados})

@app.route('/api/buscar', methods=['POST'])
def buscar():
    termo = normalizar(request.json.get('termo', ''))
    if not termo or len(termo) < 3:
        return jsonify({'ok': False, 'erro': 'Digite pelo menos 3 caracteres'})
    if not _state['pasta_raiz']:
        return jsonify({'ok': False, 'erro': 'Selecione a pasta raiz primeiro (aba Configuração)'})
    pdfs = buscar_todos_pdfs(_state['pasta_raiz'])
    resultados = []
    for pdf in pdfs:
        n = pdf['nome_norm']
        r = similaridade(termo, n)
        if r >= 0.5 or termo in n:
            resultados.append({
                'pasta': pdf['pasta'],
                'pasta_path': pdf['pasta_path'],
                'arquivo': pdf['arquivo'],
                'stem': pdf['stem'],
                'razao': round(r, 3),
            })
    resultados.sort(key=lambda x: -x['razao'])
    return jsonify({'ok': True, 'resultados': resultados[:80]})

@app.route('/api/renomear', methods=['POST'])
def renomear():
    data = request.json
    pasta_path = data.get('pasta_path', '')
    arquivo = data.get('arquivo', '')
    novo_nome = data.get('novo_nome', '').strip()
    if not novo_nome:
        return jsonify({'ok': False, 'erro': 'Nome não informado'})
    if not novo_nome.lower().endswith('.pdf'):
        novo_nome += '.pdf'
    origem = os.path.join(pasta_path, arquivo)
    destino = os.path.join(pasta_path, novo_nome)
    if not os.path.exists(origem):
        return jsonify({'ok': False, 'erro': 'Arquivo não encontrado'})
    if os.path.normcase(origem) == os.path.normcase(destino):
        return jsonify({'ok': False, 'erro': 'O nome é igual ao atual'})
    if os.path.exists(destino):
        return jsonify({'ok': False, 'erro': f'"{novo_nome}" já existe nessa pasta'})
    try:
        os.rename(origem, destino)
        add_historico('Busca (Individual)', arquivo, novo_nome, os.path.basename(pasta_path), 'Sucesso')
        return jsonify({'ok': True, 'novo_nome': novo_nome})
    except Exception as e:
        add_historico('Busca (Individual)', arquivo, novo_nome, os.path.basename(pasta_path), 'Erro', str(e))
        return jsonify({'ok': False, 'erro': str(e)})

# ─── Funções Auxiliares de Conciliação ──────────────────────────────────
def converter_xls_para_xlsx(caminho_xls):
    caminho_xlsx = os.path.splitext(caminho_xls)[0] + '.xlsx'
    if os.path.exists(caminho_xlsx):
        os.remove(caminho_xlsx)

    try:
        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        xl = win32com.client.Dispatch('Excel.Application')
        try:
            xl.DisplayAlerts = False
        except:
            pass
        wb = xl.Workbooks.Open(caminho_xls)
        wb.SaveAs(caminho_xlsx, FileFormat=51)  # 51 = xlOpenXMLWorkbook
        wb.Close()
        xl.Quit()
        pythoncom.CoUninitialize()
        return caminho_xlsx
    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        try:
            import xlrd
            rb = xlrd.open_workbook(caminho_xls)
            nome_aba = 'FGTS EM ATRASO - PROCESSOS'
            if nome_aba not in rb.sheet_names():
                nome_aba = rb.sheet_names()[0]
            df_raw = pd.read_excel(caminho_xls, sheet_name=nome_aba, header=None)
            df_raw.to_excel(caminho_xlsx, sheet_name='FGTS EM ATRASO - PROCESSOS', index=False, header=False)
            return caminho_xlsx
        except:
            return None

def normalizar_pis(valor):
    if valor is None:
        return ''
    if isinstance(valor, float):
        if math.isnan(valor) or valor == 0:
            return ''
        return str(int(valor)).zfill(11)
    texto = str(valor).strip()
    if not texto or texto == '0':
        return ''
    digitos = re.sub(r'\D', '', texto)
    return digitos.zfill(11)

# ─── Conciliação de Planilhas ──────────────────────────────────────────

@app.route('/api/carregar-planilhas', methods=['POST'])
def carregar_planilhas():
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    caminhos = filedialog.askopenfilenames(
        title="Selecione as Planilhas FGTS",
        filetypes=[("Planilhas", "*.xlsx *.xls *.ods"), ("Todos", "*.*")]
    )
    root.destroy()
    
    if not caminhos:
        return jsonify({'ok': False, 'erro': 'Nenhum arquivo selecionado'})
        
    sucessos = 0
    erros = []
    
    for c in caminhos:
        mes, ano = confinicial.extrair_competencia_excel(os.path.basename(c))
        if not mes or not ano:
            erros.append(f"Não foi possível extrair a competência do arquivo: {os.path.basename(c)}")
            continue
            
        comp_str = f"{mes}-{ano}"
        df = confinicial.ler_planilha_fgts(c)
        if df is None or df.empty:
            erros.append(f"Planilha sem dados válidos ou falha na leitura: {os.path.basename(c)}")
            continue
            
        _state['planilhas'][comp_str] = {
            'caminho': c,
            'nome_original': os.path.basename(c),
            'df': df
        }
        sucessos += 1
        
    return jsonify({'ok': True, 'sucessos': sucessos, 'erros': erros})

@app.route('/api/status-conciliacao', methods=['GET'])
def status_conciliacao():
    if not _state['planilhas']:
        return jsonify({'ok': True, 'cards': [], 'conflitos': []})
        
    # Agrupa por competência para os cards
    cards = []
    pis_dict = {} # PIS -> dict de nomes (nome -> lista de competencias)
    
    for comp, p_data in _state['planilhas'].items():
        df = p_data['df']
        cards.append({
            'competencia': comp,
            'arquivo': p_data['nome_original'],
            'total': len(df)
        })
        
        for idx, row in df.iterrows():
            pis = str(row['PIS']).strip()
            nome = str(row['NOMES']).strip()
            if pis and nome and str(pis).lower() not in ('nan', 'none', ''):
                if pis not in pis_dict:
                    pis_dict[pis] = {}
                # normalizar espacos para comparação
                n_clean = " ".join(nome.split())
                if n_clean not in pis_dict[pis]:
                    pis_dict[pis][n_clean] = []
                if comp not in pis_dict[pis][n_clean]:
                    pis_dict[pis][n_clean].append(comp)
                    
    # Encontra conflitos
    conflitos = []
    # Usar _state['nomes_ref'] (normalizado) para sugerir melhor nome
    nomes_ref_norm = _state['nomes_ref']
    
    for pis, nomes_map in pis_dict.items():
        if len(nomes_map) > 1:
            opcoes = list(nomes_map.keys())
            # Tenta achar um que bate com a pasta de referencia
            sugerido = None
            for op in opcoes:
                if normalizar(op) in nomes_ref_norm:
                    sugerido = op
                    break
            if not sugerido:
                sugerido = sorted(opcoes, key=len, reverse=True)[0]
                
            opcoes_detalhe = []
            for op in opcoes:
                opcoes_detalhe.append({
                    'nome': op,
                    'competencias': nomes_map[op]
                })
                
            conflitos.append({
                'pis': pis,
                'opcoes': opcoes_detalhe,
                'sugerido': sugerido
            })
            
    # Ordena cards
    cards.sort(key=lambda x: x['competencia'])
    return jsonify({'ok': True, 'cards': cards, 'conflitos': conflitos})

@app.route('/api/corrigir-pis', methods=['POST'])
def corrigir_pis():
    data = request.json
    pis = str(data.get('pis')).strip()
    novo_nome = data.get('novo_nome', '').strip()
    
    if not pis or not novo_nome:
        return jsonify({'ok': False, 'erro': 'Dados incompletos'})
        
    afetados = 0
    for comp, p_data in _state['planilhas'].items():
        df = p_data['df']
        # Usamos apply/lambda pra converter column pra str sem alterar tipo original
        mascara = df['PIS'].astype(str).str.strip() == pis
        if mascara.any():
            # Acha os que estao diferentes
            diferentes = mascara & (df['NOMES'] != novo_nome)
            qtd = diferentes.sum()
            if qtd > 0:
                df.loc[diferentes, 'NOMES'] = novo_nome
                afetados += qtd
                
    return jsonify({'ok': True, 'afetados': int(afetados)})

@app.route('/api/limpar-conciliacao', methods=['POST'])
def limpar_conciliacao():
    _state['planilhas'] = {}
    return jsonify({'ok': True})

@app.route('/api/analisar-conciliacao', methods=['POST'])
def analisar_conciliacao():
    if not _state['pasta_raiz']:
        return jsonify({'ok': False, 'erro': 'Selecione a Pasta Raiz de PDFs na aba de Configuração primeiro.'})
    if not _state['planilhas']:
        return jsonify({'ok': False, 'erro': 'Nenhuma planilha carregada.'})

    mapa_pastas = {}
    for d in os.listdir(_state['pasta_raiz']):
        caminho_dir = os.path.join(_state['pasta_raiz'], d)
        if os.path.isdir(caminho_dir):
            mes, ano = confinicial.extrair_competencia_pasta(d)
            if mes and ano:
                mapa_pastas[f"{mes}-{ano}"] = caminho_dir

    resultados_audit = {}

    def _status_base(s):
        for prefixo in [
            "NÃO ENCONTRADO NO .PDF",
            "POSSÍVEL ERRO NOMINAL",
            "ENCONTRADO COMO 115",
            "ENCONTRADO COM ABREVIAÇÃO",
            "ENCONTRADO",
            "ENCONTRADO VIA LLM"
        ]:
            if str(s).startswith(prefixo):
                return prefixo
        if str(s).startswith("PDF NA PASTA"):
            return "PDF NA PASTA, MAS NÃO NA PLANILHA"
        return str(s)

    for comp, p_data in _state['planilhas'].items():
        df = p_data['df'].copy()
        pasta_pdfs = mapa_pastas.get(comp)
        
        if not pasta_pdfs:
            resultados_audit[comp] = {
                'erro_pasta': True,
                'estatisticas': {},
                'erros_nominais': [],
                'nao_encontrados': [],
                'pdfs_extras': [],
                'caminho_planilha': p_data['caminho']
            }
            continue
            
        df_verificado = confinicial.verificar_pdfs(df, pasta_pdfs)
        df_verificado['Status Base'] = df_verificado['Status PDF'].apply(_status_base)
        
        totais = df_verificado['Status Base'].value_counts().to_dict()
        
        estatisticas = {
            'encontrados': totais.get('ENCONTRADO', 0) + totais.get('ENCONTRADO COMO 115', 0) + totais.get('ENCONTRADO COM ABREVIAÇÃO', 0) + totais.get('ENCONTRADO VIA LLM', 0),
            'erros_nominais': totais.get('POSSÍVEL ERRO NOMINAL', 0),
            'nao_encontrados': totais.get('NÃO ENCONTRADO NO .PDF', 0),
            'pdfs_extras': totais.get('PDF NA PASTA, MAS NÃO NA PLANILHA', 0),
            'llm': totais.get('ENCONTRADO VIA LLM', 0)
        }
        
        erros_nominais = []
        nao_encontrados = []
        pdfs_extras = []
        
        for idx, row in df_verificado.iterrows():
            status = str(row.get('Status Base', ''))
            pis = normalizar_pis(row.get('PIS', ''))
            nome_planilha = str(row.get('NOMES', ''))
            nome_pdf = str(row.get('Nome do Arquivo Encontrado', ''))
            proc = str(row.get('PROC.', ''))
            
            if status == 'POSSÍVEL ERRO NOMINAL':
                erros_nominais.append({'pis': pis, 'nome_planilha': nome_planilha, 'nome_pdf': nome_pdf})
            elif status == 'NÃO ENCONTRADO NO .PDF':
                nao_encontrados.append({'pis': pis, 'nome_planilha': nome_planilha, 'proc': proc})
            elif status == 'PDF NA PASTA, MAS NÃO NA PLANILHA':
                pdfs_extras.append({'nome_pdf': nome_pdf})

        resultados_audit[comp] = {
            'erro_pasta': False,
            'estatisticas': estatisticas,
            'erros_nominais': erros_nominais,
            'nao_encontrados': nao_encontrados,
            'pdfs_extras': pdfs_extras,
            'caminho_planilha': p_data['caminho']
        }

    return jsonify({'ok': True, 'resultados': resultados_audit})

@app.route('/api/corrigir-xls-fisico', methods=['POST'])
def corrigir_xls_fisico():
    data = request.json
    caminho_planilha = data.get('caminho_planilha')
    pis_alvo = normalizar_pis(data.get('pis'))
    novo_nome = data.get('novo_nome')
    comp = data.get('competencia')
    
    if not os.path.exists(caminho_planilha):
        return jsonify({'ok': False, 'erro': 'Arquivo xls/xlsx não encontrado fisicamente.'})
        
    ext = os.path.splitext(caminho_planilha)[1].lower()
    caminho_xlsx = caminho_planilha
    
    if ext == '.xls':
        caminho_xlsx = converter_xls_para_xlsx(caminho_planilha)
        if not caminho_xlsx:
            return jsonify({'ok': False, 'erro': 'Falha ao converter .xls para .xlsx. O Excel precisa estar instalado.'})
            
    try:
        wb = load_workbook(caminho_xlsx)
        try:
            ws = wb['FGTS EM ATRASO - PROCESSOS']
        except:
            ws = wb.active
            
        alterados = 0
        # A varredura de PIS na coluna C (3) e Nome na coluna D (4)
        for row in range(1, ws.max_row + 1):
            val_pis = ws.cell(row=row, column=3).value
            if normalizar_pis(val_pis) == pis_alvo:
                ws.cell(row=row, column=4).value = novo_nome
                alterados += 1
                
        wb.save(caminho_xlsx)
        
        # Atualiza em memória o df
        if comp in _state['planilhas']:
            # Se era .xls e virou .xlsx, atualiza o caminho na memoria
            _state['planilhas'][comp]['caminho'] = caminho_xlsx
            df = _state['planilhas'][comp]['df']
            
            # Recarrega a planilha recem salva para memoria
            novo_df = confinicial.ler_planilha_fgts(caminho_xlsx)
            if novo_df is not None:
                _state['planilhas'][comp]['df'] = novo_df
                
        return jsonify({'ok': True, 'alterados': alterados, 'novo_caminho': caminho_xlsx})
    except Exception as e:
        return jsonify({'ok': False, 'erro': str(e)})

@app.route('/api/gerar-conciliacao', methods=['POST'])
def gerar_conciliacao():
    if not _state['pasta_raiz']:
        return jsonify({'ok': False, 'erro': 'Selecione a Pasta Raiz de PDFs na aba de Configuração primeiro.'})
    if not _state['planilhas']:
        return jsonify({'ok': False, 'erro': 'Nenhuma planilha carregada.'})
        
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    pasta_destino = filedialog.askdirectory(title="Selecione a PASTA DE DESTINO para salvar os relatórios e organizar")
    root.destroy()
    
    if not pasta_destino:
        return jsonify({'ok': False, 'erro': 'Geração cancelada (nenhuma pasta destino selecionada).'})
        
    # Mapear as subpastas da raiz por competencia
    mapa_pastas = {}
    for d in os.listdir(_state['pasta_raiz']):
        caminho_dir = os.path.join(_state['pasta_raiz'], d)
        if os.path.isdir(caminho_dir):
            mes, ano = confinicial.extrair_competencia_pasta(d)
            if mes and ano:
                mapa_pastas[f"{mes}-{ano}"] = caminho_dir

    resultados = []
    
    for comp, p_data in _state['planilhas'].items():
        df = p_data['df'].copy()
        nome_orig = p_data['nome_original']
        
        pasta_pdfs = mapa_pastas.get(comp)
        if not pasta_pdfs:
            # Planilha sem pasta correspondente
            df['Status PDF'] = "PASTA NÃO ENCONTRADA"
            df['Nome do Arquivo Encontrado'] = ""
            df_final = df
            total_pdfs = 0
        else:
            total_pdfs = len([f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')])
            # Executa a verificação principal
            df_final = confinicial.verificar_pdfs(df, pasta_pdfs)
            # Organiza os arquivos
            confinicial.organizar_pdfs_por_resultado(df_final, pasta_pdfs, pasta_destino)
            
        # Salva o arquivo de relatorio
        nome_saida = f"Conciliacao_{comp}.xlsx"
        caminho_saida = os.path.join(pasta_destino, nome_saida)
        try:
            confinicial.salvar_com_resumo(df_final, caminho_saida, total_pdfs=total_pdfs)
            resultados.append(f"Gerado: {nome_saida}")
        except Exception as e:
            resultados.append(f"Erro ao salvar {nome_saida}: {str(e)}")
            
    # Ao final da geracao, limpa o estado
    _state['planilhas'] = {}
    return jsonify({'ok': True, 'resultados': resultados, 'destino': pasta_destino})

if __name__ == '__main__':
    print('=' * 50)
    print('  Gerenciador de PDFs')
    print('  Acesse: http://localhost:5000')
    print('=' * 50)
    app.run(debug=False, port=5000, use_reloader=False)
