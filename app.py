#!/usr/bin/env python3
"""
Gerenciador de PDFs — Renomeação por Planilha
Flask web app para corrigir nomes de PDFs usando planilhas de referência.
"""
from flask import Flask, render_template, request, jsonify
import os, re, unicodedata, datetime
from difflib import SequenceMatcher
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

app = Flask(__name__)
app.secret_key = 'gerenciador_pdfs_2024'

# ─── Estado global (app local, single-user) ────────────────────────────
_state = {
    'pasta_raiz': '',
    'nomes_ref': [],      # nomes normalizados da planilha de referência
    'analise': None,
}

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

def encontrar_match_ref(nome_norm, nome_stem, nomes_ref):
    """Retorna (nome_ref, tipo, razao). Tipos: correto|rename_auto|duvida|sem_match"""
    # 1. Exata
    for ref in nomes_ref:
        if normalizar(ref) == nome_norm:
            return ref, ('correto' if ref == nome_stem else 'rename_auto'), 1.0

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
    return jsonify({'ok': True, 'pasta': p, 'subpastas': len(subs), 'total_pdfs': total})

@app.route('/api/set-pasta-ref', methods=['POST'])
def set_pasta_ref():
    p = request.json.get('path', '')
    if not p or not os.path.isdir(p):
        return jsonify({'ok': False, 'erro': 'Pasta de referência inválida ou não encontrada'})
    nomes = ler_nomes_pasta_ref(p)
    if not nomes:
        return jsonify({'ok': False, 'erro': 'Nenhum nome válido (PDF) encontrado na pasta'})
    _state['nomes_ref'] = nomes
    return jsonify({
        'ok': True,
        'pasta': p,
        'total': len(nomes),
        'preview': nomes[:40],
    })

@app.route('/api/nomes-ref', methods=['GET'])
def get_nomes_ref():
    return jsonify({'nomes': _state['nomes_ref']})

@app.route('/api/analisar', methods=['POST'])
def analisar():
    if not _state['pasta_raiz']:
        return jsonify({'ok': False, 'erro': 'Selecione a pasta raiz primeiro'})
    if not _state['nomes_ref']:
        return jsonify({'ok': False, 'erro': 'Carregue a pasta de referência primeiro'})

    pdfs = buscar_todos_pdfs(_state['pasta_raiz'])
    nomes_ref = _state['nomes_ref']
    corretos = []; rename_auto = []; duvidas = []; sem_match = []

    for pdf in pdfs:
        sem_ext = os.path.splitext(pdf['arquivo'])[0]
        if eh_homonimo(sem_ext):
            corretos.append({**pdf, 'nome_ref': pdf['stem'], 'tipo': 'homonimo', 'razao': 1.0})
            continue
        ref, tipo, razao = encontrar_match_ref(pdf['nome_norm'], pdf['stem'], nomes_ref)
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
                except Exception as e:
                    resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'erro', 'msg': str(e)})
            else:
                resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'conflito',
                                   'msg': f'Conflito: "{novo_nome}" existe com conteúdo diferente'})
            continue
        try:
            os.rename(origem, destino)
            resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'ok', 'msg': f'→ {novo_nome}'})
        except Exception as e:
            resultados.append({'arquivo': arquivo, 'pasta': pasta_nome, 'status': 'erro', 'msg': str(e)})

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
        return jsonify({'ok': True, 'novo_nome': novo_nome})
    except Exception as e:
        return jsonify({'ok': False, 'erro': str(e)})

if __name__ == '__main__':
    print('=' * 50)
    print('  Gerenciador de PDFs')
    print('  Acesse: http://localhost:5000')
    print('=' * 50)
    app.run(debug=False, port=5000, use_reloader=False)
