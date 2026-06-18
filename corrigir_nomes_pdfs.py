#!/usr/bin/env python3
"""
Corrige nomes de PDFs em multiplas pastas usando UMA OU MAIS pastas de referencia.
Reutiliza a logica de matching do confinicial.py (exata, abreviacao, similaridade).
A cada correcao manual que voce faz, inclua a pasta como referencia extra e o
script fica cada vez mais preciso (autocorrecao progressiva).
"""
import os, re, shutil, unicodedata, datetime, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

EXT = ('.pdf', '.tif', '.tiff')
EXCLUIR_PADROES = [
    'ARQUIVO SEFIP', 'SEFIP', 'GRRF', 'DEPOSITADO',
    'COMPROVANTE', 'RECIBO', 'EXTRATO', 'FOLHA',
    'RELATORIO', 'GUIA', 'GPS', 'GFIP',
    'PROTOCOLO', 'COMPENSACAO', 'DECLARACAO',
]
import tkinter as tk
from tkinter import filedialog
from difflib import SequenceMatcher
try:
    import llm_matcher
except ImportError:
    llm_matcher = None

USAR_LLM = False

# ============================================================
# FUNCOES DE MATCHING (iguais ao confinicial.py)
# ============================================================

def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    nfkd = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    texto_upper = texto_sem_acento.upper()
    texto_sem_ponto = texto_upper.replace('.', '').replace('(', '').replace(')', '')
    texto_sem_hifen = texto_sem_ponto.replace('-', ' ').replace('_', ' ')
    return " ".join(texto_sem_hifen.split())

def eh_pdf_valido(nome_arquivo):
    if not nome_arquivo.lower().endswith(EXT):
        return False
    nome_sem_ext = os.path.splitext(nome_arquivo)[0].upper()
    return not any(p in nome_sem_ext for p in EXCLUIR_PADROES)

def calcular_similaridade(a, b):
    return SequenceMatcher(None, a, b).ratio()

def levenshtein(a, b):
    m, n = len(a), len(b)
    dp = list(range(n + 1))
    for i in range(1, m + 1):
        prev = dp[0]
        dp[0] = i
        for j in range(1, n + 1):
            temp = dp[j]
            dp[j] = min(dp[j] + 1, dp[j - 1] + 1, prev + (0 if a[i - 1] == b[j - 1] else 1))
            prev = temp
    return dp[n]

def verificar_abreviacao(nome_planilha, nome_pdf):
    conectores = ['DE', 'DA', 'DO', 'DOS', 'DAS']
    tokens_p = [t for t in nome_planilha.split() if t not in conectores]
    tokens_f = [t for t in nome_pdf.split() if not t.isdigit() and t not in conectores]
    if abs(len(tokens_p) - len(tokens_f)) > 2:
        return False
    idx_p = idx_f = match_count = 0
    while idx_p < len(tokens_p) and idx_f < len(tokens_f):
        t_p, t_f = tokens_p[idx_p], tokens_f[idx_f]
        if t_p == t_f:
            match_count += 1; idx_p += 1; idx_f += 1; continue
        if (len(t_p) <= 4 and t_f.startswith(t_p)) or (len(t_f) <= 4 and t_p.startswith(t_f)):
            match_count += 1; idx_p += 1; idx_f += 1; continue
        if len(t_p) >= 4 and len(t_f) >= 4 and len(tokens_p) == len(tokens_f):
            if levenshtein(t_p, t_f) <= max(1, min(len(t_p), len(t_f)) // 3):
                match_count += 1; idx_p += 1; idx_f += 1; continue
            return False
        if len(t_p) >= 4 and len(t_f) >= 4:
            idx_p += 1; idx_f += 1
        elif len(tokens_p) > len(tokens_f): idx_p += 1
        elif len(tokens_f) > len(tokens_p): idx_f += 1
        else: idx_p += 1; idx_f += 1
    return match_count >= max(len(tokens_p), len(tokens_f)) - 1

RE_VARIANTE = re.compile(r'(REC\s*\.?\s*\d+)', re.IGNORECASE)

def extrair_variante(nome_sem_ext):
    """Detecta sufixo como 'REC. 115', 'REC 115' no nome.
    Retorna (nome_stem, variante) onde variante e 'REC115' (sem pontuacao) ou None."""
    m = RE_VARIANTE.search(nome_sem_ext)
    if m:
        raw = m.group(1)
        tag = re.sub(r'[\s\.]', '', raw).upper()
        stem = nome_sem_ext[:m.start()].strip()
        return stem, tag
    return nome_sem_ext, None

def encontrar_melhor_match(nome_busca, arquivos_ref):
    """Procura o melhor match para nome_busca na lista de arquivos_ref.
    Se o nome busca tiver variante (ex: REC 115), prioriza referencias com a mesma.
    Retorna (nome_corrigido, tipo_match) ou (None, '')."""
    stem_busca, var_busca = extrair_variante(os.path.splitext(nome_busca)[0])
    nome_norm = normalizar_texto(stem_busca)

    # Se tem variante, filtra refs com a mesma; se nao tem, filtra refs sem variante
    if var_busca:
        refs_mesma_var = [a for a in arquivos_ref if extrair_variante(os.path.splitext(a)[0])[1] == var_busca]
        refs_outras = [a for a in arquivos_ref if a not in refs_mesma_var]
    else:
        refs_mesma_var = [a for a in arquivos_ref if not extrair_variante(os.path.splitext(a)[0])[1]]
        refs_outras = [a for a in arquivos_ref if a not in refs_mesma_var]

    def _buscar(grupo):
        if not grupo:
            return None, ''
        # 1. Exata (substring, preferindo tamanho mais proximo)
        candidatos = []
        for arq in grupo:
            stem_ref, _ = extrair_variante(os.path.splitext(arq)[0])
            arq_norm = normalizar_texto(stem_ref)
            if nome_norm == arq_norm:
                candidatos.append((arq, arq_norm))
        if candidatos:
            return candidatos[0][0], 'exata'

        # 2. Abreviacao
        melhor_arq = None
        melhor_tam = 999999
        for arq in grupo:
            stem_ref, _ = extrair_variante(os.path.splitext(arq)[0])
            arq_norm = normalizar_texto(stem_ref)
            if verificar_abreviacao(nome_norm, arq_norm):
                diff = abs(len(arq_norm) - len(nome_norm))
                if diff < melhor_tam:
                    melhor_tam = diff
                    melhor_arq = arq
        if melhor_arq:
            if USAR_LLM and llm_matcher:
                stem_ref2, _ = extrair_variante(os.path.splitext(melhor_arq)[0])
                ref_norm = normalizar_texto(stem_ref2)
                resp = llm_matcher.verificar_com_llm(nome_norm, ref_norm)
                if resp is True:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> SIM (abreviacao)')
                    return melhor_arq, 'abreviacao'
                elif resp == "DUVIDA":
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> DUVIDA (abreviacao)')
                    return melhor_arq, 'DUVIDA'
                elif resp is False:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> NAO')
                    melhor_arq = None
            if melhor_arq:
                return melhor_arq, 'abreviacao'

        # 3. Similaridade (exige que passe pela abreviacao primeiro)
        melhor_arq = None
        melhor_ratio = 0
        for arq in grupo:
            stem_ref, _ = extrair_variante(os.path.splitext(arq)[0])
            arq_norm = normalizar_texto(stem_ref)
            if not verificar_abreviacao(nome_norm, arq_norm):
                continue
            ratio = calcular_similaridade(nome_norm, arq_norm)
            if ratio >= 0.80 and ratio > melhor_ratio:
                melhor_ratio = ratio
                melhor_arq = arq
        if melhor_arq:
            if USAR_LLM and llm_matcher:
                stem_ref2, _ = extrair_variante(os.path.splitext(melhor_arq)[0])
                ref_norm = normalizar_texto(stem_ref2)
                resp = llm_matcher.verificar_com_llm(nome_norm, ref_norm)
                if resp is True:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> SIM')
                    return melhor_arq, f'similaridade ({melhor_ratio:.0%})'
                elif resp == "DUVIDA":
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> DUVIDA')
                    return melhor_arq, 'DUVIDA'
                elif resp is False:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> NAO')
                    melhor_arq = None
            if melhor_arq:
                return melhor_arq, f'similaridade ({melhor_ratio:.0%})'

        # 4. LLM fallback
        if USAR_LLM and llm_matcher:
            for arq in grupo:
                stem_ref2, _ = extrair_variante(os.path.splitext(arq)[0])
                ref_norm = normalizar_texto(stem_ref2)
                resp = llm_matcher.verificar_com_llm(nome_norm, ref_norm)
                if resp is True:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> SIM (fallback)')
                    return arq, 'via LLM'
                elif resp == "DUVIDA":
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> DUVIDA (fallback)')
                    return arq, 'DUVIDA'
                elif resp is False:
                    print(f'    [LLM] {nome_norm} x {ref_norm} -> NAO')

        return None, ''

    match, tipo = _buscar(refs_mesma_var)
    if match:
        return match, tipo

    match, tipo = _buscar(refs_outras)
    if match:
        return match, tipo + ' (var diff)'

    return None, ''


# ============================================================
# DIALOGOS TKINTER
# ============================================================

def dialogo_pastas(titulo_primeira, titulo_proxima):
    """Seleciona multiplas pastas, uma por vez, ate o usuario cancelar."""
    pastas = []
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    primeiro = True
    while True:
        root.update()
        titulo = titulo_primeira if primeiro else titulo_proxima
        pasta = filedialog.askdirectory(title=titulo)
        if not pasta:
            break
        if pasta not in pastas:
            pastas.append(pasta)
        primeiro = False
    root.destroy()
    return pastas


# ============================================================
# MAIN
# ============================================================

def main():
    print('=' * 60)
    print('  CORRECAO DE NOMES DE PDFs - MULTIPLAS PASTAS')
    print('=' * 60)
    print()
    print('Voce pode selecionar UMA OU MAIS pastas de REFERENCIA.')
    print('O script junta todos os PDFs delas para usar como base.')
    print('Assim, a cada correcao manual, inclua a pasta como referencia')
    print('e o script fica cada vez mais preciso (autocorrecao).')
    print()

    # Modo
    print('Modos:')
    print('  1 — Executar (renomeia os arquivos)')
    print('  2 — Executar com LLM (Gemini)')
    while True:
        modo = input('\nOpcao (1 ou 2): ').strip()
        if modo in ('1', '2'):
            break

    if modo == '2':
        global USAR_LLM
        USAR_LLM = True
        print('\nModo LLM ativado.')
        if llm_matcher:
            resp = llm_matcher.verificar_com_llm('TESTE', 'TESTE')
            if resp is None:
                print('  ATENCAO: LLM nao disponivel. Configure a chave em config_llm.py')
                input('\nPressione Enter para continuar mesmo assim...')
            else:
                print('  LLM OK')
        else:
            print('  ATENCAO: llm_matcher.py nao encontrado na pasta.')
            input('\nPressione Enter para continuar mesmo assim...')

    # Seleciona pastas referencia
    print('\n1. Selecione a(s) pasta(s) de REFERENCIA (com nomes corretos)')
    print('   (selecione uma, depois outra, e assim por diante)')
    print('   (quanto terminar, clique em Cancelar)')
    pastas_ref = dialogo_pastas(
        'SELECIONE pasta REFERENCIA (nomes corretos)',
        'SELECIONE OUTRA pasta REFERENCIA (ou Cancelar)'
    )
    if not pastas_ref:
        print('Nenhuma pasta selecionada.')
        return

    # Junta todos os PDFs das referencias
    pdfs_ref = []
    nomes_vistos = set()
    for p in pastas_ref:
        arquivos = sorted([f for f in os.listdir(p) if eh_pdf_valido(f)])
        for arq in arquivos:
            nome_norm = normalizar_texto(os.path.splitext(arq)[0])
            if nome_norm and nome_norm not in nomes_vistos:
                nomes_vistos.add(nome_norm)
                pdfs_ref.append(arq)
        print(f'  Referencia: {os.path.basename(p)} — {len(arquivos)} PDF(s)')
    print(f'  Total combinado: {len(pdfs_ref)} PDF(s) unicos')

    if not pdfs_ref:
        print('Nenhum PDF encontrado nas pastas referencia.')
        return

    # Seleciona pastas para corrigir
    print('\n2. Selecione a(s) pasta(s) com PDFs a CORRIGIR')
    print('   (selecione uma, depois outra, e assim por diante)')
    print('   (quanto terminar, clique em Cancelar)')
    pastas_corrigir = dialogo_pastas(
        'SELECIONE pasta COM PDFs A CORRIGIR',
        'SELECIONE OUTRA pasta (ou Cancelar)'
    )
    if not pastas_corrigir:
        print('Nenhuma pasta selecionada para correcao.')
        return
    print(f'  {len(pastas_corrigir)} pasta(s) selecionada(s)')
    for p in pastas_corrigir:
        print(f'    - {os.path.basename(p)}')

    # Log
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    log_path = os.path.join(os.path.dirname(pastas_ref[0]), f'log_correcao_{timestamp}.txt')
    log_lines = []
    log_lines.append('LOG DE CORRECAO DE NOMES DE PDFs')
    log_lines.append(f'Data: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}')
    log_lines.append(f'Modo: {"EXECUTAR COM LLM" if modo == "2" else "EXECUTAR"}')
    log_lines.append(f'Pastas referencia:')
    for p in pastas_ref:
        log_lines.append(f'  - {p}')
    log_lines.append(f'Pastas corrigidas:')
    for p in pastas_corrigir:
        log_lines.append(f'  - {p}')
    log_lines.append('')
    log_lines.append(f'{"=" * 70}')
    log_lines.append('')

    # Processa cada pasta
    total_geral_renomeados = 0
    total_geral_mantidos = 0
    total_geral_sem_match = 0
    
    cache_correcoes = {}

    for pasta_alvo in pastas_corrigir:
        nome_alvo = os.path.basename(pasta_alvo)
        print(f'\n--- Processando: {nome_alvo} ---')

        pdfs_alvo = sorted([f for f in os.listdir(pasta_alvo) if eh_pdf_valido(f)])
        if not pdfs_alvo:
            print(f'  Nenhum PDF nesta pasta.')

        renomeados = 0
        mantidos = 0
        sem_match = 0
        sem_match_nomes = []

        for arq_alvo in pdfs_alvo:
            if re.search(r'\d$', os.path.splitext(arq_alvo)[0].strip()):
                mantidos += 1
                msg = f'  {arq_alvo} -> IGNORADO (homonimo numerado)'
                print(msg)
                log_lines.append(msg)
                continue
                
            match, tipo = encontrar_melhor_match(arq_alvo, pdfs_ref)
            if not match:
                sem_match += 1
                sem_match_nomes.append(arq_alvo)
                continue

            if arq_alvo == match:
                mantidos += 1
                continue
                
            if modo in ('1', '2'):
                precisa_confirmar = False
                if 'DUVIDA' in tipo or 'similaridade' in tipo or 'var diff' in tipo:
                    precisa_confirmar = True
                    
                if precisa_confirmar:
                    stem_alvo, _ = extrair_variante(os.path.splitext(arq_alvo)[0])
                    norm_alvo = normalizar_texto(stem_alvo)
                    
                    if norm_alvo in cache_correcoes:
                        match = cache_correcoes[norm_alvo]
                        if match is None:
                            sem_match += 1
                            sem_match_nomes.append(arq_alvo)
                            continue
                    else:
                        print(f"\n      [DUVIDA] O arquivo '{arq_alvo}' (pasta: {nome_alvo}) eh a mesma pessoa que '{match}'?")
                        while True:
                            r = input(f"      (1 para SIM / 2 para NAO): ").strip()
                            if r in ('1', '2'): break
                        if r == '1':
                            cache_correcoes[norm_alvo] = match
                        else:
                            cache_correcoes[norm_alvo] = None
                            sem_match += 1
                            sem_match_nomes.append(arq_alvo)
                            continue

            # Vai renomear (NUNCA sobrescreve)
            origem = os.path.join(pasta_alvo, arq_alvo)
            nome_base, ext = os.path.splitext(match)
            destino = os.path.join(pasta_alvo, f'{nome_base}{ext}')

            # Se o destino ja existe e tem o mesmo tamanho, o correto ja esta la — perguntar antes de remover o duplicado
            if os.path.exists(destino) and os.path.getsize(origem) == os.path.getsize(destino):
                print(f"\n      [DUPLICADO] '{arq_alvo}' pode ser excluido, pois '{os.path.basename(destino)}' ja existe com o nome correto.")
                while True:
                    r_del = input(f"      Deseja EXCLUIR o arquivo com nome errado? (1 para SIM / 2 para NAO): ").strip()
                    if r_del in ('1', '2'): break
                if r_del == '1':
                    msg = f'  {arq_alvo} -> EXCLUIDO (duplicado, "{os.path.basename(destino)}" ja existia corretamente)'
                    print(msg)
                    log_lines.append(msg)
                    os.remove(origem)
                    renomeados += 1
                else:
                    msg = f'  {arq_alvo} -> MANTIDO (usuario optou por nao excluir o duplicado)'
                    print(msg)
                    log_lines.append(msg)
                    mantidos += 1
                continue

            # Adiciona sufixo numerico se destino ja existir (tamanho diferente = arquivo diferente)
            i = 2
            while os.path.exists(destino):
                destino = os.path.join(pasta_alvo, f'{nome_base}_{i}{ext}')
                i += 1

            msg = f'  {arq_alvo} -> {os.path.basename(destino)}  ({tipo})'
            print(msg)
            log_lines.append(msg)

            if True:  # modos 1 e 2 sempre executam
                shutil.move(origem, destino)

            renomeados += 1

        print(f'  Resultado: {renomeados} renomeado(s), {mantidos} mantido(s), {sem_match} sem match')
        if sem_match > 0:
            print(f'  SEM MATCH:')
            for n in sem_match_nomes:
                print(f'    {n}')
            log_lines.append(f'  SEM MATCH:')
            for n in sem_match_nomes:
                log_lines.append(f'    {n}')
        log_lines.append(f'  Resultado {nome_alvo}: {renomeados} renomeado(s), {mantidos} mantido(s), {sem_match} sem match')
        log_lines.append('')
        total_geral_renomeados += renomeados
        total_geral_mantidos += mantidos
        total_geral_sem_match += sem_match

    # Fim
    print(f'\n{"=" * 60}')
    print(f'  RESUMO GERAL')
    print(f'{"=" * 60}')
    print(f'  Total renomeados: {total_geral_renomeados}')
    print(f'  Total mantidos:   {total_geral_mantidos}')
    print(f'  Total sem match:  {total_geral_sem_match}')
    print(f'  Log salvo em: {log_path}')

    log_lines.append(f'{"=" * 70}')
    log_lines.append(f'RESUMO GERAL')
    log_lines.append(f'  Renomeados: {total_geral_renomeados}')
    log_lines.append(f'  Mantidos:   {total_geral_mantidos}')
    log_lines.append(f'  Sem match:  {total_geral_sem_match}')

    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(log_lines))



    input('\nPressione Enter para sair...')


if __name__ == '__main__':
    main()
