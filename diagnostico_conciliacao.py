#!/usr/bin/env python3
"""
DIAGNOSTICO DE CONCILIACAO  (para enviar ao suporte)

O que faz: compara UMA planilha FGTS com UMA pasta de PDFs e gera um relatorio
de texto mostrando TODA a estrutura (arquivos, nomes, o que casou/nao casou e o
porque). Nao altera nada — so le e escreve um .txt.

Como usar:
    python diagnostico_conciliacao.py
    1) Selecione a PLANILHA (.xls/.xlsx/.ods) da competencia.
    2) Selecione a PASTA com os PDFs daquela mesma competencia.
    3) Envie o arquivo 'diagnostico_resultado.txt' gerado.

O PIS aparece mascarado no relatorio (so os ultimos digitos).
"""
import os
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import confinicial
from match_nomes import comparar_nomes

EXT = getattr(confinicial, "EXT", ('.pdf', '.tif', '.tiff'))


def _mascara_pis(pis):
    d = ''.join(ch for ch in str(pis) if ch.isdigit())
    if len(d) < 4:
        return '***'
    return '***' + d[-4:]


def _selecionar():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        planilha = filedialog.askopenfilename(
            title="1) Selecione a PLANILHA FGTS",
            filetypes=[("Planilhas", "*.xlsx *.xls *.ods"), ("Todos", "*.*")])
        if not planilha:
            return None, None
        pasta = filedialog.askdirectory(
            title="2) Selecione a PASTA com os PDFs desta competencia")
        root.destroy()
        return planilha, pasta
    except Exception as e:
        print("Nao foi possivel abrir as janelas de selecao:", e)
        print("Uso alternativo: python diagnostico_conciliacao.py <planilha> <pasta_pdfs>")
        if len(sys.argv) >= 3:
            return sys.argv[1], sys.argv[2]
        return None, None


def main():
    if len(sys.argv) >= 3:
        planilha, pasta = sys.argv[1], sys.argv[2]
    else:
        planilha, pasta = _selecionar()

    if not planilha or not pasta:
        print("Selecao cancelada.")
        return

    linhas = []
    def w(s=""):
        linhas.append(str(s))

    w("=" * 70)
    w("DIAGNOSTICO DE CONCILIACAO")
    w("Data: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    w("Planilha: " + os.path.basename(planilha))
    w("Pasta PDFs: " + pasta)
    w("=" * 70)

    # --- Carrega a planilha ---
    try:
        df = confinicial.ler_planilha_fgts(planilha)
    except Exception as e:
        df = None
        w("\n[ERRO] Falha ao ler a planilha: " + repr(e))
    if df is None or len(df) == 0:
        w("\n[ERRO] Planilha vazia ou nao lida (verifique a aba/estrutura).")
        _salvar(linhas, pasta)
        return

    # --- Roda a conferencia real ---
    df_res = confinicial.verificar_pdfs(df.copy(), pasta)
    excluidos = df_res.attrs.get('arquivos_excluidos', []) if hasattr(df_res, 'attrs') else []

    # --- [1] Arquivos PDF na pasta ---
    arquivos = [f for f in sorted(os.listdir(pasta)) if f.lower().endswith(EXT)]
    disponiveis = []  # (real, norm, variante)
    w("\n--- [1] ARQUIVOS PDF NA PASTA (%d) ---" % len(arquivos))
    w("%-45s | %-30s | %s" % ("ARQUIVO", "NORMALIZADO (usado no match)", "VARIANTE"))
    w("-" * 90)
    for f in arquivos:
        stem, var = confinicial.extrair_variante(os.path.splitext(f)[0])
        norm = confinicial.normalizar_texto(stem)
        excl = " [IGNORADO]" if f in excluidos else ""
        if not excl:
            disponiveis.append((f, norm, var))
        w("%-45s | %-30s | %s%s" % (f[:45], norm[:30], var or "-", excl))

    if excluidos:
        w("\n--- [2] IGNORADOS por padrao (SEFIP/GUIA/IMG_/etc.) (%d) ---" % len(excluidos))
        for f in excluidos:
            w("  " + f)

    # --- [3] Nomes na planilha ---
    w("\n--- [3] NOMES NA PLANILHA (%d) ---" % len(df))
    w("%-38s | %-30s | %-8s | %s" % ("NOME", "NORMALIZADO", "TIPO", "PIS"))
    w("-" * 90)
    for _, r in df.iterrows():
        nome = str(r.get('NOMES', ''))
        norm = confinicial.normalizar_texto(nome)
        tipo = str(r.get('TIPO_LISTA', 'PADRAO'))
        w("%-38s | %-30s | %-8s | %s" % (nome[:38], norm[:30], tipo, _mascara_pis(r.get('PIS', ''))))

    # --- [4] Resultado da conferencia ---
    w("\n--- [4] RESULTADO DA CONFERENCIA ---")
    w("%-38s | %-34s | %s" % ("NOME", "STATUS", "PDF ENCONTRADO"))
    w("-" * 100)
    for _, r in df_res.iterrows():
        w("%-38s | %-34s | %s" % (
            str(r.get('NOMES', ''))[:38],
            str(r.get('Status PDF', ''))[:34],
            str(r.get('Nome do Arquivo Encontrado', ''))))

    # --- [5] Por que nao casou? (nao encontrados / duvidas) ---
    w("\n--- [5] DIAGNOSTICO: melhores candidatos para os NAO-ENCONTRADOS / DUVIDAS ---")
    for _, r in df_res.iterrows():
        status = str(r.get('Status PDF', ''))
        nome = str(r.get('NOMES', ''))
        if not nome:
            continue
        if status.startswith('ENCONTRADO'):
            continue  # ja casou bem
        norm = confinicial.normalizar_texto(nome)
        # compara contra todos os PDFs disponiveis
        ranking = []
        for real, pnorm, var in disponiveis:
            veredicto, sc = comparar_nomes(norm, pnorm)
            ranking.append((sc, veredicto, real, pnorm))
        ranking.sort(reverse=True)
        w("\n  PLANILHA: %s   [%s]" % (nome, status))
        w("           norm: %s" % norm)
        if not ranking:
            w("           (nenhum PDF disponivel para comparar)")
        for sc, veredicto, real, pnorm in ranking[:3]:
            w("     -> %-9s (score %.2f)  %s" % (veredicto, sc, real))

    _salvar(linhas, pasta)


def _salvar(linhas, pasta):
    saida = os.path.join(pasta, "diagnostico_resultado.txt")
    try:
        with open(saida, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        print("\n" + "=" * 60)
        print("Relatorio salvo em:")
        print("  " + saida)
        print("Envie esse arquivo .txt para o suporte.")
        print("=" * 60)
    except Exception as e:
        print("Erro ao salvar o relatorio:", e)
        print("\n".join(linhas))


if __name__ == "__main__":
    main()
