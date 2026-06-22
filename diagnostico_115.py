#!/usr/bin/env python3
"""Diagnostico 115 - Abre dialogos e gera log automaticamente"""
import sys, os, tkinter as tk
from tkinter import filedialog
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import confinicial, traceback

root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); root.update()

pasta_pdfs = filedialog.askdirectory(title="Selecione a PASTA DE PDFs")
if not pasta_pdfs:
    print("Nenhuma pasta selecionada.")
    sys.exit(1)

planilha = filedialog.askopenfilename(title="Selecione a PLANILHA", filetypes=[("Planilhas", "*.xls *.xlsx *.ods")])
if not planilha:
    print("Nenhuma planilha selecionada.")
    sys.exit(1)

root.destroy()

log = []
log.append("=" * 70)
log.append("DIAGNOSTICO 115")
log.append("=" * 70)
log.append(f"Planilha: {planilha}")
log.append(f"Pasta:    {pasta_pdfs}")
log.append("")

# 1. Lista todos os arquivos
log.append("--- ARQUIVOS NA PASTA ---")
try:
    for f in sorted(os.listdir(pasta_pdfs)):
        if not f.lower().endswith(('.pdf', '.tif', '.tiff')):
            log.append(f"  (ignorado) {f}")
            continue
        stem, var = confinicial.extrair_variante(os.path.splitext(f)[0])
        n_norm = confinicial.normalizar_texto(stem)
        var_tag = var or "None"
        log.append(f"  {f:55s} norm={n_norm!r:40s} var={var_tag}")
except Exception as e:
    log.append(f"  ERRO: {e}")

# 2. Le planilha - pessoas 115
log.append("")
log.append("--- PESSOAS 115 NA PLANILHA ---")
try:
    df = confinicial.ler_planilha_fgts(planilha)
    if df is not None:
        log.append(f"Total linhas: {len(df)}")
        log.append(f"Colunas: {list(df.columns)}")
        tipo_counts = df['TIPO_LISTA'].value_counts().to_dict()
        for k, v in tipo_counts.items():
            log.append(f"  {k}: {v}")
        df_115 = df[df['TIPO_LISTA'] == '115']
        for _, row in df_115.iterrows():
            nome = row['NOMES']
            n_norm = confinicial.normalizar_texto(nome)
            log.append(f"  115: {nome!r:45s} norm={n_norm!r}")
    else:
        log.append("  ERRO: Nao foi possivel ler a planilha")
except Exception as e:
    log.append(f"  ERRO: {e}")
    log.append(traceback.format_exc())

# 3. Simula matching
log.append("")
log.append("--- SIMULACAO DE MATCHING ---")
try:
    df = confinicial.ler_planilha_fgts(planilha)
    arquivos = []
    for f in sorted(os.listdir(pasta_pdfs)):
        if not f.lower().endswith(('.pdf', '.tif', '.tiff')): continue
        if any(p in os.path.splitext(f)[0].upper() for p in confinicial.EXCLUIR_PADROES): continue
        stem, var = confinicial.extrair_variante(os.path.splitext(f)[0])
        arquivos.append({"real": f, "norm": confinicial.normalizar_texto(stem), "variante": var})
    
    log.append(f"Total arquivos validos: {len(arquivos)}")
    rec_count = len([a for a in arquivos if a['variante'] == 'REC115'])
    sem_count = len([a for a in arquivos if not a['variante']])
    log.append(f"  REC115: {rec_count}")
    log.append(f"  Sem variante: {sem_count}")
    
    if df is not None:
        df_115 = df[df['TIPO_LISTA'] == '115']
        for _, row in df_115.iterrows():
            nome = row['NOMES']
            n_norm = confinicial.normalizar_texto(nome)
            log.append("")
            log.append(f"  >> {nome}")
            
            for grupo_nome, grupo_filtro in [("REC115", lambda a: a['variante'] == 'REC115'),
                                              ("SEM VAR", lambda a: not a['variante']),
                                              ("TODOS", lambda a: True)]:
                lista = [a for a in arquivos if grupo_filtro(a)]
                match = None
                for a in lista:
                    if n_norm in a["norm"] or a["norm"] in n_norm:
                        match = a; break
                if match:
                    log.append(f"     {grupo_nome:8s}: ENCONTRADO -> {match['real']} (norm={match['norm']!r})")
                else:
                    log.append(f"     {grupo_nome:8s}: nenhum match")
                    if grupo_nome == "TODOS":
                        log.append(f"     Top 5 candidatos:")
                        arqs_ord = sorted(arquivos, key=lambda a: confinicial.calcular_similaridade(n_norm, a["norm"]), reverse=True)
                        for a in arqs_ord[:5]:
                            sim = confinicial.calcular_similaridade(n_norm, a["norm"])
                            abrev = confinicial.verificar_abreviacao(n_norm, a["norm"])
                            log.append(f"       sim={sim:.0%} abrev={abrev} -> {a['real']} (norm={a['norm']!r})")
except Exception as e:
    log.append(f"  ERRO: {e}")
    log.append(traceback.format_exc())

log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'diagnostico_115.txt')
with open(log_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(log))

print(f"Log salvo em: {log_path}")
print("Cole o conteudo do arquivo diagnostico_115.txt aqui.")
input("\nPressione Enter para sair...")
