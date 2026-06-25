#!/usr/bin/env python3
"""
Consolida todos os relatorios .xlsx e lista os funcionarios
que estao faltando nos PDFs, organizados por mes e nome.
"""
import os, sys, re
import tkinter as tk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)
root.update()

arquivos = filedialog.askopenfilenames(
    title="Selecione os relatorios .xlsx",
    filetypes=[("Relatorios", "*.xlsx"), ("Todos", "*.*")]
)
root.destroy()

if not arquivos:
    print("Nenhum arquivo selecionado.")
    sys.exit(1)

todas = []
for arq in sorted(arquivos):
    nome = os.path.basename(arq)
    # Extrai mes/ano do nome do arquivo
    m = re.search(r"(?<!\d)(\d{2})[-_ ](\d{2,4})(?!\d)", nome)
    mes_ano = None
    if m:
        mes = m.group(1)
        ano = "20" + m.group(2) if len(m.group(2)) == 2 else m.group(2)
        mes_ano = f"{mes}/{ano}"

    try:
        df = pd.read_excel(arq, sheet_name='Dados')
    except Exception:
        try:
            df = pd.read_excel(arq)
        except Exception as e:
            print(f"  ERRO ao ler {nome}: {e}")
            continue

    col_status = None
    for col in df.columns:
        if 'status' in str(col).lower():
            col_status = col
            break
    if not col_status:
        print(f"  Coluna Status nao encontrada em {nome}")
        continue

    for _, row in df.iterrows():
        status = str(row[col_status])
        if 'NAO ENCONTRADO' in status or 'NAO ENCONTRADO' in status:
            nome_pessoa = str(row.get('NOMES', row.get('Nome', '')))
            if nome_pessoa and nome_pessoa not in ('', 'nan', 'None'):
                todas.append({
                    'mes_ano': mes_ano or nome,
                    'nome': nome_pessoa,
                    'arquivo': nome,
                    'status': status,
                })

if not todas:
    print("Nenhum faltante encontrado.")
    sys.exit(0)

print(f"\nTotal de registros faltantes: {len(todas)}\n")
print(f"{'Mes/Ano':<12} {'Nome':<50} {'Status':<35} {'Arquivo'}")
print("=" * 120)
for item in sorted(todas, key=lambda x: (x['mes_ano'], x['nome'])):
    print(f"{item['mes_ano']:<12} {item['nome']:<50} {item['status']:<35} {item['arquivo']}")

# Gera tambem um agrupado por pessoa
from collections import Counter
pessoas = Counter()
for item in todas:
    pessoas[item['nome']] += 1

print(f"\n\nPessoas que mais aparecem como faltantes:")
print(f"{'Nome':<50} {'Vezes':<8} {'Meses'}")
print("=" * 70)
for nome, count in pessoas.most_common(30):
    meses = [item['mes_ano'] for item in todas if item['nome'] == nome]
    print(f"{nome:<50} {count:<8} {', '.join(meses)}")

# Salva CSV
csv_path = os.path.join(os.path.dirname(arquivos[0]), 'consolidado_faltantes.csv')
import csv
with open(csv_path, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=['mes_ano', 'nome', 'status', 'arquivo'])
    w.writeheader()
    w.writerows(todas)
print(f"\nCSV salvo: {csv_path}")
input("\nPressione Enter para sair...")
