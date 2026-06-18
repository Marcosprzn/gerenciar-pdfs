"""Diagnostico - mostra EXATAMENTE o que o script ve na pasta."""
import os, sys, tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
root.update()

pasta = filedialog.askdirectory(title='SELECIONE A PASTA QUE TEM OS PDFS')
if not pasta:
    print("Nenhuma pasta selecionada.")
    root.destroy()
    sys.exit(1)

root.destroy()

print(f"\nPASTA SELECIONADA: {pasta}")
print(f"  Nome da pasta: {os.path.basename(pasta)}")
print()

# Lista TUDO na pasta
print("CONTEUDO COMPLETO DA PASTA:")
print("-" * 60)
try:
    itens = os.listdir(pasta)
    for f in sorted(itens):
        full = os.path.join(pasta, f)
        if os.path.isdir(full):
            print(f"  [SUBPSTA] {f}")
        else:
            nome, ext = os.path.splitext(f)
            print(f"  [{ext.upper() or 'SEM EXT'}] {f}")
except Exception as e:
    print(f"  ERRO: {e}")

print("-" * 60)
print(f"  Total: {len(itens)} itens")
print()

# Testa se o script enxerga os .pdf
import re
print("TESTE DE DETECCAO:")
print("-" * 60)
pdfs_encontrados = 0
try:
    for f in sorted(itens):
        full = os.path.join(pasta, f)
        if os.path.isfile(full):
            nome, ext = os.path.splitext(f)
            detectado = f.lower().endswith('.pdf')
            if detectado:
                pdfs_encontrados += 1
                print(f"  [DETECTADO] {f}")
            else:
                print(f"  [ignorado]  {f}  (extensao: '{ext}')")
        elif os.path.isdir(full):
            print(f"  [PASTA]     {f}")
except Exception as e:
    print(f"  ERRO: {e}")

print("-" * 60)
print(f"  PDFs detectados pelo script: {pdfs_encontrados}")
print()

# Testa se a data e reconhecida
nome_pasta = os.path.basename(pasta)
match = re.search(r'(\d{2})[-_.](\d{4})', nome_pasta)
if match:
    print(f"DATA RECONHECIDA: {match.group(1)}-{match.group(2)}")
    print(f"  O script vai procurar uma planilha com '01_07' ou '012007' no nome")
else:
    print("DATA NAO RECONHECIDA no nome da pasta.")
    print("  Padrao esperado: 'XX-XXXX' (ex: '01-2007') em algum lugar do nome")
    print("  Solucao: o script vai perguntar se quer selecionar a planilha manualmente")
print()

print("=" * 60)
print("COPIE E COLE ESTA SAIDA PARA EU ANALISAR")
print("=" * 60)
input("\nPressione Enter para sair...")
