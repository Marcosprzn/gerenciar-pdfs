#!/usr/bin/env python3
"""
Corrige nomes de arquivos com codificacao URL (%20, %2520, etc).
Ex: JOSE%2520PEREIRA%2520DO%2520NASCIMENTO -> JOSE PEREIRA DO NASCIMENTO
"""
import os, sys
import tkinter as tk
from tkinter import filedialog
from urllib.parse import unquote

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)
root.update()

pasta = filedialog.askdirectory(title="Selecione a PASTA com os arquivos com % no nome")
root.destroy()

if not pasta:
    print("Nenhuma pasta selecionada.")
    sys.exit(1)

def decodificar(nome):
    """Decodifica URL encoding (inclusive duplo como %2520)."""
    antigo = None
    while antigo != nome:
        antigo = nome
        nome = unquote(nome)
    return nome

renomeados = 0
erros = 0

for f in sorted(os.listdir(pasta)):
    caminho_original = os.path.join(pasta, f)
    if os.path.isdir(caminho_original):
        continue

    nome_sem_ext = os.path.splitext(f)[0]
    ext = os.path.splitext(f)[1]

    if '%' not in nome_sem_ext:
        continue

    nome_decod = decodificar(nome_sem_ext)
    if nome_decod == nome_sem_ext:
        continue

    novo_nome = nome_decod + ext
    caminho_novo = os.path.join(pasta, novo_nome)

    if os.path.exists(caminho_novo):
        print(f"  [CONFLITO] {f} -> {novo_nome} ja existe. Pulando.")
        erros += 1
        continue

    try:
        os.rename(caminho_original, caminho_novo)
        print(f"  {f:55s} -> {novo_nome}")
        renomeados += 1
    except Exception as e:
        print(f"  [ERRO] {f}: {e}")
        erros += 1

print(f"\nResumo: {renomeados} renomeado(s), {erros} erro(s)")
input("\nPressione Enter para sair...")
