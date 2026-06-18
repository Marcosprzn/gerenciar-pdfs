"""Diagnostico rapido - versao CMD (sem tkinter).
Use: python diagnostico_cmd.py "CAMINHO"  (caminho completo da pasta)
"""
import os, sys, re

if len(sys.argv) < 2:
    pasta = input("Cole o caminho da pasta com os PDFs: ").strip()
    if not pasta:
        pasta = os.getcwd()
else:
    pasta = sys.argv[1]

pasta = pasta.strip('"').strip()

print("=" * 60)
print("DIAGNOSTICO - CONFERENCIA PDF vs PLANILHA")
print("=" * 60)
print()
print(f"PASTA ANALISADA: {pasta}")
print(f"NOME DA PASTA:   {os.path.basename(pasta)}")
print()

if not os.path.exists(pasta):
    print("ERRO FATAL: Esta pasta NAO EXISTE!")
    print("Verifique se o caminho esta correto.")
    sys.exit(1)

if not os.path.isdir(pasta):
    print("ERRO FATAL: Isto nao e uma pasta!")
    sys.exit(1)

print("CONTEUDO COMPLETO DA PASTA:")
print("-" * 60)
try:
    itens = os.listdir(pasta)
    for f in sorted(itens):
        full = os.path.join(pasta, f)
        if os.path.isdir(full):
            print(f"  [PASTA]     {f}/")
        else:
            nome, ext = os.path.splitext(f)
            print(f"  [ARQUIVO]   {f}   (extensao: '{ext}')")
except Exception as e:
    print(f"  ERRO AO LISTAR: {e}")
print("-" * 60)
print(f"  Total de itens: {len(itens)}")
print()

print("ARQUIVOS QUE O SCRIPT VAI USAR (excluindo planilhas):")
print("-" * 60)
EXT_PLANILHAS = ('.xls', '.xlsx', '.ods')
arquivos_uteis = []
try:
    for f in sorted(itens):
        full = os.path.join(pasta, f)
        if os.path.isfile(full):
            if not f.lower().endswith(EXT_PLANILHAS):
                arquivos_uteis.append(f)
                print(f"  [VAI USAR] {f}")
            else:
                print(f"  [PLANILHA] {f}  (ignorada como nome)")
        elif os.path.isdir(full):
            print(f"  [PASTA]    {f}/")
except Exception as e:
    print(f"  ERRO: {e}")

print("-" * 60)
print(f"  Arquivos uteis encontrados: {len(arquivos_uteis)}")
print()

if arquivos_uteis:
    print("NOMES EXTRAIDOS (nome do arquivo sem extensao):")
    for f in arquivos_uteis:
        nome = os.path.splitext(f)[0]
        print(f"  - {nome}")
    print()

print("VERIFICACAO DE DATA NO NOME DA PASTA:")
nome_pasta = os.path.basename(pasta)
match = re.search(r'(\d{2})[-_.](\d{4})', nome_pasta)
if match:
    mes, ano = match.group(1), match.group(2)
    print(f"  DATA ENCONTRADA: {mes}-{ano}")
    print(f"  O script vai procurar planilha com: {mes}_{ano} ou {mes}_{ano[-2:]}")
else:
    print(f"  DATA NAO ENCONTRADA em: '{nome_pasta}'")
    print(f"  Padrao esperado: 'XX-XXXX' (ex: '01-2007') em algum lugar do nome")
    print(f"  O script vai permitir selecionar a planilha manualmente.")
print()

print("PLANILHAS ENCONTRADAS NA PASTA:")
print("-" * 60)
planilhas = [f for f in itens if f.lower().endswith(('.xls', '.xlsx', '.ods'))]
if planilhas:
    for p in planilhas:
        print(f"  [PLANILHA] {p}")
else:
    print("  Nenhuma planilha (.xls/.xlsx/.ods) encontrada nesta pasta.")
print("-" * 60)
print()

print("=" * 60)
print("FIM DO DIAGNOSTICO")
print("=" * 60)
print()
print("COLE TODA A SAIDA ACIMA AQUI PARA EU ANALISAR.")
print()

input("Pressione Enter para sair...")
