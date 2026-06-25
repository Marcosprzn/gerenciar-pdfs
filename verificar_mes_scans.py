import os
import re
import sys
import webbrowser
import urllib.request
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image
import fitz  # PyMuPDF
import pytesseract

# Caminho padrão do Tesseract no Windows
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TESSERACT_URL = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"

def check_tesseract():
    if os.path.exists(TESSERACT_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        return True
    
    # Se nao existe, informa o usuario
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    msg = (
        "O Tesseract OCR não está instalado ou não foi encontrado em:\n"
        f"{TESSERACT_PATH}\n\n"
        "O Tesseract é a 'lupa' necessária para ler as imagens dos PDFs offline.\n"
        "Deseja baixar o instalador oficial agora?"
    )
    if messagebox.askyesno("Tesseract OCR Ausente", msg):
        print("Iniciando o download do instalador. Por favor, aguarde...")
        try:
            exe_path = os.path.join(os.environ['TEMP'], 'tesseract_setup.exe')
            urllib.request.urlretrieve(TESSERACT_URL, exe_path)
            print("Download concluído! Iniciando a instalação...")
            messagebox.showinfo("Instalação", "Siga os passos do instalador que acabou de abrir na sua tela. APÓS CONCLUIR A INSTALAÇÃO, inicie este script novamente.")
            # Inicia o instalador
            subprocess.Popen([exe_path])
        except Exception as e:
            messagebox.showerror("Erro de Download", f"Não foi possível baixar automaticamente.\nErro: {e}\n\nBaixe manualmente pesquisando por 'Tesseract UB Mannheim'.")
    return False

def gerar_regex_mes(mes_alvo):
    # Ex: "01/2007"
    mes_alvo = mes_alvo.strip()
    
    # Substitui a barra por um regex flexível caso a OCR confunda a barra
    partes = mes_alvo.split('/')
    if len(partes) == 2:
        mes, ano = partes[0], partes[1]
        
        # As vezes OCR confunde o '0' com 'O' ou barra com '1', '7', 'l', 'i'
        mes_regex = mes.replace('0', '[0O]')
        ano_regex = ano.replace('0', '[0O]')
        
        # M[EÉ]S\s*[=:-_]?\s*01\s*[/|7lIi\-]\s*2007
        padrao = r'M[EÉ]?S\s*[=:-_]?\s*' + mes_regex + r'\s*[/|7lIi\-1]\s*' + ano_regex
        return padrao
    
    # Fallback se não tiver barra (ex: "01-2007")
    return r'M[EÉ]?S\s*[=:-_]?\s*' + re.escape(mes_alvo)

def run():
    if not check_tesseract():
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    pasta_pdfs = filedialog.askdirectory(title="Selecione a pasta com os PDFs para auditar")
    if not pasta_pdfs:
        print("Operação cancelada.")
        return

    mes_alvo = simpledialog.askstring("Mês Alvo", "Qual o mês e ano que deve estar na página?\n\nExemplo: 01/2007", parent=root)
    if not mes_alvo:
        print("Mês não informado. Operação cancelada.")
        return
        
    padrao_regex = gerar_regex_mes(mes_alvo)
    print(f"\nPasta selecionada: {pasta_pdfs}")
    print(f"Buscando por: {mes_alvo} (Regex OCR: {padrao_regex})\n")
    print("Iniciando auditoria... Isso pode levar alguns segundos por PDF.\n")
    
    arquivos = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]
    
    if not arquivos:
        print("Nenhum PDF encontrado na pasta.")
        return

    ok_list = []
    erro_list = []
    
    for f in arquivos:
        caminho_completo = os.path.join(pasta_pdfs, f)
        print(f"Lendo: {f}... ", end='', flush=True)
        
        try:
            # Abre o PDF e pega a página 1 (index 0)
            doc = fitz.open(caminho_completo)
            if len(doc) == 0:
                print("[VAZIO]")
                erro_list.append((f, "PDF vazio"))
                continue
                
            page = doc[0]
            # Extrai como imagem
            # O matrix(2,2) aumenta a resolução pra ajudar a OCR
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            
            # Converte o pixmap do PyMuPDF para Image do Pillow
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Executa o OCR
            texto_ocr = pytesseract.image_to_string(img, lang='por+eng')
            
            # Pesquisa o padrao
            if re.search(padrao_regex, texto_ocr, re.IGNORECASE):
                print("[OK]")
                ok_list.append(f)
            else:
                # Fallback: OCR as vezes nao pega a palavra MES, tenta achar apenas o "01/2007" proximo a nada
                partes = mes_alvo.split('/')
                if len(partes) == 2:
                    apenas_data_regex = partes[0].replace('0', '[0O]') + r'\s*[/|7lIi\-1]\s*' + partes[1].replace('0', '[0O]')
                    if re.search(apenas_data_regex, texto_ocr):
                        print("[OK (Apenas Data Encontrada sem palavra MES)]")
                        ok_list.append(f)
                    else:
                        print("[MÊS NÃO ENCONTRADO]")
                        erro_list.append((f, "Mês não encontrado na OCR"))
                else:
                    print("[MÊS NÃO ENCONTRADO]")
                    erro_list.append((f, "Mês não encontrado na OCR"))
                    
        except Exception as e:
            print(f"[ERRO DE LEITURA: {str(e)}]")
            erro_list.append((f, f"Erro ao processar: {str(e)}"))

    print("\n" + "="*50)
    print("                RESUMO DA AUDITORIA")
    print("="*50)
    print(f"Total de arquivos analisados: {len(arquivos)}")
    print(f"OK (Mês {mes_alvo} confere): {len(ok_list)}")
    print(f"COM PROBLEMA ou NÃO ENCONTRADO: {len(erro_list)}")
    print("="*50)
    
    if erro_list:
        print("\nARQUIVOS COM POSSÍVEL ERRO:")
        for arq, motivo in erro_list:
            print(f" - {arq} ({motivo})")
            
    print("\nAuditoria finalizada!")
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    run()
