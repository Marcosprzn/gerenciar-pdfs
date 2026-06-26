import os
import re
import sys
import platform
import urllib.request
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image
import fitz  # PyMuPDF
import pytesseract

# Caminhos possíveis do Tesseract (64-bit e 32-bit)
TESSERACT_PATHS = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
]

def _get_windows_version():
    """Retorna a versão major do Windows (ex: 6 para Win8, 10 para Win10/11)."""
    try:
        ver = sys.getwindowsversion()
        return ver.major, ver.minor
    except Exception:
        return 10, 0  # assume Win10 se não conseguir detectar

def _get_tesseract_url():
    """Retorna a URL do instalador Tesseract compatível com o Windows atual."""
    major, minor = _get_windows_version()
    if major < 10:
        # Win8: SourceForge hospeda versão 3.02 compatível
        return "https://sourceforge.net/projects/tesseract-ocr-alt/files/tesseract-ocr-setup-3.02.02.exe/download"
    else:
        print(f"[INFO] Windows {major}.{minor} detectado - baixando Tesseract 5.4")
        return "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"

# URL da pagina de download para o usuario baixar manualmente no Win8
TESSERACT_WIN8_PAGE = "https://sourceforge.net/projects/tesseract-ocr-alt/files/tesseract-ocr-setup-3.02.02.exe/download"

def check_tesseract():
    # Tenta achar o executável nos caminhos conhecidos
    for path in TESSERACT_PATHS:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            print(f"[OK] Tesseract encontrado em: {path}")
            return True
    
    # Se nao existe, informa o usuario
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    major, minor = _get_windows_version()
    
    url = _get_tesseract_url()
    versao_label = "3.02 (compatível com Windows 8/7)" if major < 10 else "5.4 (versão atual)"
    
    msg = (
        "O Tesseract OCR não está instalado ou não foi encontrado.\n\n"
        f"Seu Windows foi detectado como versão {major}.{minor}.\n"
        f"Será baixada a versão Tesseract {versao_label}.\n\n"
        "Deseja baixar e instalar agora?"
    )
    if messagebox.askyesno("Tesseract OCR Ausente", msg):
        print(f"Baixando de: {url}")
        print("Por favor, aguarde...")
        try:
            exe_path = os.path.join(os.environ['TEMP'], 'tesseract_setup.exe')
            urllib.request.urlretrieve(url, exe_path)
            print("Download concluído! Iniciando a instalação...")
            messagebox.showinfo(
                "Instalação",
                "Siga os passos do instalador que vai abrir.\n"
                "APÓS CONCLUIR A INSTALAÇÃO, inicie este script novamente."
            )
            subprocess.Popen([exe_path])
        except Exception as e:
            import webbrowser
            webbrowser.open(url)
            messagebox.showinfo(
                "Download Manual",
                f"O download automático falhou (Erro: {e}).\n\n"
                "O link de download está sendo aberto no navegador.\n"
                "Salve o arquivo e instale manualmente.\n\n"
                "APÓS CONCLUIR A INSTALAÇÃO, inicie este script novamente."
            )
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

    # Tenta extrair o mes alvo do nome da pasta (ex: "01-SEFIP FGTS - COMPETÊNCIA 01-2008")
    nome_pasta = os.path.basename(pasta_pdfs)
    match_pasta = re.search(r'(\d{2})\s*[-/]\s*(\d{4})', nome_pasta)
    if match_pasta:
        mes_sugerido = f"{match_pasta.group(1)}/{match_pasta.group(2)}"
    else:
        mes_sugerido = "01/2007"

    mes_alvo = simpledialog.askstring("Mês Alvo", "Qual o mês e ano que deve estar na página?\n\n(Detectado automaticamente da pasta)", initialvalue=mes_sugerido, parent=root)
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
                print(f"[OK - MES: {mes_alvo}]")
                ok_list.append(f)
            else:
                # Verifica se é a foto correta da capa
                is_capa = re.search(r'(SEFIP|F\.?G\.?T\.?S|DEPOSITADO)', texto_ocr, re.IGNORECASE)
                if not is_capa:
                    print("[FOTO DIFERENTE DO ESPERADO]")
                    erro_list.append((f, "Foto diferente do esperado (não achou SEFIP/FGTS)"))
                    continue

                # Tenta capturar QUALQUER mês que esteja escrito após "MES ="
                match_generico = re.search(r'M[EÉ]?S\s*[=:-_]?\s*([A-Za-z0-9]{2}\s*[/|7lIi\-1]\s*[A-Za-z0-9]{4})', texto_ocr, re.IGNORECASE)
                if match_generico:
                    mes_lido = match_generico.group(1)
                    # Limpa os erros comuns de OCR
                    mes_limpo = mes_lido.replace('O','0').replace('o','0').replace('I','1').replace('l','1').replace('i','1').replace('|','/').replace('\\','/').replace(' ', '').replace('7','/').replace('-', '/')
                    # Como 7 pode ser barra, as vezes o ano fica quebrado se tiver 7, mas para visualizacao ta ok
                    
                    if len(mes_limpo) >= 7 and mes_limpo[-4:].isdigit() and mes_limpo[:2].isdigit():
                        mes_limpo = f"{mes_limpo[:2]}/{mes_limpo[-4:]}"
                    
                    print(f"[ERRADO - LIDO: {mes_limpo}]")
                    erro_list.append((f, f"Mês errado, lido: {mes_limpo}"))
                else:
                    # Fallback: tenta achar qualquer XX/XXXX na pagina solto
                    match_data = re.search(r'(\d{2})\s*[/|7lIi\-1]\s*(\d{4})', texto_ocr)
                    if match_data:
                        print(f"[ERRADO - ENCONTRADO DATA NA PAGINA: {match_data.group(1)}/{match_data.group(2)}]")
                        erro_list.append((f, f"Data perdida na pagina: {match_data.group(1)}/{match_data.group(2)}"))
                    else:
                        print("[MÊS NÃO ENCONTRADO]")
                        erro_list.append((f, "Mês não encontrado na OCR"))
                    
                    # Salva o texto bruto em um arquivo de log para ajudar na depuração
                    try:
                        with open(os.path.join(pasta_pdfs, "DEBUG_ERROS_OCR.txt"), "a", encoding="utf-8") as logf:
                            logf.write(f"=== TEXTO BRUTO DO ARQUIVO: {f} ===\n")
                            logf.write(texto_ocr)
                            logf.write("\n===========================================\n\n")
                    except:
                        pass
                    
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
        print(f"\n[!] DICA: Um arquivo 'DEBUG_ERROS_OCR.txt' foi criado na pasta {pasta_pdfs}.")
        print("    Abra ele para ver EXATAMENTE o que a inteligência leu e me avise para ajustarmos o código!")
            
    print("\nAuditoria finalizada!")
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    run()
