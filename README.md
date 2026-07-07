# Gerenciador de PDFs — Conferência FGTS

Ferramenta local (roda no seu PC) para **conferir e renomear PDFs de trabalhadores**
a partir de planilhas de referência (FGTS/SEFIP). Compara nomes das planilhas com os
nomes dos arquivos digitalizados, aponta faltantes e corrige nomes automaticamente.

## Como rodar

1. Instale o Python 3 (com Tkinter — já vem no instalador oficial do Windows).
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. (Opcional, para OCR) Instale o **Tesseract** no sistema — necessário para o `pytesseract`.
4. Inicie o app:
   ```bash
   iniciar.bat
   ```
   ou
   ```bash
   python app.py
   ```
   Depois abra o navegador em `http://localhost:5000`.

## Configuração da chave de IA (opcional)

A verificação de nomes por IA (Gemini) é **opcional**. Para ativá-la:

- Copie `config_llm.example.py` para `config_llm.py` e cole sua chave, **ou**
- Defina a variável de ambiente `GEMINI_API_KEY`.

> ⚠️ **Nunca** commite o `config_llm.py` — ele está no `.gitignore`. Se sua chave
> já vazou em algum commit antigo, **gere uma nova** em https://aistudio.google.com/apikey.

## Estrutura dos arquivos

| Arquivo | Papel |
|---|---|
| `app.py` | Servidor web Flask — interface principal (25 rotas de API). |
| `templates/index.html` | Frontend (página única). |
| `utils.py` | **Funções compartilhadas** de normalização e comparação de nomes (`normalizar`, `levenshtein`, `similaridade`). |
| `confinicial.py` | Conferência inicial planilha × PDFs por competência (mês/ano). |
| `conferencia_unico.py` | Conferência de um mês/pasta específico. |
| `corrigir_nomes_pdfs.py` | Renomeia PDFs para bater com os nomes da planilha. |
| `corrigir_nomes_por_pis.py` | Renomeia/associa PDFs usando o número do PIS. |
| `consolidar_faltantes.py` | Gera `consolidado_faltantes.xlsx` com quem falta PDF/Excel. |
| `verificar_mes_scans.py` | Verifica scans de um mês. |
| `diagnostico_115.py` | Script de diagnóstico de casos "REC 115". |
| `llm_matcher.py` | Ponte opcional para o Gemini (verificação de nomes). |
| `config_llm.py` | Sua chave de API (local, **não versionado**). |

## Logs

Erros são gravados em `app.log` (rotativo, na pasta do projeto), além do console.
Consulte esse arquivo quando algo falhar.

## Convenções importantes

- **Não duplicar funções de normalização.** `normalizar`/`levenshtein`/`similaridade`
  moram só em `utils.py`. Importe de lá em vez de recopiar — assim uma correção vale
  para todo o sistema e a conferência permanece consistente.
