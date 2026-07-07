"""
Funcoes utilitarias compartilhadas de normalizacao e comparacao de nomes.

Antes, `normalizar`/`levenshtein`/`similaridade` estavam copiadas em varios
modulos (app.py, confinicial.py, conferencia_unico.py, corrigir_nomes_pdfs.py),
com pequenas divergencias entre elas. Isso foi centralizado aqui para que uma
correcao valha para todo o sistema.

A versao de `normalizar` abaixo e o "superset" das antigas: faz tudo o que
elas faziam e ainda trata travessoes (– —), portanto e compativel com todas.
"""
import unicodedata
from difflib import SequenceMatcher

__all__ = ["normalizar", "normalizar_texto", "levenshtein", "similaridade", "calcular_similaridade"]


def normalizar(texto):
    """Remove acentos, coloca em MAIUSCULAS e padroniza pontuacao/espacos.

    Regras:
      - Remove acentos (NFKD + descarte de combinantes).
      - Remove '.', '(' e ')'.
      - Troca '-', '_' e travessoes (– —) por espaco.
      - Colapsa espacos multiplos.

    Ex.: 'José A. Silva-Souza (REC)' -> 'JOSE A SILVA SOUZA REC'
    """
    if not isinstance(texto, str):
        return ""
    nfkd = unicodedata.normalize('NFKD', texto)
    s = "".join(c for c in nfkd if not unicodedata.combining(c)).upper()
    s = (s.replace('.', '')
          .replace('(', '')
          .replace(')', '')
          .replace('-', ' ')
          .replace('–', ' ')   # travessao curto –
          .replace('—', ' ')   # travessao longo —
          .replace('_', ' '))
    return " ".join(s.split())


# Alias historico: alguns modulos usavam o nome `normalizar_texto`.
normalizar_texto = normalizar


def levenshtein(a, b):
    """Distancia de edicao entre duas strings (implementacao O(m*n) em espaco O(n))."""
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


def similaridade(a, b):
    """Razao de similaridade [0..1] entre duas strings (difflib)."""
    return SequenceMatcher(None, a, b).ratio()


# Alias historico: alguns modulos usavam o nome `calcular_similaridade`.
calcular_similaridade = similaridade
