"""
Comparacao de nomes pessoa-a-pessoa (planilha x PDF), baseada em TOKENS.

Objetivo: maximizar a PRECISAO. Casos claros viram IGUAL; casos ambiguos
(sobrenome extra, nome do meio conflitante) viram DUVIDA para revisao humana
ou LLM, em vez de virarem match automatico e errado.

Diferencas em relacao a logica antiga:
  - NAO usa "substring" (que casava "ANA SILVA" dentro de "MARIANA SILVA ...").
  - NAO usa SequenceMatcher de caractere (que perde abreviacoes/ordem).
  - Ancora em primeiro nome + ultimo sobrenome; nomes do meio sao flexiveis.

Retorna um dos veredictos:
  "IGUAL"      -> nomes identicos (apos normalizacao) -> ENCONTRADO.
  "ABREVIACAO" -> mesma pessoa, mas com inicial/abreviatura/nome do meio
                  ausente/erro de digitacao -> ENCONTRADO COM ABREVIACAO.
  "DUVIDA"     -> pode ser a mesma pessoa; exige revisao (ou LLM)
                  -> POSSIVEL ERRO NOMINAL.
  "DIFERENTE"  -> pessoas diferentes.
"""
from utils import normalizar, levenshtein

CONECTORES = {'DE', 'DA', 'DO', 'DOS', 'DAS', 'E'}


def _tokens(nome):
    """Normaliza, remove conectores e numeros soltos, devolve lista de palavras."""
    return [t for t in normalizar(nome).split()
            if t not in CONECTORES and not t.isdigit()]


def _token_match(a, b):
    """Dois tokens (palavras) representam o mesmo nome?

    Aceita: igualdade; inicial (1 letra); abreviatura curta (prefixo de ate
    3 letras, ex 'APAR'->'APARECIDA'); erro de digitacao pequeno em palavras
    longas (levenshtein <= len//4, ou seja 1 erro a cada 4 letras).
    """
    if a == b:
        return True
    # inicial de 1 letra: 'J' x 'JOAO'
    if len(a) == 1 and b.startswith(a):
        return True
    if len(b) == 1 and a.startswith(b):
        return True
    # abreviatura curta (prefixo de <=3 letras): 'APAR' x 'APARECIDA'
    if 2 <= len(a) <= 3 and b.startswith(a):
        return True
    if 2 <= len(b) <= 3 and a.startswith(b):
        return True
    # tolerancia a erro de digitacao em palavras longas (Luiz x Luis)
    if len(a) >= 4 and len(b) >= 4:
        if levenshtein(a, b) <= min(len(a), len(b)) // 4:
            return True
    return False


def _todos_casam(menor, maior):
    """Todo token da lista menor casa com um token DISTINTO da maior?
    Devolve (bool, n_casados)."""
    usados = [False] * len(maior)
    casados = 0
    for tm in menor:
        for i, tM in enumerate(maior):
            if not usados[i] and _token_match(tm, tM):
                usados[i] = True
                casados += 1
                break
        else:
            return False, casados
    return True, casados


def _score(tp, tf):
    """Proporcao de tokens que casam (0..1), para ordenar candidatos."""
    menor, maior = (tp, tf) if len(tp) <= len(tf) else (tf, tp)
    _, casados = _todos_casam(menor, maior)
    return casados / max(len(maior), 1)


def comparar_nomes(nome_planilha, nome_pdf):
    """Compara dois nomes. Retorna (veredicto, score) com veredicto em
    {"IGUAL", "ABREVIACAO", "DUVIDA", "DIFERENTE"} e score float [0..1]."""
    tp = _tokens(nome_planilha)
    tf = _tokens(nome_pdf)
    if not tp or not tf:
        return ("DIFERENTE", 0.0)

    # Nomes identicos (apos normalizacao/remocao de conectores) -> ENCONTRADO.
    if tp == tf:
        return ("IGUAL", 1.0)

    score = _score(tp, tf)

    primeiro_ok = _token_match(tp[0], tf[0])
    ultimo_ok = _token_match(tp[-1], tf[-1])

    # Primeiro nome diferente -> pessoa diferente.
    if not primeiro_ok:
        return ("DIFERENTE", score)

    menor, maior = (tp, tf) if len(tp) <= len(tf) else (tf, tp)

    if ultimo_ok:
        # Mesmo primeiro e ultimo, mas NAO identicos. Os tokens do menor
        # precisam todos casar (subset) com o maior. Se sim, a diferenca e
        # apenas inicial/abreviatura/nome do meio ausente -> mesma pessoa.
        casam, _ = _todos_casam(menor, maior)
        if casam:
            return ("ABREVIACAO", score)
        # Ha um nome do meio que conflita (ex.: MARIA JOSE SILVA x MARIA HELENA SILVA)
        return ("DUVIDA", score)

    # Ultimo sobrenome difere.
    # Caso "sobrenome extra no fim": o ultimo do menor casa com o penultimo do
    # maior e todo o menor esta contido no maior. Ex.: JOSE PRAZERES x
    # JOSE PRAZERES VIEIRA -> pode ser parente/homonimo -> DUVIDA.
    if len(maior) > len(menor) and _token_match(menor[-1], maior[-2]):
        casam, _ = _todos_casam(menor, maior)
        if casam:
            return ("DUVIDA", score)

    # Sobrenomes finais realmente diferentes -> pessoas diferentes.
    return ("DIFERENTE", score)
