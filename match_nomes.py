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
# Rotulos de grupo/codigo que aparecem no nome do PDF mas NAO fazem parte do
# nome da pessoa. Ex.: "JOSE ... SANTOS COD 115" -> o "COD 115" so indica que a
# pessoa e do grupo 115; o numero ja e descartado por ser digito, e o rotulo
# textual (COD/GRUPO/REC) e ignorado aqui.
MARCADORES = {'COD', 'GRUPO', 'REC'}

# Sufixos geracionais: sao DISTINTIVOS. "JOSE SILVA" e "JOSE SILVA FILHO" sao
# pessoas diferentes (pai x filho). So casam se ambos tiverem o mesmo sufixo.
SUFIXOS_GERACIONAIS = {'FILHO', 'FILHA', 'NETO', 'NETA', 'JUNIOR', 'SOBRINHO', 'SOBRINHA', 'IRMAO', 'IRMA'}


def _tokens(nome):
    """Normaliza; remove conectores, numeros soltos e rotulos de grupo (COD/REC/GRUPO)."""
    return [t for t in normalizar(nome).split()
            if t not in CONECTORES and t not in MARCADORES and not t.isdigit()]


def _sufixo_geracional(tokens):
    """Devolve o sufixo geracional (FILHO/NETO/...) se for o ultimo token, senao None."""
    return tokens[-1] if tokens and tokens[-1] in SUFIXOS_GERACIONAIS else None


def _token_match(a, b, insercao_ok=False):
    """Dois tokens (palavras) representam o mesmo nome?

    Aceita: igualdade; inicial (1 letra); abreviatura curta (prefixo de ate
    3 letras, ex 'APAR'->'APARECIDA'); UM erro de digitacao (troca de 1 letra)
    em palavras de MESMO tamanho e 4+ letras (ex 'LUIZ'<->'LUIS').

    Por padrao NAO aceita insercao/remocao de letra (tamanhos diferentes), para
    nao ligar nomes distintos como 'SEVERINO'(8) x 'SEVERIANO'(9). Mas quando
    `insercao_ok` (ha uma inicial abreviada no nome -> e uma versao abreviada),
    aceita 1 letra a mais/menos: ex 'RAMPACHE' x 'RAMPCHE'.
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
    if len(a) >= 4 and len(b) >= 4:
        # 1 troca de letra em palavras de MESMO tamanho (Luiz x Luis)
        if len(a) == len(b) and levenshtein(a, b) == 1:
            return True
        # 1 letra a mais/menos, so quando ha abreviatura no nome (Rampache x Rampche)
        if insercao_ok and abs(len(a) - len(b)) == 1 and levenshtein(a, b) == 1:
            return True
    return False


def _todos_casam(menor, maior, insercao_ok=False):
    """Todo token da lista menor casa com um token DISTINTO da maior?
    Devolve (bool, n_casados)."""
    usados = [False] * len(maior)
    casados = 0
    for tm in menor:
        for i, tM in enumerate(maior):
            if not usados[i] and _token_match(tm, tM, insercao_ok):
                usados[i] = True
                casados += 1
                break
        else:
            return False, casados
    return True, casados


def _alinhar(menor, maior, insercao_ok=False):
    """Casa cada token de `menor` com um de `maior` (greedy, via _token_match).
    Retorna (faltantes, sobras): tokens de `menor` sem par e tokens de `maior`
    que sobraram (nao usados)."""
    usados = [False] * len(maior)
    faltantes = []
    for tm in menor:
        for i, tM in enumerate(maior):
            if not usados[i] and _token_match(tm, tM, insercao_ok):
                usados[i] = True
                break
        else:
            faltantes.append(tm)
    sobras = [maior[i] for i in range(len(maior)) if not usados[i]]
    return faltantes, sobras


def _relacionado(a, b):
    """a e b sao a mesma raiz de nome (abreviacao/truncamento)? Ex.: APAR ~ APARECIDA.

    So considera "parente" quando uma palavra e PREFIXO PROPRIO da outra (uma e
    o comeco da outra). Compartilhar apenas um prefixo comum NAO basta:
    'SEVERINO' e 'SEVERIANO' compartilham 'SEVERI' mas divergem depois -> sao
    nomes/pessoas diferentes, nao abreviacao.

    Usado so para graduar a DUVIDA (abreviacao -> revisar; nome de fato
    diferente -> outra pessoa).
    """
    return a == b or a.startswith(b) or b.startswith(a)


def _relacionado_a_algum(t, tokens):
    return any(_relacionado(t, o) for o in tokens)


def _score(tp, tf, insercao_ok=False):
    """Proporcao de tokens que casam (0..1), para ordenar candidatos."""
    menor, maior = (tp, tf) if len(tp) <= len(tf) else (tf, tp)
    _, casados = _todos_casam(menor, maior, insercao_ok)
    return casados / max(len(maior), 1)


def comparar_nomes(nome_planilha, nome_pdf):
    """Compara dois nomes (strings). Retorna (veredicto, score).
    Atalho para comparar_tokens; para lacos grandes, pre-calcule os tokens uma
    vez (ver `tokens`) e chame comparar_tokens diretamente (bem mais rapido)."""
    return comparar_tokens(_tokens(nome_planilha), _tokens(nome_pdf))


# Nome publico para pre-calcular tokens fora deste modulo (otimizacao).
tokens = _tokens


def comparar_tokens(tp, tf):
    """Como comparar_nomes, mas recebe as listas de tokens JA calculadas.
    Retorna (veredicto, score) com veredicto em
    {"IGUAL", "ABREVIACAO", "DUVIDA", "DIFERENTE"} e score float [0..1]."""
    if not tp or not tf:
        return ("DIFERENTE", 0.0)

    # Nomes identicos (apos normalizacao/remocao de conectores) -> ENCONTRADO.
    if tp == tf:
        return ("IGUAL", 1.0)

    # Ha uma inicial abreviada (token de 1 letra)? Entao o nome e uma versao
    # abreviada -> toleramos 1 letra a mais/menos nos outros tokens (Rampache x
    # Rampche). Sem abreviatura, mantemos estrito (Severino != Severiano).
    insercao_ok = any(len(t) == 1 for t in tp) or any(len(t) == 1 for t in tf)

    # Rejeicoes BARATAS antes de calcular o score (score varre tokens e custa
    # mais). Sufixo geracional (FILHO/NETO/...) e distintivo: pai x filho.
    if _sufixo_geracional(tp) != _sufixo_geracional(tf):
        return ("DIFERENTE", 0.0)

    # Primeiro nome diferente -> pessoa diferente.
    if not _token_match(tp[0], tf[0], insercao_ok):
        return ("DIFERENTE", 0.0)

    score = _score(tp, tf, insercao_ok)
    ultimo_ok = _token_match(tp[-1], tf[-1], insercao_ok)

    menor, maior = (tp, tf) if len(tp) <= len(tf) else (tf, tp)

    if ultimo_ok:
        # Mesmo primeiro nome e mesmo ultimo sobrenome, mas NAO identicos.
        faltantes, sobras = _alinhar(menor, maior, insercao_ok)
        if faltantes:
            # Nome(s) do meio do lado menor sem par:
            #   - abreviatura/truncamento (compartilha prefixo) -> DUVIDA (revisar)
            #     ex.: MARIA APAR SILVA x MARIA APARECIDA SILVA
            #   - nome COMPLETAMENTE diferente -> outra pessoa
            #     ex.: LUIZ PEDRO DA SILVA x LUIZ PEREIRA DA SILVA
            if all(_relacionado_a_algum(t, maior) for t in faltantes):
                return ("DUVIDA", score)
            return ("DIFERENTE", score)
        # Todo o lado menor casou. Olha os tokens EXTRAS do lado maior:
        #   - so iniciais/curtos (<=3 letras) -> abreviatura -> mesma pessoa
        #     ex.: MARIA A SILVA x MARIA APARECIDA SILVA
        #   - um NOME INTEIRO (>=4) presente so de um lado -> ambiguo, revisar
        #     ex.: JOSE SEVERINO DA SILVA x JOSE SEVERINO SALES DA SILVA
        if any(len(t) >= 4 for t in sobras):
            return ("DUVIDA", score)
        return ("ABREVIACAO", score)

    # Ultimo sobrenome difere.
    # Caso "sobrenome extra no fim": o ultimo do menor casa com o penultimo do
    # maior e todo o menor esta contido no maior. Ex.: JOSE PRAZERES x
    # JOSE PRAZERES VIEIRA -> pode ser parente/homonimo -> DUVIDA.
    if len(maior) > len(menor) and _token_match(menor[-1], maior[-2], insercao_ok):
        casam, _ = _todos_casam(menor, maior, insercao_ok)
        if casam:
            return ("DUVIDA", score)

    # Sobrenomes finais realmente diferentes -> pessoas diferentes.
    return ("DIFERENTE", score)
