import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from config_llm import GEMINI_API_KEY as _KEY
except (ImportError, AttributeError):
    _KEY = os.getenv("GEMINI_API_KEY", "")

_CLIENTE = None

def _iniciar():
    global _CLIENTE
    if _CLIENTE is not None:
        return True
    if not _KEY:
        return False
    try:
        from google import genai
        _CLIENTE = genai.Client(api_key=_KEY)
        return True
    except Exception:
        return False

def verificar_com_llm(nome1, nome2, timeout=10):
    if not _iniciar():
        return None

    prompt = (
        f"Você é um assistente rigoroso na conferência de trabalhadores brasileiros.\n"
        f"Nome A: {nome1}\n"
        f"Nome B: {nome2}\n\n"
        f"Sua tarefa é decidir se é a MESMA pessoa. Regras vitais:\n"
        f"1. Abreviaturas de nomes do meio são comuns e aceitáveis (ex: MARIA A SILVA x MARIA APARECIDA SILVA).\n"
        f"2. Conectores (DE, DA, DO, DOS, DAS) ausentes são ignoráveis.\n"
        f"3. ATENÇÃO MÁXIMA: Nomes com o último sobrenome principal DIFERENTE (ex: SILVA vs SANTOS) ou FALTANDO COMPLETAMENTE (ex: 'JOSE PRAZERES' vs 'JOSE PRAZERES VIEIRA') indicam muito provavelmente que NÃO é a mesma pessoa.\n"
        f"4. Se houver divergência forte em sobrenomes (um inteiro a mais que não seja mera abreviatura), pode ser homônimo ou parente.\n\n"
        f"Responda APENAS com UMA destas palavras:\n"
        f"SIM - Certeza absoluta que é a mesma pessoa.\n"
        f"NAO - Nomes têm diferenças relevantes de sobrenome, indicando pessoas diferentes.\n"
        f"DUVIDA - Pode ser a mesma pessoa, mas a divergência nominal exige revisão humana."
    )

    try:
        resp = _CLIENTE.models.generate_content(
            model="gemini-3.1-flash-lite",
            contents=prompt,
            config={"max_output_tokens": 10, "temperature": 0.0},
        )
        texto = resp.text.strip().upper()
        if "SIM" in texto:
            return True
        elif "DUVIDA" in texto:
            return "DUVIDA"
        else:
            return False
    except Exception:
        return None
