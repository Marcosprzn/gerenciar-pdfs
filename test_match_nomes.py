"""Testes do matcher de nomes. Rode:  python test_match_nomes.py"""
from match_nomes import comparar_nomes

# (nome_planilha, nome_pdf, veredicto_esperado)
CASOS = [
    # --- identicos (IGUAL) ---
    ("Maria Aparecida Silva", "Maria Aparecida Silva", "IGUAL"),      # identico
    ("Maria de Souza",        "Maria Souza",             "IGUAL"),      # conector removido
    ("Ana Silva",             "Ana Silva",               "IGUAL"),      # nome curto identico
    ("Jose Carlos de Oliveira","Jose Carlos Oliveira",   "IGUAL"),      # conector

    # --- mesma pessoa via abreviacao/inicial/erro (ABREVIACAO) ---
    ("Maria Aparecida Silva", "Maria A Silva",          "ABREVIACAO"), # meio abreviado (inicial)
    ("Luiz Gonzaga Lima",     "Luis Gonzaga Lima",       "ABREVIACAO"), # 1 erro de digitacao Luiz/Luis
    ("Joao Paulo Santos",     "J P Santos",              "ABREVIACAO"), # iniciais
    # "COD 115" / rotulo de grupo nao faz parte do nome -> mesma pessoa
    ("Jose Carlos dos Santos", "Jose Carlos dos Santos COD 115", "IGUAL"),
    # Apelido entre parenteses nao faz parte do nome -> mesma pessoa
    ("Jose Severino Alves",    "Jose Severino Alves ( FI )",       "IGUAL"),
    ("Jose Severino da Silva", "Jose Severino da Silva ( CALADO )", "IGUAL"),

    # --- ambiguos, devem exigir revisao (DUVIDA) ---
    # Prefixo de 4+ letras no meio -> pode ser abreviatura -> revisao.
    ("Maria Aparecida Silva", "Maria Apar Silva",       "DUVIDA"),   # meio truncado (prefixo)
    ("Maria Aparecida Silva", "Maria Silva",            "DUVIDA"),   # nome do meio inteiro so de um lado
    ("Jose Severino da Silva","Jose Severino Sales da Silva","DUVIDA"),# sobrenome inteiro a mais no meio
    ("Jose Prazeres",         "Jose Prazeres Vieira",    "DUVIDA"),   # sobrenome extra no fim
    ("Antonio Carlos",        "Antonio Carlos de Jesus", "DUVIDA"),   # sobrenome extra no fim
    ("Maria Silva",           "Maria Silva Santos",      "DUVIDA"),   # sobrenome extra no fim
    ("Ana Paula Silva",       "Ana Paula Silva Souza",   "DUVIDA"),   # sobrenome extra

    # --- devem ser rejeitados (DIFERENTE) ---
    ("Ana Silva",             "Mariana Silva Santos",    "DIFERENTE"),# o BUG antigo (substring)
    ("Jose Carlos Silva",     "Jose Carlos Santos",      "DIFERENTE"),# ultimo sobrenome diferente
    ("Carlos Eduardo Lima",   "Carla Eduardo Lima",      "DIFERENTE"),# primeiro nome diferente (genero)
    ("Pedro Almeida",         "Paulo Almeida",           "DIFERENTE"),# primeiro nome diferente
    ("Ana Lima",              "Adriana Lima",            "DIFERENTE"),# substring no primeiro nome
    ("Joao Silva Costa",      "Joao Pereira Ramos",      "DIFERENTE"),# meio e fim diferentes
    # Meio COMPLETAMENTE diferente (primeiro+ultimo iguais) -> outra pessoa.
    # Casos reais reportados que antes davam falso "POSSIVEL ERRO NOMINAL":
    ("Maria Jose Silva",      "Maria Helena Silva",      "DIFERENTE"),# meio diferente
    ("Jose Carlos do Nascimento Santos", "Jose Martiliano dos Santos", "DIFERENTE"),
    ("Luiz Pedro da Silva",   "Luiz Pereira da Silva",   "DIFERENTE"),# PEDRO != PEREIRA
    ("Luiz Pereira da Silva", "Luiz Rodolfo da Silva",   "DIFERENTE"),# PEREIRA != RODOLFO
    ("Antonio Julio da Cunha Filho", "Antonio Anastacio da Silva Filho", "DIFERENTE"),
    ("Reginaldo Manoel da Silva", "Edinaldo Vieira da Silva", "DIFERENTE"),
    # SEVERINO x SEVERIANO: nomes distintos (insercao de letra) -> pessoas diferentes
    ("Jose Severino da Silva", "Jose Severiano da Silva", "DIFERENTE"),
    ("Jose Severino da Silva Filho", "Jose Severiano da Silva Filho", "DIFERENTE"),
    # Sufixo geracional (FILHO/NETO/...) e distintivo: pai x filho = diferentes.
    ("Severino Antonio da Silva", "Severino Antonio da Silva Filho", "DIFERENTE"),
    ("Severino Antonio da Silva Filho", "Severino Antonio da Silva Neto", "DIFERENTE"),
    ("Jose Francisco Barbosa Irmao", "Jose Francisco Barbosa", "DIFERENTE"),
    # ...mas se AMBOS tem o mesmo sufixo, seguem casando normalmente:
    ("Antonio Anastacio da Silva Filho", "Antonio Anastacio da Silva Filho", "IGUAL"),
    ("Joao Franco da Silva Filho", "J Franco da Silva Filho", "ABREVIACAO"),
    # Com abreviatura (A.=ALMEIDA) tolera 1 letra a mais/menos: RAMPACHE~RAMPCHE
    ("Jayme Milagres Melo de A. Rampache", "Jayme Milagres Melo de Almeida Rampche", "ABREVIACAO"),
]


def main():
    falhas = 0
    for planilha, pdf, esperado in CASOS:
        veredicto, score = comparar_nomes(planilha, pdf)
        ok = veredicto == esperado
        if not ok:
            falhas += 1
        marca = "ok " if ok else "FALHOU"
        print(f"[{marca}] {planilha!r:38} x {pdf!r:34} -> {veredicto:10} (score {score:.2f})"
              + ("" if ok else f"   ESPERADO {esperado}"))
    total = len(CASOS)
    print(f"\n{total - falhas}/{total} casos corretos.")
    if falhas:
        raise SystemExit(f"{falhas} FALHA(S)")
    print("TODOS OS CASOS PASSARAM")


if __name__ == "__main__":
    main()
