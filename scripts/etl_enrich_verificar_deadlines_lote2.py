# -*- coding: utf-8 -*-
"""
Data enrichment em LOTE (2a passada): segunda tentativa de Deep Research,
por angulos diferentes da primeira, para as 3 linhas que continuaram com
"Deadline Aplicação" = "Verificar" apos o lote 1
(etl_enrich_verificar_deadlines_lote1.py).

Pedido da usuaria: "tente mais uma vez achar a data de aplicacao de todos
os programas que ainda estao em verificar".

Resultado:
  - Southern Connecticut State University (SCSU): CONFIRMADO. A primeira
    passada so tinha checado paginas de marketing/admissoes gerais; desta
    vez o subagente encontrou o portal oficial de aplicacao
    (apply.southernct.edu), pagina dedicada ao programa "Business
    Analytics" MS, com uma secao explicita "Terms and Deadlines": Fall
    17/ago, Spring 3/jan (datas fixas anuais, recorrentes -- nao rolling
    como a primeira passada tinha concluido). Alta confianca.
  - Texas A&M University-Texarkana: prazo GERAL de aplicacao continua
    "program specific" sem data publicada (confirmado de novo). Mas desta
    vez o subagente conseguiu ler o calendario academico oficial Fall
    2026 (PDF, que na 1a passada nao tinha sido encontrado) e confirmou o
    1o dia de aula: 24/ago/2026. Aplicando a regra ja confirmada na 1a
    passada (documentos de internacionais ate 45 dias antes do 1o dia de
    aula), chega-se a uma data CALCULADA (nao publicada diretamente como
    deadline, mas derivada de 2 fatos oficiais confirmados): ~10/jul/2026.
    Registrado como estimativa, nao como deadline oficial confirmado.
  - UCF (Professional), trilha Business Analytics do Professional MSM:
    NAO encontrado de novo, mesmo com angulos diferentes (catalogo
    oficial, portal do MSM, Executive Development Center, variacoes de
    busca). Confirma que a lacuna e real (a trilha Human Resources do
    mesmo MSM tem calendario proprio publicado, o que reforca que a
    ausencia de dado para Business Analytics nao e falha de busca).
    Permanece "Verificar" -- nao alterado por este script alem da nota.

Fontes:
- SCSU: https://apply.southernct.edu/portal/programs (portal oficial de
  aplicacao, pagina do programa "Business Analytics" MS -- renderizada via
  JS, exige navegador, nao so fetch)
- Texas A&M-Texarkana: calendario academico oficial Fall 2026
  (tamut.edu/enrollment-services/registrar/academic-calendar/files/fall-2026-full-term.pdf)
  + regra de 45 dias ja confirmada no lote 1 (catalog.tamut.edu)
- UCF (Professional): sem fonte nova (ver nota acima)

Uso: python etl_enrich_verificar_deadlines_lote2.py (executar a partir da
raiz do repositorio, mesma convencao dos demais etl_enrich_*.py).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

# Cada entrada: (Nome da Universidade, Estado, Cidade) -> {campo: novo_valor}
SCHOOL_UPDATES = {
    ("Southern Connecticut State University (SCSU)", "CT", "New Haven"): {
        "Deadline Aplicação": "Fall: 17/ago; Spring: 3/jan",
        "Rounds Aplicação": (
            "Datas fixas anuais (não rolling -- correção da pesquisa "
            "anterior). Confirmado no portal oficial de aplicação "
            "(apply.southernct.edu), página dedicada ao programa 'Business "
            "Analytics' MS, seção 'Terms and Deadlines': Fall 17/ago; "
            "Spring 3/jan. Datas recorrentes anualmente (sem ano específico "
            "exibido na página). Não é o prazo geral de 1/mai encontrado "
            "antes em busca (esse pertence a outro programa/orientação "
            "geral). Sem distinção de prazo internacional x doméstico "
            "publicada nesta página específica."
        ),
    },
    ("Texas A&M University-Texarkana (A&M-Texarkana)", "TX", "Texarkana"): {
        "Deadline Aplicação": (
            "Internacional (documentos): ~10/jul/2026 (calculado -- ver "
            "nota); prazo geral de aplicação não publicado"
        ),
        "Rounds Aplicação": (
            "Verificar (prazo geral) -- catálogo oficial confirma, mesmo em "
            "segunda pesquisa aprofundada, que os prazos são 'program "
            "specific' sem data listada para este programa. Único dado novo: "
            "calendário acadêmico oficial Fall 2026 localizado e lido -- 1º "
            "dia de aula confirmado em 24/ago/2026. Aplicando a regra já "
            "confirmada (documentos de candidatos internacionais devem "
            "chegar ao Office of International Studies até 45 dias antes do "
            "1º dia de aula), o resultado é uma data CALCULADA de "
            "~10/jul/2026 -- não é um deadline oficial publicado "
            "diretamente, é derivado de 2 fatos oficiais confirmados. "
            "Recomenda-se contato direto: internationalprograms@tamut.edu "
            "ou graduateadmissions@tamut.edu para confirmar o prazo geral "
            "de aplicação (não só o de documentos)."
        ),
    },
    ("UCF (Professional)", "FL", "Orlando"): {
        "Rounds Aplicação": (
            "Verificar -- não foi encontrado prazo específico publicado "
            "para a trilha 'Professional' (Professional MSM, trilha "
            "Business Analytics), distinta do MSBA padrão da UCF (que tem "
            "prazo próprio já registrado nesta base, na linha 'UCF'). "
            "Página oficial (business.ucf.edu) diz apenas 'Application "
            "deadlines vary by track -- contact Office of Professional "
            "Development (407-823-2185) or Graduate Business Programs "
            "(407-235-3917)'. SEGUNDA pesquisa aprofundada (catálogo "
            "oficial, portal do MSM, Executive Development Center, "
            "variações de busca) confirmou a mesma ausência de dado -- a "
            "trilha Human Resources do mesmo Professional MSM TEM "
            "calendário próprio publicado (early 1/mar, final 1/jul, "
            "admite só em anos ímpares), o que reforça que a lacuna para a "
            "trilha Business Analytics é real, não falha de busca. "
            "Recomenda-se contato direto antes de aplicar."
        ),
    },
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    all_fields = set()
    for updates in SCHOOL_UPDATES.values():
        all_fields.update(updates.keys())
    unknown_cols = [h for h in all_fields if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    no_fill = PatternFill(fill_type=None)

    report = []
    not_found = []

    for (nome, estado, cidade), updates in SCHOOL_UPDATES.items():
        target_rows = [
            row[0].row
            for row in ws.iter_rows(min_row=2)
            if row[col_uni - 1].value == nome
            and row[col_estado - 1].value == estado
            and row[col_cidade - 1].value == cidade
        ]
        if len(target_rows) == 0:
            not_found.append(nome)
            continue
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha corresponde a '{nome}' (linhas {target_rows}). Abortando por seguranca.")

        row_idx = target_rows[0]
        before = {}
        after = {}
        for header, new_value in updates.items():
            col_idx = header_to_col[header]
            cell = ws.cell(row=row_idx, column=col_idx)
            before[header] = cell.value
            cell.value = new_value
            cell.fill = no_fill
            after[header] = new_value
        report.append((nome, row_idx, before, after))

    if not_found:
        raise ValueError(f"Universidades nao encontradas na planilha (abortando, nada foi salvo): {not_found}")

    wb.save(DB_PATH)

    print(f"\n{len(report)} faculdades atualizadas:\n")
    for nome, row_idx, before, after in report:
        print(f"=== {nome} (linha {row_idx}) ===")
        for header in after:
            print(f"  - {header}:")
            print(f"      Antes: {before[header]!r}")
            print(f"      Depois: {after[header]!r}")
        print()


if __name__ == "__main__":
    main()
