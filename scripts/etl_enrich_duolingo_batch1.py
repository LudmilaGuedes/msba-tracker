"""
Data enrichment (lote único): atualiza 'Testes Inglês Aceitos' em
MÚLTIPLAS linhas do msba_programs_database.xlsx, a pedido do usuário:
confirmar quais faculdades aceitam o Duolingo English Test para
admissão de pós-graduação (e o score mínimo), cobrindo as escolas que
já tinham o campo preenchido mas SEM menção a Duolingo.

METODOLOGIA: pesquisa via WebSearch/WebFetch em vários lotes paralelos
(2026-09-09). Padrão confirmado repetidamente nesta base: a página
ESPECÍFICA do programa de MS Business Analytics frequentemente só
lista TOEFL/IELTS e omite Duolingo, mas a página GERAL/university-wide
de graduate admissions ("English Language Proficiency Requirements")
com frequência aceita Duolingo mesmo assim -- por isso a pesquisa
priorizou essa página geral quando a do programa não tinha a
informação.

LIMITAÇÃO IMPORTANTE DESTA RODADA: o orçamento de WebSearch da sessão
(200 buscas) se esgotou no meio da varredura. Das 45 faculdades que
precisavam ser checadas, 22 foram CONFIRMADAS (algumas com alta
confiança via WebFetch direto na página oficial, outras via busca
antes do limite bater) e 23 ficaram sem confirmação por falta de
orçamento de busca -- essas continuam com o campo como estava antes
(sem menção a Duolingo), NÃO foram marcadas como "Duolingo não
aceito" só por falta de dado (distinção importante: "não encontrado"
!= "não aceito"). A única exceção são Georgia Tech e University of
Maryland, onde a fonte oficial confirma EXPLICITAMENTE que Duolingo
NÃO é aceito (não é uma omissão, é uma exclusão documentada).

Esta rodada é PARCIAL por desenho -- as 23 faculdades restantes
precisam de uma nova sessão (ou orçamento de busca ampliado) para
serem confirmadas: UGA, Virginia Tech, VT Global BA, Georgetown, GWU,
Dakota State, UC Berkeley, UC Irvine, U of Arizona, SMU, U of
Michigan, Northwestern, Texas Tech, Illinois Tech, University of
North Texas (UNT), Rensselaer Polytechnic Institute (RPI), Southern
Connecticut State University (SCSU), Wake Forest University, William
& Mary, Oakland University, Florida Institute of Technology, Texas
A&M University-Texarkana, West Texas A&M University (WTAMU).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

# Cada item: (Nome da Universidade, Estado, Cidade) -> novo valor completo de 'Testes Inglês Aceitos'
SCHOOL_UPDATES = [
    (
        {"Nome da Universidade": "USC", "Estado": "CA", "Cidade": "Los Angeles"},
        {
            "Testes Inglês Aceitos": (
                "TOEFL 100 total (mín. 20 por seção) / IELTS 7.0 (mín. "
                "6 por seção) / PTE aceito / Duolingo aceito mas SEM "
                "mínimo formal de admissão (usado só para 'placement' "
                "-- decidir se o aluno precisa fazer o exame ISE de "
                "nivelamento ao chegar)"
            )
        },
    ),
    (
        {"Nome da Universidade": "UT Austin", "Estado": "TX", "Cidade": "Austin"},
        {"Testes Inglês Aceitos": "TOEFL 100 / IELTS 7.0 / Duolingo aceito como alternativa (mínimo específico não confirmado)"},
    ),
    (
        {"Nome da Universidade": "UC Davis", "Estado": "CA", "Cidade": "Davis/San Francisco"},
        {
            "Testes Inglês Aceitos": (
                "TOEFL 100 (iBT) mínimo / IELTS 7.0 mínimo / Duolingo "
                "mín. 115 (quem submete só Duolingo pode precisar "
                "fazer o TOEP -- Test of Oral English Proficiency -- "
                "da UC Davis se admitido; abaixo de 130 pode exigir "
                "curso de inglês adicional já matriculado)"
            )
        },
    ),
    (
        {"Nome da Universidade": "U of Minnesota", "Estado": "MN", "Cidade": "Minneapolis"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/PTE/Duolingo (mín. 115) -- obrigatório se o candidato não for falante nativo de inglês ou não tiver se formado em instituição de ensino em inglês"},
    ),
    (
        {"Nome da Universidade": "Texas A&M", "Estado": "TX", "Cidade": "Houston"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo (mín. 100)"},
    ),
    (
        {"Nome da Universidade": "U of Utah", "Estado": "UT", "Cidade": "Salt Lake City"},
        {"Testes Inglês Aceitos": "TOEFL iBT competitivo: 90 / IELTS Academic obrigatório: 7.0 (de 9) / Duolingo mín. 110 (university-wide; alguns departamentos exigem até 130)"},
    ),
    (
        {"Nome da Universidade": "Michigan State", "Estado": "MI", "Cidade": "East Lansing"},
        {"Testes Inglês Aceitos": "TOEFL mín. 100 / IELTS mín. 7.0 / Duolingo mín. 110 (faixa de admissão provisória: 95-110)"},
    ),
    (
        {"Nome da Universidade": "U of Florida", "Estado": "FL", "Cidade": "Gainesville"},
        {
            "Testes Inglês Aceitos": (
                "TOEFL 80 (iBT) / IELTS 6.0 mínimo (MELAB 77 também "
                "aceito) / Duolingo aceito, mas score mínimo NÃO "
                "confirmado com confiança direto na fonte oficial da "
                "UF (indício cruzado de documento da FAU aponta 140, "
                "não confirmado oficialmente -- tratar com cautela)"
            )
        },
    ),
    (
        {"Nome da Universidade": "UT Knoxville", "Estado": "TN", "Cidade": "Knoxville"},
        {"Testes Inglês Aceitos": "TOEFL 80/IELTS 6.5/Duolingo mín. 120"},
    ),
    (
        {"Nome da Universidade": "U at Buffalo SUNY", "Estado": "NY", "Cidade": "Buffalo"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo mín. 120 (university-wide; pode ser maior no departamento específico)"},
    ),
    (
        {"Nome da Universidade": "Stony Brook SUNY", "Estado": "NY", "Cidade": "Long Island"},
        {"Testes Inglês Aceitos": "TOEFL 80/IELTS 6.5/Duolingo mín. 105 (fonte não-oficial, confirmar antes de aplicar)"},
    ),
    (
        {"Nome da Universidade": "U of Maryland", "Estado": "MD", "Cidade": "College Park"},
        {
            "Testes Inglês Aceitos": (
                "TOEFL iBT mín. 96; IELTS Academic mín. 7.0; PTE "
                "Academic mín. 68; dispensado para diploma de "
                "instituição de língua inglesa qualificada -- "
                "Duolingo NÃO É ACEITO pela Graduate School (confirmado "
                "oficialmente; a UMD aceita Duolingo para undergrad, "
                "mas não para pós-graduação)"
            )
        },
    ),
    (
        {"Nome da Universidade": "U of Tulsa", "Estado": "OK", "Cidade": "Tulsa"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo mín. 105"},
    ),
    (
        {"Nome da Universidade": "SF State", "Estado": "CA", "Cidade": "San Francisco"},
        {"Testes Inglês Aceitos": "TOEFL mín. 88 (570 PBT) / IELTS mín. 6.5 / PTE mín. 65 / Duolingo mín. 120"},
    ),
    (
        {"Nome da Universidade": "Cal Poly Pomona", "Estado": "CA", "Cidade": "Pomona"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS; Duolingo mín. 105 (alguns programas exigem 120); CPP English Language Institute (CPELI) nível 6 também aceito"},
    ),
    (
        {"Nome da Universidade": "Baylor", "Estado": "TX", "Cidade": "Waco"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/PTE/Duolingo mín. 125 (dispensado com diploma de instituição dos EUA credenciada ou país de língua oficial inglesa; validade de 2 anos)"},
    ),
    (
        {"Nome da Universidade": "Santa Clara University", "Estado": "CA", "Cidade": "Santa Clara"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo (mín. 105-110 para Engenharia, 110 geral pelo Graduate Admissions Office) exigido -- dispensado com diploma de país anglófono (EUA, UK, Canadá anglófono, Irlanda, NZ, Austrália) obtido nos últimos 4 anos"},
    ),
    (
        {"Nome da Universidade": "Loyola Marymount University (LMU)", "Estado": "CA", "Cidade": "Los Angeles"},
        {"Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo (mín. 130) exigido"},
    ),
    (
        {"Nome da Universidade": "University of Illinois Chicago (UIC)", "Estado": "IL", "Cidade": "Chicago"},
        {
            "Testes Inglês Aceitos": (
                "IELTS 6.5 / TOEFL 80 / PTE 54 (médias de admitidos -- "
                "mínimos oficiais ainda não confirmados) / Duolingo "
                "mín. 130 (piso institucional para pós-graduação) -- "
                "não vale para quem for atuar como teaching assistant, "
                "que precisa de TOEFL/IELTS ou English Proficiency "
                "Interview nesse caso"
            )
        },
    ),
    (
        {"Nome da Universidade": "Georgia Tech", "Estado": "GA", "Cidade": "Atlanta"},
        {
            "Testes Inglês Aceitos": (
                "TOEFL iBT mín. 100 (mín. 19 por seção) / IELTS "
                "Academic mín. 7.5 (mín. 6.5 em Reading/Listening/"
                "Speaking, mín. 5.5 em Writing) / GT Language "
                "Institute Test; dispensado para graduação cursada "
                "integralmente em instituição de ensino em inglês -- "
                "Duolingo NÃO É ACEITO (listado explicitamente na "
                "página oficial em 'Tests Not Accepted', junto com "
                "TOEFL Essentials, IELTS General Training e PTE)"
            )
        },
    ),
    (
        {"Nome da Universidade": "U of South Carolina", "Estado": "SC", "Cidade": "Columbia"},
        {"Testes Inglês Aceitos": "TOEFL mín. 100 (internet-based) / IELTS mín. 7.5 / TOEIC / PTE Academic / Duolingo mín. 120; dispensável para quem se formou em programa ministrado em inglês"},
    ),
    (
        {"Nome da Universidade": "U of Houston", "Estado": "TX", "Cidade": "Houston"},
        {"Testes Inglês Aceitos": "TOEFL 100 (iBT) / IELTS 6.5 / Duolingo mín. 105"},
    ),
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for match, updates in SCHOOL_UPDATES:
        unknown_cols = [h for h in updates if h not in header_to_col]
        if unknown_cols:
            raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

        col_uni = header_to_col["Nome da Universidade"]
        col_estado = header_to_col["Estado"]
        col_cidade = header_to_col["Cidade"]

        target_rows = []
        for row in ws.iter_rows(min_row=2):
            if (
                row[col_uni - 1].value == match["Nome da Universidade"]
                and row[col_estado - 1].value == match["Estado"]
                and row[col_cidade - 1].value == match["Cidade"]
            ):
                target_rows.append(row[0].row)

        if len(target_rows) == 0:
            raise ValueError(f"Nenhuma linha encontrada para {match}. Abortando.")
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha corresponde a {match} (linhas {target_rows}). Abortando por seguranca.")

        row_idx = target_rows[0]
        print(f"\n{match['Nome da Universidade']} (linha {row_idx}):")

        for header, new_value in updates.items():
            col_idx = header_to_col[header]
            cell = ws.cell(row=row_idx, column=col_idx)
            before = cell.value
            cell.value = new_value
            cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
            print(f"  - {header}: {before!r} -> {new_value!r}")

    wb.save(DB_PATH)
    print("\nSalvo com sucesso.")


if __name__ == "__main__":
    main()
