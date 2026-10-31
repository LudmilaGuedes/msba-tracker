# -*- coding: utf-8 -*-
"""
Data enrichment em LOTE: atualiza multiplas linhas classificadas pela
usuaria como "Pesquisar mais" ou "Falta informacao" no dataset exportado
(msba-meus-dados-*.xlsx), aplicando o mesmo processo de Deep Research ja
usado nas faculdades individuais desta sessao -- mas em escala (55
faculdades pesquisadas em 6 lotes paralelos via subagentes).

Este script contem APENAS os achados de ALTA CONFIANCA: fonte oficial
(.edu) verificada por fetch direto (nao so resumo de busca), sem conflito
nao resolvido com dado ja cadastrado, e sem risco de mistura de programas
(MBA vs MS, campus principal vs satelite, presencial vs online).

Achados de confianca MODERADA (so resumo de busca, nao fetch direto
oficial) ou com CONFLITO nao resolvido foram DELIBERADAMENTE EXCLUIDOS
deste lote, para manter o mesmo rigor aplicado nas ~15 faculdades
pesquisadas individualmente antes nesta sessao. Exemplos descartados:
GWU (deadline/GMAT via busca agregada), UW-Madison (TA/RA institucional
generico, conflito de ranking QS nao resolvido), WashU (pre-requisitos/
GMAT medio via busca agregada, e um conflito real entre duas buscas sobre
qual numero e salario vs tuition -- NENHUM dos dois foi usado), UC Irvine
(PDF oficial ilegivel em 2 tentativas de fetch), U Michigan
(michiganross.umich.edu bloqueou fetch com 403 em 2 tentativas), SCSU
(GMAT waiver via busca agregada), American University (taxa de aceitacao
e salario via busca agregada, nao fetch direto).

Tambem NAO alterados: casos onde o dado encontrado apenas RECONFIRMA o
que ja estava cadastrado (Wake Forest Tamanho Turma/Taxa Empregabilidade,
Bentley Taxa Empregabilidade, Seton Hall GPA Minimo, UT Knoxville Salario,
Cornell tuition, RPI Salario) -- sem mudanca necessaria.

Fontes por escola (todas .edu, verificadas por fetch direto):
- CMU: cmu.edu/tepper/programs/master-business-analytics/careers/compensation-highlights
- U of Maryland: networth.rhsmith.umd.edu/welcome-ms-fall/financial-information/tuition-fees-business-masters-programs
- UConn: analytics.business.uconn.edu/overview/faq/
- ASU: wpcarey.asu.edu/masters-programs/business-analytics/class-profile
- Cal Poly Pomona: cpp.edu/cba/masters-business-analytics.shtml (2025 Graduate Status Report)
- RPI: lally.rpi.edu/programs/graduate-programs/ms-business-analytics-stem-designated
- Rutgers Camden: business.camden.rutgers.edu (FAQ oficial)
- George Mason: business.gmu.edu/programs/graduate-degree-programs/ms-business-analytics
- U Washington: foster.uw.edu/foster-degree/master-of-science-in-business-analytics-msba
- U Denver: daniels.du.edu/business-information-analytics/masters/
- Wayne State: ilitchbusiness.wayne.edu/tisa/msdsba/requirements; gradschool.wayne.edu/students/funding/scholarships
- American University (Kogod): kogod.american.edu/programs-admissions/masters/analytics/admissions
- Bentley: bentley.edu/academics/graduate-programs/masters-business-analytics/msba-overview
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

# Cada entrada: (Nome da Universidade, Estado, Cidade) -> {campo: novo_valor}
SCHOOL_UPDATES = {
    ("CMU", None, None): {
        "Taxa Empregabilidade %": 100,
        "Salário Médio Pós USD": 112625,
    },
    ("U of Maryland", None, None): {
        "Custo/Crédito (Residente) USD": 1631,
        "Tuition Total (Residente) USD": 48930,
    },
    ("UConn", None, None): {
        "Duração (meses)": "~9 (aproximado -- página oficial diz que tempo integral termina em 3 semestres Fall/Spring, sem número exato de meses divulgado)",
    },
    ("ASU", None, None): {
        "Taxa Empregabilidade %": 95,  # substitui 91 (sem fonte registrada) por valor oficial: wpcarey.asu.edu/masters-programs/business-analytics/class-profile
    },
    ("Cal Poly Pomona", None, None): {
        "Salário Médio Pós USD": 95375,
        "Taxa Empregabilidade %": 88,
    },
    ("Rensselaer Polytechnic Institute (RPI)", None, None): {
        "Duração (meses)": 12,
        "Total Créditos": 30,
        "GPA Mínimo": "3.0 (3.3+ preferencial/competitivo)",
        "GMAT/GRE Obrigatório": "Opcional",
        "Testes Inglês Aceitos": "TOEFL mín. 88 (iBT) ou 4.5 (nova escala 6.0); IELTS e Duolingo não confirmados com fonte oficial",
    },
    ("Rutgers University–Camden", None, None): {
        "Início": "Fall (set), Spring (jan), Summer (mai)",
    },
    ("George Mason University", None, None): {
        "Início": "Fall ou Spring",
        "Deadline Aplicação": "Fall (início em agosto): 1º de maio",
        "GMAT/GRE Obrigatório": "Test-optional; recomendado ≥ 50º percentil se enviado",
        "Principais Empregadores": "Citi Private Bank, Dell Technologies, Lockheed Martin, DoorDash, U.S. Department of Justice, Navy Federal",
        "Bolsas Disponíveis": "Sim -- avaliação automática durante revisão da candidatura, sem inscrição separada; valores não divulgados publicamente",
    },
    ("University of Washington", None, None): {
        "Início": "Junho (turma única por ano)",
    },
    ("University of Denver", None, None): {
        "Início": "Fall (meados de setembro) ou Spring (meados/fim de março)",
        "Taxa Empregabilidade %": 94,
    },
    ("Wayne State University", None, None): {
        "Matérias Pré-Requisito": "Conhecimento básico de probabilidade/estatística + pelo menos 1 curso de programação (C, C++, Java, Python, R ou FORTRAN)",
        "Documentos Necessários": "GRE/GMAT (scores), resume profissional, Statement of Purpose, 2 cartas de recomendação, teste de inglês (internacionais), diploma de bacharel",
        "Tipos de Bolsa": "AGRADE Scholarship -- cobre até 4 créditos na tarifa in-state por 2 semestres consecutivos (Fall+Winter ou Winter+Fall)",
    },
    ("American University", None, None): {
        "Rounds Aplicação": "Rolling (a partir de 1º/jul); R1: 15/nov; R2: 15/jan; R3: 15/mar; R4: 15/mai",
        "Taxa Aplicação USD": 100,
        "Documentos Necessários": "Resume, ensaio de propósito (até 750 palavras), 1 carta de recomendação, transcripts, GMAT/GRE opcional (ou ensaio de isenção), teste de inglês",
        "Matérias Pré-Requisito": "Nenhum -- não é exigida graduação em área relacionada",
    },
    ("Bentley University", None, None): {
        "Salário Médio Pós USD": 80000,
    },
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    # valida que todas as colunas referenciadas existem no schema
    all_fields = set()
    for updates in SCHOOL_UPDATES.values():
        all_fields.update(updates.keys())
    unknown_cols = [h for h in all_fields if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]

    no_fill = PatternFill(fill_type=None)

    report = []
    not_found = []

    for (nome, _estado, _cidade), updates in SCHOOL_UPDATES.items():
        target_rows = [
            row[0].row for row in ws.iter_rows(min_row=2) if row[col_uni - 1].value == nome
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
