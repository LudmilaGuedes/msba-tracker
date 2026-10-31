# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): UAH (University of Alabama in Huntsville,
College of Business), Huntsville, AL -- MS in Business Analytics,
presencial, STEM.

Fontes (verificadas por WebFetch/WebSearch em set/2026):
- uah.edu/business/grad/graduate-tuition (calculadora oficial AY 2026-27)
- uah.edu/bursar/tuition/tuition-rates-2 (tarifas 2025-26, para conferir)
- uah.edu/business/grad/degrees/analytics/program-snapshot
- uah.edu/admissions/graduate/apply-for-admission/international-students
- uah.edu/admissions/graduate/apply-for-admission/test-score-requirements
- uah.edu/business/grad/graduate-assistantships
- uah.edu/admissions/graduate/financial-aid/scholarships
- uah.edu/admissions/graduate/apply-for-admission (isencao de taxa)

ACHADOS:

1) *** TUITION CORRIGIDO: a rodada anterior registrou US$1.730/credito
   "igual para residentes e nao-residentes". Isso esta ERRADO. Na
   calculadora 2026-27, campus/misto: US$1.730/credito para um perfil e
   US$736/credito para o outro, com teto fixo por semestre a partir de 9
   creditos (US$14.420 vs US$6.010). O texto extraido da pagina mostra os
   rotulos Resident/Non-Resident invertidos, mas a tabela oficial do
   bursar de 2025-26 resolve a duvida: nao-residente paga US$1.680/h e
   US$14.000 fixo, residente paga US$715/h e US$5.835 fixo. Os valores de
   2026-27 sao esses mesmos com ~3% de reajuste, entao NAO-RESIDENTE =
   US$1.730 e RESIDENTE = US$736.
   Total do programa: segue a propria calculadora (plano de 12 meses com
   12 + 12 + 6 creditos, Fall/Spring/Summer):
     Intl:      2 x 14.420 + 6 x 1.730 = US$39.220
     Residente: 2 x 6.010  + 6 x 736   = US$16.436
   Fees Adicionais (US$136/credito x 30 = US$4.080) foram re-confirmadas
   para 2026-27, sem mudanca. A International Fee (US$150/semestre) segue
   fora, como na convencao da rodada anterior. ***

2) BOLSAS: a GEM Fellowship que estava registrada e restrita a cidadaos e
   residentes permanentes dos EUA, entao nao serve para internacional e
   foi removida. O GA do College of Business existe (10h ou 20h/semana),
   mas a pagina diz "limited and highly competitive". O GA cobre ate 12h
   de tuition calculada pela tarifa IN-STATE, mais stipend e fee
   scholarship. Nao esta claro se a diferenca da tarifa de nao-residente
   tambem e coberta. Prazos de prioridade do GA: 31/out (Spring) e 31/mar
   (Fall/Summer), com aplicacao separada via formulario. Bolsas
   competitivas da universidade: aplicacao ate 15/dez, sem detalhes de
   elegibilidade para internacionais. Realismo rebaixado de "Moderada"
   para "Baixa-Moderada".

3) INICIO: Fall e Summer (antes so Fall). FORMATO: presencial noturno,
   com algumas opcoes online.

4) PERFIL DA TURMA (2018, unico disponivel): media de 5,8 anos de
   experiencia e 30 anos de idade. O programa atende profissionais que ja
   trabalham e recem-formados. Anotado em Experiencia Profissional.

5) GMAT WAIVER: faltava a opcao "GPA 2.8 + 5 anos de experiencia".

6) TAXA DE APLICACAO: US$60, confirmado. So ha isencao para alunos JUMP e
   funcionarios da UAH, e desconto para ex-alunos da UAH. A UAH diz que
   nao tem outro processo de isencao.

Nao encontrado (continua "Verificar"): taxa de empregabilidade, salario
medio, tamanho da turma, taxa de aceitacao, entrevista.

Uso: python scripts/etl_enrich_uah_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UAH",
    "Estado": "AL",
    "Cidade": "Huntsville",
}

UPDATES = {
    "Formato": "Presencial (noturno; algumas disciplinas com opção online)",
    "Início": "Fall, Summer",
    "Custo/Crédito (Intl) USD": 1730,
    "Custo/Crédito (Residente) USD": 736,
    "Tuition Total (Intl) USD": 39220,
    "Tuition Total (Residente) USD": 16436,
    "Tipos de Bolsa": (
        "Graduate Assistantship do College of Business (part-time 10h/sem ou full-time "
        "20h/sem: até 12h de tuition pela tarifa in-state + stipend + fee scholarship + "
        "seguro saúde; aplicação separada via formulário); Competitive Scholarships da "
        "UAH (aplicação até 15/dez; sem detalhes publicados)"
    ),
    "Quão Realista (Intl)": (
        "Baixa-Moderada -- GAs do College of Business são 'limited and highly "
        "competitive' e a cobertura de tuition é pela tarifa in-state (não está claro se "
        "cobre a diferença de não-residente); a GEM Fellowship é exclusiva para "
        "cidadãos/residentes dos EUA"
    ),
    "Caminho Tuition Residente": (
        "Não identificado -- internacionais pagam tarifa de não-residente (US$1.730/"
        "crédito ou US$14.420 fixo por semestre com 9+ créditos, 2026-27); GA cobre "
        "tuition com base na tarifa in-state"
    ),
    "GMAT/GRE Obrigatório": (
        "GMAT (mín. 500) ou GRE (mín. 300 verbal+quant, 3.0 analytical writing); waiver "
        "com GPA 3.0 + 3 anos de experiência, GPA 2.8 + 5 anos de experiência, ou "
        "pós-graduação já concluída"
    ),
    "Experiência Profissional": (
        "Não obrigatória -- mas a turma de 2018 (único perfil publicado) tinha média de "
        "5,8 anos de experiência e 30 anos de idade; o programa atende tanto "
        "profissionais quanto recém-formados"
    ),
    "Deadline Aplicação": (
        "Doméstico: Fall 1º jun, Spring 30 nov, Summer 15 mar; Internacional (I-20): "
        "Fall 1º ago, Spring 1º dez. Prioridade para GA: 31/mar (Fall/Summer) e 31/out "
        "(Spring); bolsas competitivas: 15/dez"
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados; admissão por deadline de semestre (ver Deadline "
        "Aplicação). Sem isenção de taxa para candidatos externos (só para alunos "
        "JUMP e funcionários da UAH; ex-alunos UAH pagam US$35)"
    ),
    "Documentos Necessários": (
        "Histórico acadêmico (oficial ou não-oficial via Charger Status) + avaliação "
        "course-by-course se estudou fora dos EUA; resume; GMAT/GRE (se não dispensado); "
        "teste de inglês; affidavit of financial support com comprovação de fundos para "
        "1 ano de tuition + custo de vida (internacionais)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    unknown_cols = [h for h in UPDATES if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    target_rows = []
    for row in ws.iter_rows(min_row=2):
        if (
            row[col_uni - 1].value == MATCH["Nome da Universidade"]
            and row[col_estado - 1].value == MATCH["Estado"]
            and row[col_cidade - 1].value == MATCH["Cidade"]
        ):
            target_rows.append(row[0].row)

    if len(target_rows) == 0:
        raise ValueError("Nenhuma linha encontrada para UAH (Huntsville, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UAH, Huntsville, AL)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (UAH):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
