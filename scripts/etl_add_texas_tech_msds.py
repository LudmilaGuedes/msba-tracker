# -*- coding: utf-8 -*-
"""
Adiciona a Texas Tech (Rawls College of Business, Lubbock, TX) com o
Master of Science in Data Science (MSDS).

Contexto: a Texas Tech foi removida da base em set/2026
(etl_remove_no_msba_program.py) porque a linha estava rotulada como
"MSBA", e o Rawls nao tem MSBA. Em 27/set/2026 foi pedido para
reincluir a faculdade com o programa correto, depois de confirmado que o
MSDS e voltado ao mercado de trabalho: sem tese, com projeto final
(capstone), dentro da escola de negocios e STEM. O nome oficial do
programa deixa claro que nao e um MSBA.

Fontes (set/2026):
- depts.ttu.edu/rawlsbusiness/graduate/ms/ (lista de MS; nenhum MSBA)
- depts.ttu.edu/rawlsbusiness/graduate/ms/datascience/ e /index.php
  (36 creditos, 1 ou 2 anos, presencial ou online, lock-step, STEM,
  capstone "Data Science Project", requisitos, prazos Spring 1/dez e
  Summer 1/mai)
- depts.ttu.edu/gradschool/admissions/InternationalProspectiveStudents.php
  (ingles: TOEFL 79 / Essentials 8.5 / IELTS 6.5 / Duolingo 100 / PTE 60 /
  CPE 180 / CAE 175; taxa US$75, que sobe para US$100 depois de certas
  datas; prazo recomendado para internacionais: 15/jan para Fall e Summer)
- depts.ttu.edu/gradschool/financial-support/tuition.php (custo estimado
  AY 2026-27 para programas de BUSINESS, com 18 creditos/ano letivo:
  tuition & fees de US$13.836 para residente e US$21.216 para
  nao-residente [38.996 - 14.534 - 3.246]; living expenses US$14.534;
  seguro US$3.246)

CALCULOS:
- Custo/credito, ja com fees e com o diferencial de business:
    residente     13.836 / 18 = 768,67
    nao-residente 21.216 / 18 = 1.178,67
- Tuition total (36 creditos): residente 27.672; intl 42.432. E uma
  estimativa proporcional; agregadores citam ~US$39.8k para intl, na mesma
  ordem de grandeza.
- Fees Adicionais: ja incluidas nos valores acima (a fonte oficial junta
  tuition e fees).
- Custo Vida Total: 14.534 / 9 x 12 = 19.378,67 (ano letivo de 9 meses
  pro-rateado para o programa de 12 meses; mesmo metodo da UAH).
- Aluguel/mes US$720: pesquisa de mercado de aluguel estudantil em Lubbock
  da rodada anterior (etl_enrich_texas_tech.py).

INCERTEZAS REGISTRADAS:
- GMAT/GRE: as fontes conflitam. A versao anterior da base dizia
  "opcional"; snippets de busca dizem "exigido, sem nota minima, com
  dispensa para 3-5 anos de experiencia", mas podem ser da pagina do STEM
  MBA. Nao foi possivel ler o FAQ do MSDS (HTTP 403).
- Bolsas/GA: a Graduate School diz que a maior parte do funding vem dos
  departamentos. O Rawls oferece GA institucionalmente (~US$22.750/ano,
  meio periodo, segundo a rodada anterior), mas nao ha confirmacao para o
  MSDS.
- Inicio: o programa presencial de 1 ano comeca no Summer (Summer I). A
  entrada de Spring e so para o formato online/part-time.

Uso: python scripts/etl_add_texas_tech_msds.py (a partir da raiz).
"""

import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
VERIFICAR = "Verificar"

NEW_PROGRAM = {
    "Nome da Universidade": "Texas Tech",
    "Estado": "TX",
    "Cidade": "Lubbock",
    "Business School": "Rawls College of Business",
    "Tipo (Pub/Priv)": "Pública",
    "Nome Oficial do Programa": (
        "Master of Science in Data Science (MSDS) -- o Rawls não tem MSBA; é o programa "
        "equivalente mais próximo (voltado ao mercado, sem tese)"
    ),
    "STEM Designated": "Sim",
    "Formato": (
        "Presencial (a versão de 1 ano também existe online; a versão de 2 anos é só "
        "online)"
    ),
    "Duração (meses)": 12,
    "Início": (
        "Summer (maio) para o programa presencial de 1 ano; entrada de Spring só no "
        "formato online/part-time"
    ),
    "Tamanho Turma": VERIFICAR,
    "Destaques Currículo": (
        "Statistics for Data Science, Database Concepts, Scripting Languages, Machine "
        "Learning, Big Data Strategy, Business Intelligence; capstone 'Data Science "
        "Project' (sem tese); formato lock-step em blocos de 7 semanas"
    ),
    "Total Créditos": 36,
    "Custo/Crédito (Intl) USD": 1178.67,
    "Custo/Crédito (Residente) USD": 768.67,
    "Tuition Total (Intl) USD": 42432,
    "Tuition Total (Residente) USD": 27672,
    "Fees Adicionais USD": (
        "Incluídas nos valores de tuition (a estimativa oficial 2026-27 da Graduate "
        "School junta tuition e fees, com o diferencial de business)"
    ),
    "Aluguel Privativo/mês USD": 720,
    "Mercado/mês USD": VERIFICAR,
    "Seguro Saúde (programa) USD": 3246,
    "Seguro Residencial/mês USD": VERIFICAR,
    "Utilities/mês USD": VERIFICAR,
    "Celular/mês USD": VERIFICAR,
    "WiFi/mês USD": VERIFICAR,
    "Transporte/mês USD": VERIFICAR,
    "Outros Gastos/mês USD": VERIFICAR,
    "Custo Vida Total (programa) USD": 19378.67,
    "Bolsas Disponíveis": VERIFICAR,
    "Tipos de Bolsa": (
        "Rawls College oferece GA/RA institucionalmente (~US$22.750/ano, meio período); "
        "não confirmado para o MSDS. A Graduate School diz que a maior parte do funding "
        "vem dos departamentos"
    ),
    "Quão Realista (Intl)": (
        "Incerto -- não há dado de bolsa ou GA específico do MSDS; perguntar ao programa"
    ),
    "TA/RA Disponível": VERIFICAR,
    "Processo TA/RA": VERIFICAR,
    "Caminho Tuition Residente": VERIFICAR,
    "Graduate Assistantships": (
        "Rawls College oferece GA/RA institucionalmente (~US$22.750/ano, meio período, "
        "com waiver de tuition out-of-state); não confirmado para o MSDS"
    ),
    "Ranking US News": VERIFICAR,
    "Ranking QS": VERIFICAR,
    "Taxa Empregabilidade %": VERIFICAR,
    "Salário Médio Pós USD": VERIFICAR,
    "Principais Empregadores": VERIFICAR,
    "GPA Mínimo": VERIFICAR,
    "GPA Médio Admitidos": VERIFICAR,
    "GMAT/GRE Obrigatório": (
        "Verificar -- fontes conflitam: versão anterior da base dizia que não era exigido; snippets "
        "de busca citam exigido, sem nota mínima, com dispensa possível para 3-5 anos de "
        "experiência profissional"
    ),
    "GMAT Médio": VERIFICAR,
    "Taxa Aceitação %": VERIFICAR,
    "Experiência Profissional": "Não obrigatória (programa também voltado a profissionais que já trabalham)",
    "Matérias Pré-Requisito": (
        "Nenhuma obrigatória; recomendado ter noções de R, SQL e Python e ter cursado "
        "cálculo, estatística e probabilidade"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 79 / TOEFL Essentials 8.5 / IELTS Academic 6.5 / Duolingo 100 / PTE "
        "60 / Cambridge CPE 180 / CAE 175 -- dispensado para diplomas dos EUA ou de "
        "países isentos"
    ),
    "Deadline Aplicação": (
        "Internacionais: 15 jan (recomendado pela Graduate School para Summer/Fall); "
        "Summer: 1 mai (prazo do programa); Spring: 1 dez (entrada só online/part-time) "
        "-- o programa pede que internacionais apliquem com 6+ meses de antecedência"
    ),
    "Taxa Aplicação USD": 75,
    "Rounds Aplicação": (
        "Sem rounds numerados; a taxa de US$75 sobe para US$100 depois de certas datas "
        "(aplicar cedo)"
    ),
    "Documentos Necessários": (
        "Transcripts não oficiais (com tradução oficial se não estiverem em inglês); "
        "resume detalhado com mês/ano de cada posição; teste de inglês; comprovação "
        "financeira (pode ser pedida após a admissão); cartas de recomendação não "
        "mencionadas na página do programa"
    ),
    "Entrevista": VERIFICAR,
    "Dificuldade Entrada": VERIFICAR,
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    missing = [h for h in headers if h not in NEW_PROGRAM]
    extra = [h for h in NEW_PROGRAM if h not in col]
    if missing or extra:
        raise ValueError(f"Schema divergente. Faltando: {missing}; sobrando: {extra}")

    for row in ws.iter_rows(min_row=2, values_only=True):
        if (row[col["Nome da Universidade"] - 1], row[col["Estado"] - 1], row[col["Cidade"] - 1]) == (
            "Texas Tech", "TX", "Lubbock"
        ):
            raise ValueError("Texas Tech (Lubbock, TX) já existe na base. Abortando.")

    ref_cell = ws.cell(row=2, column=1)
    base_font = Font(name=ref_cell.font.name, sz=ref_cell.font.sz)
    side = Side(style="thin", color="00D0D0D0")
    base_border = Border(left=side, right=side, top=side, bottom=side)
    base_alignment = Alignment(vertical="top", wrap_text=True)
    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")

    new_row = ws.max_row + 1
    for header, col_idx in col.items():
        value = NEW_PROGRAM[header]
        cell = ws.cell(row=new_row, column=col_idx, value=value)
        cell.font = base_font
        cell.border = base_border
        cell.alignment = base_alignment
        if value == VERIFICAR:
            cell.fill = verificar_fill

    last_col_letter = ws.cell(row=1, column=len(headers)).column_letter
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    # O Resumo estava desatualizado (dizia 114, mas a base tinha 109 linhas,
    # porque as remoções anteriores não o atualizaram). Recalcula pela contagem real.
    total = ws.max_row - 1
    summary = wb[SHEET_SUMMARY]
    for row in summary.iter_rows(min_row=2):
        label = row[0].value
        if label in ("Total de Programas", "Programas Presenciais STEM"):
            print(f"Resumo '{label}': {row[1].value} -> {total}")
            row[1].value = total
        if label == "Data de Compilação":
            row[1].value = f"{row[1].value} + Texas Tech (MSDS) em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"\nTexas Tech adicionada na linha {new_row}. Campos preenchidos:")
    filled = {h: v for h, v in NEW_PROGRAM.items() if v != VERIFICAR}
    for h, v in filled.items():
        print(f"  - {h}: {v}")
    print(f"\n{len(filled)} de {len(NEW_PROGRAM)} campos preenchidos; {len(NEW_PROGRAM) - len(filled)} como 'Verificar'.")


if __name__ == "__main__":
    main()
