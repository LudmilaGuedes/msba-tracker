# -*- coding: utf-8 -*-
"""
Data enrichment: University of San Diego (Knauss School of Business),
San Diego, CA -- Master of Science in Business Analytics (presencial, STEM).
Nao confundir com "UC San Diego" (Rady, La Jolla), outra linha da base.

Fontes (set/2026; paginas baixadas via curl):
- sandiego.edu/business/academics/graduate/business-analytics/ (full-time,
  on campus, inicio no Fall, 10 meses, STEM; tuition 2026-27 US$64,440;
  82% empregados em 6 meses e salario inicial medio US$74,237 -- media de 3
  anos; 78% recebem auxilio; 3 projetos: international consulting de 2
  semanas no exterior, social impact e capstone; SQL, Python, R, Tableau,
  Power BI, IBM SPSS Modeler; experiencia da turma de 0 a 6.5 anos)
- .../admission-and-aid/graduate/msba-how-to-apply.php (sem GPA minimo, media
  3.4 na coorte 2024; estatistica universitaria com B-; 2 anos de experiencia
  preferidos, nao exigidos; Rounds 1 nov, 15 jan, 1 mar, 1 abr, 1 mai +
  estendido 15 jul so domesticos; rolling, vagas e bolsas por ordem de
  chegada; resume, essay de 250-500 palavras, taxa US$125 salvo fee waiver,
  ate 2 cartas; entrevista por convite; test-optional; Dean's Merit; 36
  unidades em 10 meses, 18 por semestre; ISIP obrigatorio p/ internacionais)
- .../admission-and-aid/tuition-and-scholarships.php (US$1,790/unidade;
  MSBA: Dean's Merit automatica + Graduate Assistantship de 40h com US$1,200
  de reducao de tuition; bolsas gerais da pos com aplicacao)
- torero-hub/tuition-and-fees/graduate.php (2026-27: seguro saude US$4,158
  por 12 meses; por semestre: Wellness US$170, Facility US$150, GSA US$64,
  Student Life Pavilion US$35; custo de vida de pos off-campus independente,
  9 meses: moradia 16,920, alimentacao 7,884, pessoais 5,706, transporte
  1,656, livros 1,320)
- admission-and-aid/graduate/apply/.../international-applicants/
  application-checklist.php (TOEFL iBT 4.5 / 83 antes de 21 jan 2026; IELTS
  7.0; Duolingo 120; avaliacao course-by-course por WES/ECE/TEC/IERF/FACS)
- businessresources.sandiego.edu (ex-alunos da Knauss em HelloFresh, Axos
  Bank, Chosen Foods, ResMed, Qral Group, EDF Renewables, LiveGoode; fee
  waiver de US$125 em workshops de aplicacao / reuniao com advisor)
- U.S. News: Knauss #105 (empate) em Best Business Schools; sem ranking
  especifico do MSBA.

Custo de vida: o orcamento oficial cobre 9 meses; o programa dura 10, entao
os gastos mensais foram multiplicados por 10 (livros contam uma vez).

Uso: python scripts/etl_enrich_usd.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'University of San Diego', 'Estado': 'CA', 'Cidade': 'San Diego'}

TUITION_PER_UNIT = 1790
TOTAL_UNITS = 36
TUITION_TOTAL = TUITION_PER_UNIT * TOTAL_UNITS  # 64,440
DURATION_MONTHS = 10

# Taxas obrigatorias por semestre (2026-27), 2 semestres; seguro saude fica em coluna propria
FEES_PER_SEMESTER = 170 + 150 + 64 + 35  # Wellness + Facility + GSA (>6.5 unidades) + Pavilion
FEES_TOTAL = FEES_PER_SEMESTER * 2  # 838

# Custo de vida de pos off-campus independente, 9 meses (2026-27)
HOUSING_9M, MEALS_9M, PERSONAL_9M, TRANSPORT_9M, BOOKS = 16920, 7884, 5706, 1656, 1320
CUSTO_VIDA_TOTAL = round((HOUSING_9M + MEALS_9M + PERSONAL_9M + TRANSPORT_9M) / 9 * DURATION_MONTHS) + BOOKS

UPDATES = {
    'Formato': 'Presencial',
    'Duração (meses)': DURATION_MONTHS,
    'Início': 'Fall',
    'Destaques Currículo': '36 unidades em 10 meses (18 por semestre, full-time); 3 projetos '
                           'práticos: international consulting de 2 semanas no exterior (ex.: '
                           'Madri, Munique), Data for Social Good e capstone com empresa parceira; '
                           'Advanced Analytics for AI, Prescriptive Analytics; SQL, Python, R, '
                           'Tableau, Power BI e IBM SPSS Modeler; tarefas preparatórias de Python, '
                           'banco de dados e estatística antes do início',
    'Total Créditos': TOTAL_UNITS,
    'Custo/Crédito (Intl) USD': TUITION_PER_UNIT,
    'Custo/Crédito (Residente) USD': TUITION_PER_UNIT,
    'Tuition Total (Intl) USD': TUITION_TOTAL,
    'Tuition Total (Residente) USD': TUITION_TOTAL,
    'Fees Adicionais USD': FEES_TOTAL,
    'Aluguel Privativo/mês USD': round(HOUSING_9M / 9),
    'Mercado/mês USD': round(MEALS_9M / 9),
    'Seguro Saúde (programa) USD': 4158,
    'Transporte/mês USD': round(TRANSPORT_9M / 9),
    'Outros Gastos/mês USD': round(PERSONAL_9M / 9),
    'Custo Vida Total (programa) USD': CUSTO_VIDA_TOTAL,
    'Bolsas Disponíveis': 'Sim',
    'Tipos de Bolsa': "Dean's Merit (automática na admissão, valor variável informado na "
                      'oferta; vagas e bolsas por ordem de chegada); Graduate Assistantship do '
                      'MSBA (40h com professor da Knauss = US$1,200 de redução de tuition, com '
                      'aplicação); bolsas gerais da pós com aplicação após o início (John '
                      'Ronchetto até US$6,450 dividida; BusinessLink até US$4,125); desconto de '
                      'tuition no study abroad; empréstimo Prodigy Finance para internacionais; '
                      '78% dos alunos recebem algum auxílio',
    'Quão Realista (Intl)': "Moderada -- a Dean's Merit é automática e não exige cidadania, mas o "
                            'valor não é publicado e as bolsas acabam por ordem de chegada '
                            '(aplicar no Round 1); o GA vale só US$1,200',
    'TA/RA Disponível': 'Sim (Graduate Assistantship: 40h com professor da Knauss, US$1,200 de '
                        'redução de tuition)',
    'Processo TA/RA': 'Aplicação pela página de bolsas da Knauss (Tuition and Scholarships), '
                      'para alunos matriculados no MSBA',
    'Caminho Tuition Residente': 'N/A (privada -- mesma tuition para internacionais)',
    'Graduate Assistantships': 'US$1,200 de redução de tuition por uma assistantship de 40 horas '
                               '(valor simbólico, sem stipend)',
    'Ranking US News': 'Sem ranking específico do MSBA; Knauss #105 (empate) em Best Business '
                       'Schools (U.S. News)',
    'Taxa Empregabilidade %': 82,
    'Salário Médio Pós USD': 74237,
    'Principais Empregadores': 'HelloFresh, Axos Bank, Chosen Foods, ResMed, Qral Group, EDF '
                               'Renewables, LiveGoode (ex-alunos da Knauss citados na página do '
                               'MSBA)',
    'GPA Mínimo': 'Sem mínimo (avaliação holística)',
    'GPA Médio Admitidos': 3.4,
    'GMAT/GRE Obrigatório': 'Opcional (test-optional)',
    'Experiência Profissional': 'Não obrigatória (2 anos preferidos; média da turma 2 anos, faixa '
                                'de 0 a 6.5)',
    'Matérias Pré-Requisito': '1 semestre de estatística universitária com B- ou mais (ou curso '
                              'online credenciado de Business Statistics antes da decisão); após '
                              'a admissão, tarefas preparatórias de Python, banco de dados e '
                              'estatística',
    'Testes Inglês Aceitos': 'TOEFL iBT 4.5 na nova escala (83 para provas antes de 21/jan/2026) / '
                             'IELTS 7.0 / Duolingo 120; válidos por 2 anos e numa única data; '
                             'dispensa só com bacharelado de 4 anos em inglês e GPA 3.0+',
    'Deadline Aplicação': 'Round 1: 1 nov 2026; Round 2: 15 jan 2027; Round 3: 1 mar 2027; Round '
                          '4: 1 abr 2027; Round 5: 1 mai 2027 (Fall 2027); prazo estendido 15 jul '
                          'só para domésticos',
    'Taxa Aplicação USD': 125,  # isenção em workshops de aplicação da Knauss / reunião com advisor
    'Rounds Aplicação': '5 rounds + prazo estendido (15 jul, só domésticos); admissão rolling -- '
                        'vagas e bolsas por ordem de chegada, então aplicar no Round 1',
    'Documentos Necessários': 'Aplicação online; transcripts oficiais + avaliação course-by-course '
                              '(WES, ECE, TEC, IERF ou FACS); resume; essay de 250-500 palavras; '
                              'até 2 cartas (professores ou supervisores); teste de inglês; '
                              'GMAT/GRE opcional; após admissão, documentos financeiros para o '
                              'I-20 e ISIP obrigatório em agosto',
    'Entrevista': 'Por convite apenas',
    'Dificuldade Entrada': 'Moderada',
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    unknown = [h for h in UPDATES if h not in col]
    if unknown:
        raise ValueError(f"Colunas inexistentes no schema: {unknown}")

    rows = [
        r[0].row for r in ws.iter_rows(min_row=2)
        if all(r[col[k] - 1].value == v for k, v in MATCH.items())
    ]
    if len(rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, encontrou {len(rows)}. Abortando.")
    row_idx = rows[0]

    changed = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        if cell.value == new_value:
            continue
        changed.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = PatternFill(fill_type=None)

    wb.save(DB_PATH)
    print(f"\n{len(changed)} campos atualizados na linha {row_idx} ({MATCH['Nome da Universidade']}):")
    for header, before, after in changed:
        print(f"  - {header}:")
        print(f"      Antes:  {before!r}")
        print(f"      Depois: {after!r}")


if __name__ == "__main__":
    main()
