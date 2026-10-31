# -*- coding: utf-8 -*-
"""
Data enrichment: Gonzaga University (School of Business Administration),
Spokane, WA -- Master of Science in Business Analytics (presencial, STEM).

Fontes (set/2026; gonzaga.edu bloqueia WebFetch, paginas baixadas via curl):
- gonzaga.edu/school-of-business-administration/graduate/msba (STEM; 30
  creditos; 2 semestres; merit automatica ate US$4,200/ano, sem aplicacao
  separada, anunciada junto com a admissao)
- .../msba/business-administration-track (on campus; aulas a tarde, noite e
  fim de semana; grade Fall/Spring; Tableau, Python, Orange, Excel;
  pre-requisito BMIS 235 com B+)
- .../msba/how-to-apply (Fall 2027: abre 1 fev 2027, prazo 1 ago 2027;
  transcripts; statement of purpose com 3 prompts de ~400 palavras; 2 cartas;
  resume/CV; uso de IA = negacao; nenhuma mencao a GMAT/GRE ou entrevista)
- .../msba/how-to-apply/international-students (TOEFL 80 ou 4.5 na nova
  escala; IELTS 6.5; avaliacao NACES course-by-course; passaporte; OPT STEM)
- .../msba/tuition-financial-aid (US$1,265/credito; 30 creditos)
- 2627-Graduate-Direct-Cost.pdf (MSBA 2026-27: tuition US$37,950 + Required
  Fees US$514 = US$38,464)
- graduate-cost-of-attendance 2026-27 (9 meses: housing US$12,648 (~1,410/mes),
  meals US$5,610 (~623/mes), books 528, personal 1,968, transport 2,796)
- international-graduate-scholarships ("limited number" de bolsas p/
  internacionais de programas presenciais, full-time; aplicacao separada)
- costs-scholarships-international-students (privada: internacional paga a
  mesma tuition)
- school-of-business-administration/about/facts-and-figures (U.S. News 2025:
  #107 Business Analytics; First Destination 2023 da pos da SBA inteira:
  97.5% "success rate" e top employers -- NAO especificos do MSBA)
- global.gonzaga.edu/graduate-admissions (Shorelight: TOEFL, IELTS, iTEP,
  Duolingo e PTE; GPA 3.0, 2.5+ pode ser avaliado -- pagina lista MBA/MAcc,
  nao o MSBA)
- Taxa de aplicacao US$50: guia oficial de aplicacao internacional da SBA
  (2018) -- valor antigo, confirmar no portal gradapply.

Mantidos como "Verificar": Tamanho Turma, Seguro Saude (obrigatorio, valor nao
publicado), Utilities/Celular/WiFi/Seguro Residencial, Ranking QS,
Empregabilidade e Salario do MSBA, GPA medio, GMAT medio, Taxa de Aceitacao.

Uso: python scripts/etl_enrich_gonzaga.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'Gonzaga University', 'Estado': 'WA', 'Cidade': 'Spokane'}

TUITION_PER_CREDIT = 1265
TOTAL_CREDITS = 30
TUITION_TOTAL = TUITION_PER_CREDIT * TOTAL_CREDITS  # 37,950

# Cost of Attendance 2026-27 (9 meses, Fall + Spring)
HOUSING_9M, MEALS_9M, BOOKS, PERSONAL_9M, TRANSPORT_9M = 12648, 5610, 528, 1968, 2796
CUSTO_VIDA_TOTAL = HOUSING_9M + MEALS_9M + BOOKS + PERSONAL_9M + TRANSPORT_9M  # 23,550

UPDATES = {
    'Formato': 'Presencial',
    'Duração (meses)': 9,
    'Início': 'Fall',
    'Destaques Currículo': 'Trilha Business Administration: Intro to Business Analytics, '
                           'Fundamental Tech for Business Analytics, Database Management & SQL, '
                           'Quant Stats, Descriptive, Regression e Predictive Analytics, Emerging '
                           'Technologies, Communication Skills: Persuasion; eletivas de 1 crédito em '
                           'people/marketing/finance analytics; Tableau, Python, Power BI, Orange; '
                           'aulas à tarde, noite e fim de semana; AACSB; trilha alternativa em '
                           'Accounting Analytics',
    'Total Créditos': TOTAL_CREDITS,
    'Custo/Crédito (Intl) USD': TUITION_PER_CREDIT,
    'Custo/Crédito (Residente) USD': TUITION_PER_CREDIT,
    'Tuition Total (Intl) USD': TUITION_TOTAL,
    'Tuition Total (Residente) USD': TUITION_TOTAL,
    'Fees Adicionais USD': 514,
    'Aluguel Privativo/mês USD': 1410,
    'Mercado/mês USD': 623,
    'Transporte/mês USD': round(TRANSPORT_9M / 9),
    'Outros Gastos/mês USD': round(PERSONAL_9M / 9),
    'Custo Vida Total (programa) USD': CUSTO_VIDA_TOTAL,
    'Bolsas Disponíveis': 'Sim',
    'Tipos de Bolsa': 'Merit automática de até US$4,200/ano (avaliada na admissão, sem aplicação '
                      'separada); International Graduate Student Scholarship ("limited number", '
                      'para programas presenciais full-time, com aplicação própria); AmeriCorps '
                      'US$1,000/ano; banco de bolsas externas',
    'Quão Realista (Intl)': 'Moderada -- a merit automática (até US$4.2k, ~11% da tuition) vale '
                            'para todos os candidatos; a bolsa internacional existe mas é limitada '
                            'e sem valor publicado; não há GA documentado para o MSBA',
    'TA/RA Disponível': 'Não documentado para o MSBA (a admissão internacional diz que "many" '
                        'programas presenciais oferecem GA -- perguntar)',
    'Processo TA/RA': 'Sem processo formal publicado; perguntar ao Admissions Specialist da '
                      'trilha Business Administration (Trevor Heilman, heilman@gonzaga.edu)',
    'Caminho Tuition Residente': 'N/A (privada -- internacionais pagam a mesma tuition)',
    'Graduate Assistantships': 'Não documentado para o MSBA',
    'Ranking US News': '#107 Business Analytics (U.S. News Best Graduate Business Schools 2025)',
    'Principais Empregadores': 'Moss Adams, KPMG, Deloitte, BDO USA, Grant Thornton, PwC -- '
                               'pós-graduação da School of Business como um todo (First '
                               'Destination 2023, 97.5% "success rate"), não específico do MSBA',
    'GPA Mínimo': 'Sem mínimo publicado na página do MSBA; referência da Gonzaga Global para '
                  'pós de negócios: 3.0 (2.5+ pode ser avaliado); GPA 3.0 exigido para '
                  'permanecer no programa',
    'GMAT/GRE Obrigatório': 'Não (não consta nos requisitos oficiais do MSBA)',
    'Experiência Profissional': 'Não obrigatória (apenas resume/CV)',
    'Matérias Pré-Requisito': 'BMIS 235 Management Information Systems (nota B+ ou superior); a '
                              'trilha Accounting exige também Intermediate Accounting I/II, Cost '
                              'Accounting, Federal Taxation I e Auditing (podem ser cursadas '
                              'após a admissão)',
    'Testes Inglês Aceitos': 'TOEFL iBT 80 (ou 4.5 na nova escala, desde jan/2026) / IELTS 6.5 '
                             '(página oficial do MSBA); Duolingo, iTEP e PTE aparecem só na '
                             'página da Gonzaga Global (Shorelight), sem nota mínima -- '
                             'confirmar para o MSBA',
    'Deadline Aplicação': '1 ago 2027 (Fall 2027, prazo único); aplicações abrem 1 fev 2027',
    'Taxa Aplicação USD': 50,
    'Rounds Aplicação': 'Sem rounds; janela única (abre 1 fev, fecha 1 ago)',
    'Documentos Necessários': 'Aplicação online; 1 transcript oficial de cada instituição + '
                              'avaliação course-by-course por NACES (TEC, ECE ou WES); statement '
                              'of purpose com 3 prompts de ~400 palavras (uso de IA = negação '
                              'automática); 2 cartas de recomendação (gestores ou professores); '
                              'resume/CV; teste de inglês; cópia do passaporte; comprovação '
                              'financeira (Certificate of Finances) para o I-20',
    'Entrevista': 'Não mencionada nos requisitos oficiais',
    'Dificuldade Entrada': 'Acessível',
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
