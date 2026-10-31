"""
Data enrichment: atualiza APENAS a linha 'Southern Connecticut State
University (SCSU)' (School of Business, New Haven, CT) no
msba_programs_database.xlsx.

Programa: STEM Master of Science in Business Analytics (MSBA),
presencial (on-ground), STEM.

FONTES OFICIAIS (acessadas em 2026-10-03):
  - https://www.southernct.edu/academics/programs/business-analytics-ms
  - https://catalog.southernct.edu/graduate/programs-and-degrees/business-analytics-ms.html
  - https://go.southernct.edu/docs/drupal/tuition-fees-fulltime-2026-2027.pdf
  - https://go.southernct.edu/docs/drupal/tuition-fees-parttime-2026-2027.pdf
  - https://grad.southernct.edu/admissions/
  - https://catalog.southernct.edu/graduate/admission-standards/international-students.html
  - https://inside.southernct.edu/sites/default/files/inline-files/international-student-profile-graduate.pdf
  - https://grad.southernct.edu/graduate-assistantships-2/
  - https://grad.southernct.edu/graduate-scholarships/
  - https://inside.southernct.edu/international/study-at-scsu/matriculated-students/insurance

ACHADOS / CORREÇÕES:
- Tuition: em tempo integral (pós = 9+ créditos/semestre, obrigatório
  para F-1) a SCSU cobra um valor FIXO por semestre, não por crédito.
  Tabela 2026-27, coluna "Graduate" (o MSBA não está nas colunas de MBA/
  MLIS/Sports Mgmt.):
    Out-of-State/International (tarifa NEBHE desde o Fall 2026):
      Tuition US$ 6.537/semestre
    In-State: Tuition US$ 4.358/semestre
    Fees (iguais para os dois): University General Fee US$ 2.547 +
      University Fee US$ 528 + Student Activity Fee US$ 27 + Writing
      Center Fee US$ 20 = US$ 3.122/semestre
  O programa leva "typically two years" (página oficial) = 4 semestres:
    Tuition Total Intl = 4 x 6.537 = US$ 26.148 (efetivo US$ 871,60/crédito)
    Tuition Total Residente = 4 x 4.358 = US$ 17.432 (efetivo US$ 581,07/crédito)
    Fees = 4 x 3.122 = US$ 12.488
  Os valores antigos (US$ 702,10/crédito, US$ 21.063 e US$ 10.619 de
  fees) não batem com a tabela vigente. Dá para terminar em 3 semestres
  (12+9+9 créditos); como o valor é fixo, isso reduziria o custo em 1/4.
- Bolsas: a única bolsa interna da pós (US$ 1.000/ano) é só para quem
  veio do programa acelerado de graduação da SCSU. O formulário
  internacional diz que internacionais "may be considered for university
  scholarships", sem detalhes. A ajuda real são as Graduate
  Assistantships: US$ 5.280/semestre (US$ 17,60/h, 20 h/semana) + isenção
  de University Fee, University General Fee (exceto parte do seguro de
  acidentes) e Student Activity Fee -- NÃO isenta tuition. Vagas
  publicadas uma a uma no site, candidatura por e-mail, nomeação pelo
  dean por recomendação do chefe do departamento.
- Prazos: o catálogo 2026-27 diz que internacionais novos precisam
  entregar as credenciais à International Student Services até 1 abr
  (Fall) e 1 out (Spring). O formulário PDF (provavelmente antigo) diz
  1 mai. O prazo do programa no portal é 17 ago (Fall) / 3 jan (Spring).
- Documentos: internacionais também entregam ao OIE o International
  Student Profile Form e um extrato financeiro original cobrindo pelo
  menos 1 ano de tuition, fees e moradia; avaliação course-by-course
  NACES com tradução certificada. Nenhuma página oficial pede GRE/GMAT,
  cartas ou redação para o MSBA (requisitos: bacharelado + GPA 3.0).
- Seguro saúde: obrigatório para internacionais (plano ISO ou equivalente
  aprovado); a SCSU não publica o preço -- permanece 'Verificar'.

NÃO ALTERADOS: Salário Médio Pós (a página só cita a mediana nacional
do BLS para "management analysts", US$ 101.190 -- dado genérico),
Taxa Empregabilidade, Rankings, Tamanho Turma, GPA Médio, Taxa Aceitação.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Southern Connecticut State University (SCSU)",
    "Estado": "CT",
    "Cidade": "New Haven",
}

UPDATES = {
    "Destaques Currículo": (
        "30 créditos (+ até 3 de pré-requisito dispensáveis: ECO 500 Business "
        "Statistics / BIS 505 Database Systems): 18 de core (visualização e "
        "dashboards, machine learning, simulação e otimização, IA para "
        "decisão, banco de dados, métodos quantitativos) + 9 de eletivas "
        "(People, Accounting, Financial, Healthcare, Marketing e Operational "
        "Analytics, Forecasting, Law & Ethics) + capstone BUS 750 de "
        "consultoria com empresas; R, Python, SQL, Power BI, Tableau, "
        "Bloomberg; AACSB; OPT STEM de 3 anos"
    ),
    "Custo/Crédito (Intl) USD": 871.6,  # efetivo: 4 x US$ 6.537 / 30 créditos (tarifa fixa por semestre)
    "Custo/Crédito (Residente) USD": 581.07,  # efetivo: 4 x US$ 4.358 / 30 créditos
    "Tuition Total (Intl) USD": 26148,
    "Tuition Total (Residente) USD": 17432,
    "Fees Adicionais USD": 12488,  # 4 x (2.547 + 528 + 27 + 20)
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Graduate Assistantships: US$ 5.280/semestre (US$ 17,60/h, 20 "
        "h/semana) + isenção de ~US$ 3.100/semestre em fees (University "
        "Fee, University General Fee e Student Activity Fee), sem isenção de "
        "tuition. A única bolsa interna da pós (US$ 1.000/ano) é só para "
        "egressos do programa acelerado da SCSU; internacionais 'may be "
        "considered for university scholarships', sem detalhes. Sem ajuda "
        "federal para internacionais; bolsas externas listadas no site"
    ),
    "Quão Realista (Intl)": (
        "Baixa a moderada -- GA é a única ajuda concreta e vale bastante "
        "(~US$ 8,4k por semestre entre salário e fees), mas as vagas são "
        "publicadas uma a uma, não são do MSBA e são disputadas por toda a "
        "pós. Planejar o custo integral"
    ),
    "Processo TA/RA": (
        "Depois da admissão: acompanhar as vagas publicadas na página de "
        "Graduate Assistantships e enviar por e-mail cover letter, resume e "
        "horários disponíveis; nomeação pelo dean da unidade por "
        "recomendação do chefe do departamento. Exige admissão plena e GPA "
        "3.0 na pós"
    ),
    "Graduate Assistantships": (
        "Sim, em nível de universidade: US$ 5.280/semestre (US$ 17,60/h, 20 "
        "h/semana) + isenção de University Fee, University General Fee "
        "(exceto a parte do seguro de acidentes) e Student Activity Fee; NÃO "
        "isenta tuition; vagas para o MSBA não divulgadas"
    ),
    "GMAT/GRE Obrigatório": "Não exigido (os requisitos oficiais do MSBA são só bacharelado e GPA 3.0)",
    "Experiência Profissional": "Não obrigatória (não consta nos requisitos)",
    "Testes Inglês Aceitos": (
        "TOEFL iBT 79 (antes de 2026) ou 4.5 (a partir de 1 jan 2026) / "
        "TOEFL PBT 550 / IELTS 6.5 / Duolingo 105 / PTE 53 / ELS nível 112; "
        "nota de no máximo 2 anos, enviada direto à SCSU (código 3662)"
    ),
    "Deadline Aplicação": (
        "1 abr (Fall, internacional: aplicação, credenciais e International "
        "Student Profile Form com comprovação financeira; o formulário PDF "
        "diz 1 mai); 17 ago (Fall, prazo do programa no portal). Spring: 1 "
        "out (internacional) / 3 jan (programa). Para ingresso (Fall "
        "2028): ~1 abr 2028"
    ),
    "Rounds Aplicação": (
        "Sem rounds -- prazos fixos por semestre; internacionais novos têm "
        "prazo antecipado (1 abr Fall / 1 out Spring, catálogo 2026-27). "
        "Completar o arquivo até o prazo dá prioridade; depois do prazo não "
        "se aceitam aplicações"
    ),
    "Documentos Necessários": (
        "Aplicação online (taxa US$ 50, não reembolsável); transcripts "
        "oficiais de todas as instituições, com cópias e traduções "
        "certificadas para inglês; avaliação course-by-course NACES (WES, "
        "Josef Silny, GCE ou ECE); comprovante de inglês. Para o OIE: "
        "International Student Profile Form e extrato financeiro original "
        "cobrindo pelo menos 1 ano de tuition, fees e moradia. Nenhuma "
        "página oficial pede GRE/GMAT, cartas ou redação para o MSBA"
    ),
    "Entrevista": "Não mencionada nas páginas oficiais",
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
        raise ValueError("Nenhuma linha encontrada para SCSU (New Haven, CT). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SCSU, New Haven, CT)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
