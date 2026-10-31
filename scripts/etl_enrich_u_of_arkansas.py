"""
Data enrichment: atualiza APENAS a linha da 'U of Arkansas' (Walton
College of Business, Fayetteville, AR) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (walton.uark.edu).

Foco: Master of Applied Business Analytics (MABA), trilha FULL-TIME
(distinta da trilha "Professional MABA", part-time, 5 semestres com
boot camp — não tocada aqui). 30 créditos, 1 ano, STEM.

TUITION: PREENCHIDO (estava 'Verificar'). Tabela oficial de custos
2026-27 do MABA full-time: $741,88/crédito (residente AR) / $1.974,78/
crédito (fora do estado/internacional). 'Tuition Total': Residente =
30 × $741,88 = $22.256,40 (bate exatamente com o total oficial
divulgado). Internacional: a própria página divulga um total OFICIAL
de $62.583,40 -- MAIOR que o total "fora do estado" puro (30 ×
$1.974,78 = $59.243,40) por $3.340,00. A página explica essa diferença
como custos adicionais específicos de internacionais ("a one-time
orientation fee, a per-semester student fee, and an annual insurance
fee"), mas SEM abrir o valor de cada um separadamente -- por isso o
delta de $3.340,00 foi usado como 'Fees Adicionais USD' (ao invés de
tentar separar arbitrariamente entre taxa de orientação/taxa
estudantil/seguro), e 'Seguro Saúde (programa) USD' permanece
'Verificar' para não correr risco de dupla contagem com esse bundle já
preenchido.

'GMAT/GRE Obrigatório' preenchido com achado IMPORTANTE para o perfil
de busca (aluna internacional): ao contrário de candidatos domésticos
(que podem dispensar o teste com GPA ≥3.2 de instituição dos EUA, 3+
anos de experiência profissional, ou credenciais excepcionais como JD/
MD/PhD), CANDIDATOS INTERNACIONAIS DEVEM APRESENTAR GMAT OU GRE
OBRIGATORIAMENTE, sem direito à dispensa -- confirmado explicitamente
na página oficial de admissão.

'GPA Mínimo' preenchido com nota de que o 3.2 citado é o limiar de
DISPENSA de teste para domésticos, não um mínimo geral de admissão (não
divulgado separadamente). 'Testes Inglês Aceitos' refinado: TOEFL/
IELTS (já presente) + Duolingo English Test também aceito (mínimos não
divulgados). 'Taxa Aplicação USD' preenchido: $60 (doméstico) / $75
(internacional). 'Documentos Necessários' preenchido com a lista
oficial (application, transcripts, statement of purpose até 2 páginas,
resume, 3 referências profissionais via portal).

'Deadline Aplicação' preenchido com os prazos ESPECÍFICOS da trilha
full-time (foco desta pesquisa): Fall 2026 full-time JÁ FECHOU; Fall
2027 full-time -- internacional: 15 fev; doméstico: 1 jul. CUIDADO: a
mesma seção de admissões geral do Walton também lista prazos "1 nov
(internacional) / 1 mar (doméstico prioritário)" associados a um
"boot camp de 9 ago" e "5 semestres" -- esses parecem pertencer à
trilha Professional/part-time, não à full-time, e por isso NÃO foram
usados aqui (documentado para evitar confusão futura).

'TA/RA Disponível', 'Graduate Assistantships', 'Bolsas Disponíveis' e
'Tipos de Bolsa' preenchidos com CONFIRMAÇÃO NEGATIVA OFICIAL: a
própria página de tuition/funding do Walton declara explicitamente que
"No other Walton master's degree provides any direct funding (such as
Graduate Assistantship, Research Assistantship, Teaching Assistantship,
Scholarship or Fellowship)" -- aplicável ao MABA. Não é uma lacuna de
pesquisa, é uma negativa confirmada oficialmente. 'Quão Realista
(Intl)' preenchido como 'Baixa' com base nessa mesma confirmação.

'Caminho Tuition Residente' preenchido com nota geral de
inaplicabilidade prática para internacionais em programa de 1 ano
(mesmo padrão usado para outras universidades públicas nesta sessão).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Seguro Saúde (programa) USD' (mandatório confirmado, mas sem valor de
prêmio divulgado na página consultada), 'Tamanho Turma', 'GPA Médio
Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Ranking US News',
'Ranking QS', 'Taxa Empregabilidade %', 'Salário Médio Pós USD'
(a página só oferece um vídeo de outcomes, sem estatísticas em texto),
'Entrevista', e todos os campos mensais de custo de vida ainda vazios
(Utilities, Celular, WiFi, Transporte, Seguro Residencial, Outros
Gastos, Custo Vida Total) -- página de cost of attendance da
universidade retornou erro de DNS.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://walton.uark.edu/graduate-programs/business-analytics-masters-degree/index.php
  (formato full-time/part-time, duração, créditos, currículo, STEM,
  deadlines da trilha full-time)
- https://walton.uark.edu/graduate-programs/admissions/tuition-cost.php
  (tuition oficial 2026-27, ausência confirmada de GA/RA/TA/bolsa/
  fellowship para o MABA)
- https://walton.uark.edu/graduate-programs/apply.php (GPA/GMAT-GRE
  com critérios de dispensa — internacionais SEM dispensa —, Duolingo
  aceito, taxa de aplicação, documentos exigidos, deadlines gerais da
  trilha Professional/part-time)
- http://health.uark.edu/billing-insurance/insurance.php/ (seguro
  saúde obrigatório para internacionais, sem valor de prêmio
  divulgado)
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
    "Nome da Universidade": "U of Arkansas",
    "Estado": "AR",
    "Cidade": "Fayetteville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,  # oficial, trilha full-time
    "Total Créditos": 30,  # oficial
    "Custo/Crédito (Intl) USD": 1974.78,  # oficial 2026-27, mesma tarifa de out-of-state
    "Custo/Crédito (Residente) USD": 741.88,  # oficial 2026-27
    "Tuition Total (Intl) USD": 62583.40,  # total oficial, inclui custos extras bundled de internacionais
    "Tuition Total (Residente) USD": 22256.40,  # 30 x $741,88, oficial
    "Fees Adicionais USD": 3340.00,  # diferença entre total Intl oficial e total out-of-state puro (orientação + taxa estudantil + seguro, não itemizado pela fonte)
    "Caminho Tuition Residente": "Não aplicável na prática para internacionais em programa de 1 ano — tarifa não-residente/internacional é a relevante",
    "GMAT/GRE Obrigatório": (
        "OBRIGATÓRIO para candidatos internacionais (sem direito a "
        "dispensa); domésticos podem dispensar com GPA ≥3.2 de "
        "instituição dos EUA, 3+ anos de experiência profissional, "
        "ou credenciais excepcionais (JD/MD/PhD)"
    ),
    "GPA Mínimo": "3.2 é o limiar de dispensa de GMAT/GRE para domésticos; sem mínimo geral de admissão divulgado separadamente",
    "Testes Inglês Aceitos": "TOEFL/IELTS; Duolingo English Test também aceito (mínimos não divulgados)",
    "Taxa Aplicação USD": 75,  # internacional (doméstico: $60)
    "Documentos Necessários": (
        "Online application; transcripts (não-oficiais aceitos "
        "inicialmente); Statement of Purpose (máx. 2 páginas); "
        "resume; 3 referências profissionais via portal"
    ),
    "Deadline Aplicação": "Full-time Fall 2026: encerrado; Full-time Fall 2027 — internacional: 15 fev; doméstico: 1 jul",
    "TA/RA Disponível": "Não disponível — confirmado oficialmente: nenhum master's do Walton (incl. MABA) oferece GA/RA/TA/bolsa/fellowship",
    "Graduate Assistantships": "Não disponível — mesma confirmação oficial de TA/RA",
    "Bolsas Disponíveis": "Não disponível — confirmado oficialmente (sem funding direto para o MABA)",
    "Tipos de Bolsa": "N/A — sem bolsas/fellowships para este programa (confirmado oficialmente)",
    "Quão Realista (Intl)": "Baixa — sem bolsas nem assistantships; tuition integral, financiamento via fundos próprios/empréstimos",
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para U of Arkansas (Fayetteville, AR). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Arkansas, Fayetteville, AR)")

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
