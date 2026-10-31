"""
Data enrichment: atualiza APENAS a linha da 'CU Boulder' (Leeds School
of Business, Boulder, CO) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, presencial, 10 meses, 33 créditos (já
confirmado), STEM, com 3 trilhas (Decision Science, Healthcare
Analytics, Marketing Analytics).

'Início' CORRIGIDO de 'July' para 'June' -- a página oficial de
admissões confirma "Program Start: June 2027" para o ciclo mais
recente, não julho.

TUITION: CORRIGIDO com base na página oficial "Estimated Cost of
Attendance" (rótulo explícito: rates 2025-2026, 3 semestres/33
créditos): tuition pura -- Residente $41.632 / Não-residente e
Internacional $57.853 (mesma tuition para ambos, só as FEES diferem).
Os valores anteriores ($41.315 residente / $57.160 Intl) eram
próximos mas não idênticos, provavelmente de uma rodada anterior com
rates de um ano diferente -- CORRIGIDOS para os valores oficiais mais
recentes e datados. 'Custo/Crédito' derivado dividindo pelos 33
créditos: $1.261,58 (residente) / $1.753,12 (internacional).

'Fees Adicionais USD' CORRIGIDO de $1.724 para $1.424 -- a mesma
página oficial discrimina fees por categoria de residência: Residente/
Não-residente doméstico $1.221; INTERNACIONAL $1.424 (fee adicional
específica para status internacional). Usado o valor internacional
por ser o foco desta pesquisa; o valor anterior ($1.724) não batia com
nenhuma das duas categorias oficiais.

CUSTO DE VIDA: a mesma página oficial fornece um orçamento anual
ITEMIZADO (não um bundle vago) para a categoria "US non-resident"
(usada como proxy para internacional, já que a fonte não separa
"international" dos demais itens de custo de vida, apenas de tuition/
fees): Room & Board (combinado, off-campus) $18.684/ano; Books &
Supplies $1.200/ano; Transportation $2.016/ano; Medical $8.136/ano;
Personal $1.782/ano; total de 12 meses = $31.818.

Mapeado para o schema: 'Transporte/mês USD' = $2.016 ÷ 12 = $168,00;
'Seguro Saúde (programa) USD' = $8.136 (linha "Medical" do orçamento
oficial, interpretada como estimativa do seguro saúde obrigatório);
'Outros Gastos/mês USD' = (Books $1.200 + Personal $1.782) ÷ 12 =
$248,50 (combinados por não haver campo específico de material
didático no schema); 'Custo Vida Total (programa) USD' = $31.818
(total oficial de 12 meses). 'Aluguel Privativo/mês' e 'Mercado/mês'
PERMANECEM 'Verificar' -- a fonte só dá "Room & Board" como valor
COMBINADO (aluguel + alimentação), sem separação; decompor usando uma
proporção inventada violaria o princípio já estabelecido nesta sessão
de nunca inventar uma decomposição sobre um bundle oficial.

'Testes Inglês Aceitos' refinado com todos os mínimos oficiais: TOEFL
iBT mín. 80 (preferencial 100); TOEFL iBT26 mín. 4; IELTS mín. 6.5;
Duolingo mín. 115 (escala de 160); PTE Pearson mín. 58; Cambridge mín.
180. 'Taxa Aplicação USD' preenchido: $0 -- taxa base é $70, mas está
ATUALMENTE DISPENSADA para todos os candidatos (confirmado
oficialmente).

'Documentos Necessários' preenchido: resume profissional; 1 carta de
recomendação; essay de até 750 palavras; transcripts não-oficiais de
todas as instituições; GMAT/GRE opcional. 'Entrevista' preenchido:
NÃO exigida para candidatos de MS (confirmado oficialmente, negativa
explícita).

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com os 5 prazos
oficiais do ciclo de 2027: Round 1 (5 out 2026); Round 2 (30 nov
2026); Round 3 (1 fev 2027); Final Internacional (3 fev 2027); Final
Doméstico (10 mai 2027).

'Ranking US News' preenchido: '#39 Best Business Analytics Specialty
(U.S. News 2026)' -- dado apresentado diretamente na página oficial do
programa.

'Salário Médio Pós USD' preenchido com $83.885 -- oficial e específico
do programa, mas da coorte de 2023 (RESSALVA: dado com ~2-3 anos,
documentado como tal). 'Principais Empregadores' preenchido com uma
amostra da lista oficial de +50 empresas: Deloitte, Accenture, KPMG,
Epic Systems, Lockheed Martin, Comcast, DaVita.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação
%', 'Ranking QS', 'Matérias Pré-Requisito', 'Processo TA/RA', 'Caminho
Tuition Residente', 'Aluguel Privativo/mês' e 'Mercado/mês' (ver nota
acima sobre bundle "Room & Board"), 'Seguro Residencial/mês',
'Utilities/mês', 'Celular/mês', 'WiFi/mês'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.colorado.edu/business/ms-programs/masters-program-business-analytics
  (formato, duração, início corrigido, STEM, currículo, deadlines,
  ranking US News #39)
- .../masters-programs-admissions (mínimos de inglês completos, taxa
  de aplicação dispensada, documentos, ausência de entrevista)
- .../masters-program-business-analytics/msba-msma-and-mssca-estimated-cost-attendance
  (tabela oficial e itemizada de custo de atendimento 2025-26: tuition
  por residência, fees por residência incluindo internacional
  específico, orçamento de 12 meses de custo de vida)
- .../masters-program-business-analytics/career-success (salário
  médio 2023, empregadores)
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
    "Nome da Universidade": "CU Boulder",
    "Estado": "CO",
    "Cidade": "Boulder",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Início": "June",  # CORREÇÃO: oficial "Program Start: June 2027", não July
    "Custo/Crédito (Intl) USD": 1753.12,  # 57.853 / 33, oficial 2025-26
    "Custo/Crédito (Residente) USD": 1261.58,  # 41.632 / 33, oficial 2025-26
    "Tuition Total (Intl) USD": 57853,  # CORREÇÃO: oficial 2025-26 (era 57.160)
    "Tuition Total (Residente) USD": 41632,  # CORREÇÃO: oficial 2025-26 (era 41.315)
    "Fees Adicionais USD": 1424,  # CORREÇÃO: oficial, fee específica internacional (era 1.724, não batia com nenhuma categoria)
    "Transporte/mês USD": 168.00,  # $2.016/ano (US non-resident/intl) ÷ 12, oficial
    "Seguro Saúde (programa) USD": 8136,  # linha "Medical" do orçamento oficial anual
    "Outros Gastos/mês USD": 248.50,  # (Books $1.200 + Personal $1.782) / 12, oficial
    "Custo Vida Total (programa) USD": 31818,  # total oficial de 12 meses (US non-resident/intl)
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 80 (preferencial 100); TOEFL iBT26 mín. 4; "
        "IELTS mín. 6.5; Duolingo mín. 115 (escala 160); PTE Pearson "
        "mín. 58; Cambridge mín. 180"
    ),
    "Taxa Aplicação USD": 0,  # taxa base $70, ATUALMENTE dispensada para todos (confirmado oficialmente)
    "Documentos Necessários": (
        "Resume profissional; 1 carta de recomendação; essay até 750 "
        "palavras; transcripts não-oficiais de todas as instituições; "
        "GMAT/GRE opcional"
    ),
    "Entrevista": "Não exigida para candidatos de MS (confirmado oficialmente)",
    "Deadline Aplicação": "Round 1: 5 out 2026; Round 2: 30 nov 2026; Round 3: 1 fev 2027; Final Internacional: 3 fev 2027; Final Doméstico: 10 mai 2027",
    "Rounds Aplicação": "5 prazos: 5 out, 30 nov, 1 fev, 3 fev (final intl), 10 mai (final doméstico)",
    "Ranking US News": "#39 Best Business Analytics Specialty (U.S. News 2026)",
    "Salário Médio Pós USD": 83885,  # oficial, coorte 2023 (dado com ~2-3 anos)
    "Principais Empregadores": "Deloitte, Accenture, KPMG, Epic Systems, Lockheed Martin, Comcast, DaVita",
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
        raise ValueError("Nenhuma linha encontrada para CU Boulder (Boulder, CO). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (CU Boulder, Boulder, CO)")

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
