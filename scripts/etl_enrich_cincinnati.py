"""
Data enrichment: atualiza APENAS a linha da 'U of Cincinnati' (Carl H.
Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial (full-time,
~9 meses), STEM. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e
custo de vida local.

AUDITORIA: 'Tuition Total (Intl) USD' ($33.195) foi CONFIRMADA por dois
valores de terceiros independentes e próximos ($33.070 e $33.320,
atualizados em 2026), sem alteração. Como 'Caminho Tuition Residente' já
estava cadastrado como "Não" (sem trilha de tarifa reduzida para
residentes -- outra fonte confirma tarifa única/flat para o "specialty
master's" independente de residência), 'Tuition Total (Residente) USD'
foi preenchido com o MESMO valor da tuition internacional.

ACHADO CRITICO: 'Testes Inglês Aceitos' estava com limiares desatualizados
(TOEFL 80/IELTS 6.5/Duolingo 110). A página oficial de admissions
requirements da Lindner (fonte atual) lista limiares mais altos: TOEFL 90 /
IELTS 7.0 / Duolingo 115. CORRIGIDO (o PTE 54 pré-existente foi mantido, já
que a nova fonte não mencionou PTE e não há motivo para descartá-lo).

Custo de vida: apenas 'Seguro Saúde (programa)' pôde ser preenchido com
confiança -- a taxa oficial do UC Student Health Insurance Plan (Aetna)
para o ano letivo 2026-27 foi anunciada em agosto/2026 e bate exatamente
com a duração do programa (Fall+Spring = 9 meses). Os demais 8 componentes
mensais de custo de vida permanecem "Verificar": os únicos dados de aluguel
encontrados eram listagens de mercado genéricas (studio/1BR em Cincinnati),
não específicas de "quarto individual + banheiro privativo" -- mesmo
critério já aplicado em outras linhas deste dataset (UC Davis, UMN, Utah).

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://www.business.uc.edu/programs-degrees/graduate/specialized-masters/business-analytics.html
  (duração, 40 créditos totais, placement rate 92% em até 3 meses, salário
  médio $92.498, empregadores Google/Disney/Amazon)
- https://www.business.uc.edu/future-students/admissions-requirements/graduate.html
  (GPA competitivo 3.0; GMAT/GRE opcional a partir de Fall 2025; TOEFL 90/
  IELTS 7.0/Duolingo 115; Business Foundations courses como pré-requisito
  condicional; Lindner não oferece isenção de taxa de inscrição)
- Busca por deadlines: internacionais (visto F-1) recomendados a se
  inscrever até 15/out; Rounds 1-2 dão melhor chance de vaga (programa
  pode fechar por lotação); Round 3 (final): 15/mar, decisão 1/abr; MSBA
  presencial só admite para Fall, não para Spring
- Busca por assistantships: candidatos precisam estar admitidos e com taxa
  de matrícula paga antes de se candidatar a posições via Handshake,
  tipicamente concedidas na chegada; Graduate Incentive Award (GIA) e
  Graduate Assistant Scholarship (GAS) cobrem parte/toda a tuition
- shb/grad.uc.edu: taxa oficial 2026-27 do UC Student Health Insurance
  (Aetna), anunciada ago/2026: $1.107,90 (Fall) + $1.571,22 (Spring) =
  $2.679,12/ano letivo
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
    "Nome da Universidade": "U of Cincinnati",
    "Estado": "OH",
    "Cidade": "Cincinnati",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 40,
    "Tuition Total (Residente) USD": 33195,  # tarifa única/flat, igual à internacional (consistente com "Caminho Tuition Residente: Não")
    "Seguro Saúde (programa) USD": 2679,  # UC SHI 2026-27 oficial: 1.107,90 (Fall) + 1.571,22 (Spring)
    "GPA Mínimo": 3.0,  # sem mínimo formal rígido; 3.0 é o patamar competitivo divulgado
    "GMAT/GRE Obrigatório": (
        "Opcional a partir de Fall 2025; avaliado apenas se enviado, "
        "dentro de uma revisão holística da candidatura"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL mín. 90 / IELTS mín. 7.0 / Duolingo mín. 115 / PTE mín. 54"
    ),  # CORRECAO: planilha tinha limiares desatualizados (80/6.5/110)
    "Matérias Pré-Requisito": (
        "Business Foundations courses podem ser exigidos para quem não "
        "tem graduação credenciada AACSB em Business; alguns podem ser "
        "cursados durante o próprio programa"
    ),
    "Deadline Aplicação": (
        "Internacionais (visto F-1): recomendado até 15/out; Rounds 1-2 "
        "dão melhor chance de vaga (programa pode fechar por lotação); "
        "Round 3 (final): 15/mar, decisão 1/abr; só admite para Fall "
        "(não admite para Spring)"
    ),
    "Rounds Aplicação": (
        "3 rounds + admissão rolling; internacionais devem se inscrever "
        "até 15/out para dar tempo ao processo de visto F-1"
    ),
    "Processo TA/RA": (
        "Candidato precisa estar admitido e com a taxa de matrícula paga "
        "antes de poder se candidatar a vagas via Handshake; posições "
        "tipicamente concedidas na chegada, no início do programa"
    ),
    "Graduate Assistantships": (
        "Graduate Incentive Award (GIA) e Graduate Assistant Scholarship "
        "(GAS) -- bolsas financiadas pela universidade que cobrem parte "
        "ou toda a tuition e certas taxas"
    ),
    "Taxa Empregabilidade %": 92,  # CORRECAO: planilha tinha 90; oficial "92% placement rate within 3 months"
    "Salário Médio Pós USD": 92498,
    "Principais Empregadores": "Amazon, Deloitte, P&G, Kroger, Google, Disney",  # mescla da lista já cadastrada com a fonte oficial (Google, Disney)
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
        raise ValueError("Nenhuma linha encontrada para U of Cincinnati (Cincinnati, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Cincinnati, Cincinnati, OH)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
