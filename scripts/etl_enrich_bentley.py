"""
Data enrichment: atualiza APENAS a linha da 'Bentley University'
(McCallum Graduate School of Business, Waltham, MA) no
msba_programs_database.xlsx.

Foco: MS in Business Analytics, formato presencial full-time (também
part-time, não tocado aqui), 10 cursos mínimos = 30 créditos (1
foundation + 7 obrigatórios + 3 eletivos, confirmado oficialmente;
consistente com a nota já presente "10-16 cursos (c/ waivers)" -- o
teto de 16 se aplica a quem precisa de cursos de nivelamento
adicionais, não tocado por já estar corretamente enquadrado como
faixa), STEM.

TUITION: CORREÇÃO SIGNIFICATIVA. A página oficial de tuition 2026-27
confirma que a Bentley fatura POR CURSO (não por crédito): "$6.335 per
course". Para os 10 cursos mínimos do MS (= 30 créditos): $6.335 × 10
= $63.350 -- MAIOR que o valor anterior ($50.310), que não batia com
essa tarifa oficial nem com nenhum outro cálculo óbvio. CORRIGIDO
'Tuition Total' (Intl e Residente, mesma tarifa -- privada, sem
distinção de residência) para $63.350. 'Custo/Crédito' preenchido
dividindo a tarifa por curso pelos 3 créditos/curso (30 créditos ÷ 10
cursos = 3): $6.335 ÷ 3 = $2.111,67. NOTA: quem precisar de até 16
cursos (sem waivers de nivelamento) pagaria proporcionalmente mais
(até $101.360) -- o valor usado aqui reflete o cenário de 10 cursos
(mínimo, com waivers), consistente com o texto já presente na
planilha enquadrando isso como uma faixa.

'Seguro Saúde (programa) USD' preenchido: $2.650/ano -- prêmio oficial
de seguro saúde, obrigatório para todo aluno de pós-graduação
matriculado em mais de 6 créditos.

'Taxa Aplicação USD' preenchido: $133 (com descontos disponíveis, não
quantificados). 'GMAT/GRE Obrigatório' preenchido: opcional --
aceita-se GMAT, GMAT Focus Edition, GRE, ou Executive Assessment,
nenhum obrigatório. 'Documentos Necessários' preenchido: resume atual
com datas e responsabilidades de emprego; 1 carta de recomendação;
essay; transcripts (não-oficiais aceitos para revisão inicial); teste
de inglês (TOEFL/IELTS/Duolingo) + avaliação de credencial de agência
NACES para internacionais.

'Entrevista' preenchido com INFERÊNCIA CAUTELOSA: a página oficial
declara que "Full-Time MBA and MSHFID applicants are also required to
attend an evaluative interview" -- o MS in Business Analytics NÃO é
mencionado nessa lista de programas que exigem entrevista, sugerindo
(mas sem confirmação explícita e direta) que NÃO é exigida para o
MSBA. Documentado como inferência, não certeza absoluta.

'Testes Inglês Aceitos' refinado: TOEFL/IELTS/PTE/Duolingo aceitos
para internacionais (Duolingo adicionado à lista já existente);
mínimos numéricos específicos ainda não encontrados, mantido
'(mínimos Verificar)' no valor do campo.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Fees Adicionais USD' (além do seguro saúde já preenchido, não há
outras fees discriminadas nas fontes consultadas), 'GPA Médio
Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Ranking US News',
'Ranking QS', 'Salário Médio Pós USD', 'Principais Empregadores',
'Matérias Pré-Requisito', 'Experiência Profissional', 'Caminho
Tuition Residente', 'Processo TA/RA', 'TA/RA Disponível', 'Bolsas
Disponíveis'/'Tipos de Bolsa' (a página menciona "Dean's Leadership
scholarships" mas especifica que são majoritariamente para candidatos
de MBA, não confirmado para o MSBA -- não usado para não sugerir
elegibilidade indevida), e todos os campos mensais de custo de vida
em Waltham, MA.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.bentley.edu/academics/graduate-programs/masters-business-analytics
  (formato, 10 cursos confirmados, STEM/OPT, currículo, tamanho de
  turma qualitativo)
- https://www.bentley.edu/graduate/tuition-aid (tuition oficial
  2026-27 por curso, seguro saúde obrigatório)
- https://www.bentley.edu/graduate/apply/deadlines (taxa de
  aplicação $133, GMAT/GRE opcional, documentos, entrevista restrita
  a MBA/MSHFID — MSBA não mencionado)
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
    "Nome da Universidade": "Bentley University",
    "Estado": "MA",
    "Cidade": "Waltham",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2111.67,  # $6.335/curso ÷ 3 créditos/curso, oficial 2026-27
    "Custo/Crédito (Residente) USD": 2111.67,  # privada, tarifa flat (igual ao Intl)
    "Tuition Total (Intl) USD": 63350,  # CORREÇÃO: 10 cursos x $6.335, oficial (era 50.310, sem fonte clara)
    "Tuition Total (Residente) USD": 63350,  # privada, tarifa flat (igual ao Intl)
    "Seguro Saúde (programa) USD": 2650,  # oficial anual, obrigatório para >6 créditos
    "Taxa Aplicação USD": 133,  # descontos disponíveis, não quantificados
    "GMAT/GRE Obrigatório": "Opcional (GMAT, GMAT Focus Edition, GRE, ou Executive Assessment aceitos, nenhum obrigatório)",
    "Documentos Necessários": (
        "Resume atual com datas/responsabilidades de emprego; 1 "
        "carta de recomendação; essay; transcripts (não-oficiais "
        "aceitos para revisão inicial); teste de inglês (TOEFL/"
        "IELTS/Duolingo) + avaliação de credencial NACES "
        "(internacionais)"
    ),
    "Entrevista": (
        "Provavelmente não exigida para o MSBA — a fonte oficial "
        "lista apenas Full-Time MBA e MSHFID como programas com "
        "entrevista avaliativa obrigatória (inferência, sem "
        "confirmação explícita)"
    ),
    "Testes Inglês Aceitos": "IELTS/TOEFL/PTE/Duolingo aceitos (mínimos numéricos específicos ainda não confirmados)",
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
        raise ValueError("Nenhuma linha encontrada para Bentley University (Waltham, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Bentley University, Waltham, MA)")

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
