"""
Data enrichment: atualiza APENAS a linha da 'ASU' (W.P. Carey School of
Business, Tempe, AZ) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, trilha "Big Data" (9 meses, presencial,
início em agosto) -- já confirmado na planilha via 'Duração (meses)' =
9 e 'Destaques Currículo' = 'Big Data track (9mo) + 4 specialized
(16mo)'. As 4 trilhas especializadas de 16 meses (Cloud/tech
consulting, Fintech, Marketing Analytics, Supply Chain Analytics) NÃO
são tocadas aqui por serem um produto diferente (duração e tuition
distintas). 30 créditos (já confirmado), STEM.

TUITION: CORREÇÃO CRÍTICA. Os valores anteriores ('Tuition Total
(Residente)' = $33.800; '(Intl)' = $61.058) estavam ERRADOS -- o valor
residente ($33.800) era IDÊNTICO ao valor que a U of Arizona (linha
vizinha nesta mesma planilha, também no Arizona) tinha ANTES de ser
corrigido nesta mesma sessão, um forte indício de erro de cópia entre
linhas adjacentes de duas universidades do Arizona. A página oficial
de custos do W.P. Carey confirma valores MUITO diferentes e
específicos por trilha: trilha Big Data (9 meses) -- Residente
$40.000 / Não-residente doméstico $64.000 / INTERNACIONAL $67.500
(ASU distingue "Nonresident" doméstico de "International", com o
internacional sendo ligeiramente mais caro). CORRIGIDO: 'Tuition Total
(Residente)' = $40.000; 'Tuition Total (Intl)' = $67.500 (usando a
coluna "International", a mais relevante para o perfil desta busca).
'Custo/Crédito' derivado dividindo pelos 30 créditos: $1.333,33
(residente) / $2.250,00 (internacional). Depósito de matrícula de $500
(único, não-reembolsável) mencionado mas não incluído em nenhum campo
recorrente do schema.

'GMAT/GRE Obrigatório' refinado: opcional para admissão -- scores
podem fortalecer a candidatura dependendo do histórico acadêmico/
profissional, mas não são exigidos (códigos institucionais: GMAT
0GZ-1R-80, GRE 4007). 'Testes Inglês Aceitos' refinado com uma
exclusão IMPORTANTE: ASU NÃO aceita TOEFL iBT Home Edition (antes de
21/jan/26), TOEFL iBT Paper Edition, nem TOEFL Essentials -- apenas o
formato padrão online é aceito.

'Taxa Aplicação USD' preenchido: $70 (doméstico) / $115
(internacional) -- usado o valor internacional como principal por ser
o foco desta busca; dispensável para militares/AmeriCorps/Peace Corps/
Teach for America. 'Documentos Necessários' preenchido com a lista
oficial: informações pessoais, resume, respostas de ensaio de
admissão (personal statements genéricos NÃO são aceitos, precisa
responder às perguntas específicas), contato de recomendadores,
transcripts oficiais, GMAT/GRE opcional, comprovante de proficiência
em inglês + traduções oficiais para internacionais.

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com o ciclo
relevante para entrada em agosto (Fall 2027, consistente com 'Início'
= 'August' já presente): Round I (1 dez 2026); Round II, prioridade
internacional (1 fev 2027); Round III, final internacional (1 abr
2027); Round IV, final geral (1 jul 2027); deadline rolling até 15 jul
2027.

'Tipos de Bolsa' refinado: bolsas competitivas limitadas baseadas em
desempenho acadêmico/liderança/envolvimento (aplicação exigida após
admissão); parcerias acadêmicas com empregadores oferecem bolsas de
$2.000 a $25.000 (elegibilidade não detalhada -- provavelmente ligada
a empresas parceiras específicas, não uma bolsa geral).

CUIDADO APLICADO: a fonte cita "95% of MS-BA students employed within
6 months", divergindo do valor já presente em 'Taxa Empregabilidade %'
(91%, não 'Verificar', não alterado). Como nenhuma das duas fontes
identifica claramente a coorte/ano de referência, não foi possível
determinar qual é mais atual -- o valor existente foi MANTIDO por
cautela (não é uma correção clara, apenas uma possível divergência de
coorte), e a discrepância fica registrada aqui para revisão futura.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Mínimo', 'GPA Médio Admitidos', 'GMAT Médio',
'Taxa Aceitação %', 'Matérias Pré-Requisito', 'Entrevista' (não
mencionado em nenhuma fonte oficial consultada), 'Ranking US News',
'Ranking QS', 'Fees Adicionais USD' (os valores de tuition já são
descritos como "estimated total program costs", possivelmente já
incluindo fees, sem discriminação separada -- não alterado para não
arriscar dupla contagem), 'Seguro Saúde (programa) USD', 'Processo
TA/RA' (já havia nota geral de "GA after committing", sem processo
detalhado encontrado nesta rodada), e todos os campos mensais de
custo de vida em Tempe, AZ.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://wpcarey.asu.edu/masters-programs/business-analytics
  (formato, duração das duas trilhas, STEM, currículo, tuition por
  trilha e por residência/internacional, GMAT/GRE opcional com
  códigos, exclusões de formato de TOEFL, deadlines completos por
  ciclo, taxas de aplicação, documentos, bolsas, 95% empregabilidade)
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
    "Nome da Universidade": "ASU",
    "Estado": "AZ",
    "Cidade": "Tempe",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2250.00,  # 67.500 / 30, oficial (trilha Big Data)
    "Custo/Crédito (Residente) USD": 1333.33,  # 40.000 / 30, oficial (trilha Big Data)
    "Tuition Total (Intl) USD": 67500,  # CORREÇÃO: oficial (coluna "International"), era 61.058 sem fonte clara
    "Tuition Total (Residente) USD": 40000,  # CORREÇÃO CRÍTICA: era 33.800, valor idêntico ao antigo da U of Arizona vizinha (provável erro de cópia)
    "GMAT/GRE Obrigatório": (
        "Opcional para admissão; scores podem fortalecer a "
        "candidatura dependendo do histórico (códigos "
        "institucionais: GMAT 0GZ-1R-80, GRE 4007)"
    ),
    "Testes Inglês Aceitos": (
        "IELTS/TOEFL/PTE/Duolingo exigidos para não-nativos; ASU NÃO "
        "aceita TOEFL iBT Home Edition (antes de 21/jan/26), TOEFL "
        "iBT Paper Edition, nem TOEFL Essentials"
    ),
    "Taxa Aplicação USD": 115,  # internacional (doméstico: $70); dispensável para militares/AmeriCorps/Peace Corps/Teach for America
    "Documentos Necessários": (
        "Informações pessoais; resume (datas em Mês/Ano); respostas "
        "de ensaio de admissão (personal statement genérico NÃO "
        "aceito); contato de recomendadores; transcripts oficiais; "
        "GMAT/GRE opcional; comprovante de proficiência em inglês + "
        "traduções oficiais (internacionais)"
    ),
    "Deadline Aplicação": "Entrada Fall (agosto): Round I (1 dez 2026); Round II intl prioridade (1 fev 2027); Round III intl final (1 abr 2027); Round IV final (1 jul 2027); rolling até 15 jul 2027",
    "Rounds Aplicação": "4 rounds (1 dez, 1 fev, 1 abr, 1 jul) + rolling até 15 jul",
    "Tipos de Bolsa": (
        "Bolsas competitivas limitadas por desempenho acadêmico/"
        "liderança/envolvimento (aplicação exigida após admissão); "
        "parcerias com empregadores oferecem bolsas de $2.000 a "
        "$25.000 (elegibilidade não detalhada)"
    ),
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
        raise ValueError("Nenhuma linha encontrada para ASU (Tempe, AZ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (ASU, Tempe, AZ)")

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
