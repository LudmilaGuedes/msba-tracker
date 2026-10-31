"""
Data enrichment: atualiza APENAS a linha da 'U of Arizona' (Eller
College of Management, Tucson, AZ) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, trilha FULL-TIME de 10 meses (início
Verão, cursos de nivelamento online + core presencial em Tucson --
distinta da trilha de 16 meses e do campus Chandler, nenhuma das duas
tocada aqui), 33 créditos (já confirmado), STEM.

TUITION: os valores já presentes ($33.800 residente / $56.100 Intl)
foram REFINADOS com as cifras exatas da página oficial de custos:
$33.845,50 (residente) / $56.167,50 (internacional) -- pequena
correção de precisão (diferença de ~$45-67, provavelmente
arredondamento na rodada anterior de coleta). 'Custo/Crédito' derivado
dividindo pelos 33 créditos: $1.025,62 (residente) / $1.702,05
(internacional).

'Fees Adicionais USD' permanece 'Verificar' DELIBERADAMENTE: a página
oficial de custos afirma explicitamente que "Program fees are included
in tuition estimates" (ou seja, já embutidas nos totais acima) e cita
apenas uma faixa vaga ("$3.834–$5.750 dependendo do semestre") sem
valor fechado -- preencher esse campo separadamente somando essa faixa
ao total já preenchido DUPLICARIA a cobrança, já que ela já está
embutida. Documentado para não gerar dupla contagem.

'GPA Mínimo' preenchido: 3.00 (graduação); candidatos internacionais
devem usar a "UA Graduate College GPA Calculator" para determinar o
GPA de entrada equivalente. 'Testes Inglês Aceitos' refinado com
mínimos oficiais: TOEFL 85 iBT (antes de 21/jan/26) ou 4.5 (novo
formato, depois); IELTS composto 7 (nenhuma sub-área abaixo de 6);
PTE não confirmado nesta fonte (mantido na lista por já estar presente
na planilha, mas sem mínimo específico localizado).

'Documentos Necessários' preenchido com a lista oficial: transcripts
oficiais; resume; statement of purpose; 2-3 cartas de recomendação de
associados profissionais; Graduate College Application; TOEFL/IELTS
(internacionais). 'Entrevista' preenchido/refinado: entrevista virtual
de 15-25 minutos pode ser solicitada à maioria dos candidatos, usada
para avaliar preparo quantitativo -- SUBSTITUI o GMAT/GRE (já
registrado no campo 'GMAT/GRE Obrigatório', não alterado por não ser
'Verificar').

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com os 5 prazos
oficiais (ciclo 2026): Round 1 (19 out); Round 2 (14 dez); Round 3 +
deadline internacional (1 fev); Round 4 (31 mar); deadline doméstico
final (14 jun).

'Processo TA/RA' preenchido: candidatos devem já estar admitidos ao
programa; posições limitadas por ano; benefício = stipend + tuition
parcial (valores exatos não divulgados); responsabilidades incluem
correção de provas de graduação, projetos departamentais, ou operações
do programa.

'Salário Médio Pós USD' preenchido com a FAIXA oficial encontrada
($65.000–$105.000, dependendo de experiência e localização) -- não
foi convertido para um único valor médio/mediano porque a fonte não
fornece essa cifra central, apenas a faixa; inventar um ponto médio
seria uma estimativa não confirmada. 'Principais Empregadores'
preenchido com uma amostra da lista oficial: Deloitte, Amazon, KPMG,
EY, Honeywell, Walmart, Intuit, Carvana, DHL.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio' (não aplicável --
o programa não usa GMAT/GRE, usa entrevista), 'Taxa Aceitação %',
'Ranking US News', 'Ranking QS' (a página de careers confirma
acreditação AACSB mas não cita ranking numérico específico), 'Taxa
Aplicação USD', 'Matérias Pré-Requisito' (nenhum pré-requisito formal
de disciplinas confirmado -- programação e diploma em business não
são exigidos), 'Seguro Saúde (programa) USD', 'Transporte/mês USD',
'Utilities/mês USD', 'Custo Vida Total (programa) USD'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://eller.arizona.edu/programs/masters/business-analytics
  (formato, duração, STEM, currículo, tuition, deadlines das 5 rounds,
  entrevista substitui GMAT/GRE)
- .../admissions (GPA, mínimos de inglês, documentos, pré-requisitos)
- .../roi (bolsas de mérito automáticas, GA sem valor fechado,
  confirmação de que fees já estão embutidas no total de tuition)
- .../careers (90% empregabilidade — já presente —, faixa salarial
  $65k-$105k, empregadores)
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
    "Nome da Universidade": "U of Arizona",
    "Estado": "AZ",
    "Cidade": "Tucson",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1702.05,  # 56.167,50 / 33, oficial (refinado)
    "Custo/Crédito (Residente) USD": 1025.62,  # 33.845,50 / 33, oficial (refinado)
    "Tuition Total (Intl) USD": 56167.50,  # refinado (era 56.100)
    "Tuition Total (Residente) USD": 33845.50,  # refinado (era 33.800)
    "GPA Mínimo": "3.00 (graduação); internacionais usam a UA Graduate College GPA Calculator para GPA de entrada equivalente",
    "Testes Inglês Aceitos": "TOEFL 85 iBT (antes de 21/jan/26) ou 4.5 (novo formato, depois); IELTS composto 7 (nenhuma sub-área <6); PTE não confirmado nesta fonte",
    "Documentos Necessários": (
        "Transcripts oficiais; resume; statement of purpose; 2-3 "
        "cartas de recomendação de associados profissionais; "
        "Graduate College Application; TOEFL/IELTS (internacionais)"
    ),
    "Entrevista": "Virtual, 15-25 min, pode ser solicitada à maioria dos candidatos; substitui GMAT/GRE, avalia preparo quantitativo",
    "Deadline Aplicação": "Round 1: 19 out; Round 2: 14 dez; Round 3 + deadline internacional: 1 fev; Round 4: 31 mar; deadline doméstico final: 14 jun",
    "Rounds Aplicação": "5 rounds/deadlines: 19 out, 14 dez, 1 fev (+ final intl), 31 mar, 14 jun (final doméstico)",
    "Processo TA/RA": (
        "Requer admissão prévia ao programa; posições limitadas por "
        "ano; benefício = stipend + tuition parcial (valores exatos "
        "não divulgados); responsabilidades: correção de provas de "
        "graduação, projetos departamentais, ou operações do programa"
    ),
    "Salário Médio Pós USD": "Faixa oficial: $65.000–$105.000 (dependendo de experiência e localização; sem cifra central divulgada)",
    "Principais Empregadores": "Deloitte, Amazon, KPMG, EY, Honeywell, Walmart, Intuit, Carvana, DHL",
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
        raise ValueError("Nenhuma linha encontrada para U of Arizona (Tucson, AZ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Arizona, Tucson, AZ)")

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
