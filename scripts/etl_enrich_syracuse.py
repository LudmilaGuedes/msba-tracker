"""
Data enrichment: atualiza APENAS a linha da 'Syracuse' (Whitman School
of Management, Syracuse, NY) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (whitman.syracuse.edu). Páginas
específicas de custo de vida e seguro saúde para internacionais
(hospedadas em admissions.syr.edu / financialaid.syracuse.edu)
retornaram erro de DNS.

Foco: M.S. in Business Analytics, trilha FULL-TIME on-campus
(distinta da trilha 100% online, também oferecida pela Whitman —
não tocada aqui), 36 créditos (já confirmado), STEM.

'Duração (meses)' CORRIGIDO de 12 para 16 -- a página oficial do
programa confirma "16 months" para a trilha padrão full-time com
estágio (há também uma trilha acelerada de 16 meses e uma part-time de
21 meses, todas com os mesmos 36 créditos; a trilha padrão de 16 meses
foi usada por ser a mais representativa do formato presencial buscado).

TUITION: como a Whitman é uma escola privada com tuition flat
(já documentado na planilha: 'Caminho Tuition Residente' = 'N/A
(privada)'), 'Custo/Crédito (Residente)' e 'Tuition Total (Residente)'
foram preenchidos com os MESMOS valores já confirmados para Intl
($2.086/crédito; $75.096 total, 36 créditos) -- a própria página
oficial de custos rotula esse valor apenas como "U.S. students" mas,
consistente com o padrão de outras escolas privadas já tratadas nesta
sessão (RIT, Cornell, American University), tuition é uniforme sem
distinção de residência/cidadania em programas privados desta lista.

'GMAT/GRE Obrigatório' refinado: confirmado oficialmente que o teste
NÃO é exigido no ciclo de admissão atual ("not required for the
current admission cycle"), mas o envio é incentivado por poder
melhorar as chances de admissão e de bolsas de mérito.

'Testes Inglês Aceitos' preenchido com os mínimos oficiais: TOEFL iBT
100 (speaking preferencial 23), IELTS Academic 7.0 (speaking
preferencial 7.0), Duolingo 130.

'Deadline Aplicação' preenchido com os 4 prazos oficiais: Priority 15
nov; Early 15 jan; Regular 15 mar; Final 15 abr.

'Principais Empregadores' preenchido com a lista oficial de egressos
citada na página do programa: BAE Systems, Deloitte, EY, Goldman
Sachs, Raytheon Missiles and Defense, T-Mobile.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Mínimo', 'GPA Médio Admitidos', 'GMAT Médio',
'Taxa Aceitação %', 'Taxa Aplicação USD', 'Documentos Necessários'
(a página de admissões apenas referencia um checklist sem detalhar),
'Entrevista', 'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade
%', 'Salário Médio Pós USD', 'Fees Adicionais USD', 'Seguro Saúde
(programa) USD' e todos os campos mensais de custo de vida -- as
páginas de "Cost of Attendance for International Graduate Students"
referenciadas pela própria Whitman (hospedadas em domínios separados
da universidade) retornaram erro de DNS em todas as tentativas via
WebFetch.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://whitman.syracuse.edu/full-time-masters-program/masters-in-business-analytics
  (duração 16 meses, 36 créditos, $2.086/crédito, deadlines, bolsa de
  mérito automática, empregadores)
- https://whitman.syracuse.edu/graduate-programs/full-time-ms-admissions
  (GMAT/GRE não exigido no ciclo atual, mínimos de TOEFL/IELTS/
  Duolingo)
- https://whitman.syracuse.edu/graduate-programs/full-time-ms-admissions/full-time-ms-costs
  (confirma tuition $2.086/crédito "U.S. students"; direciona para
  páginas de custo de internacionais que estavam inacessíveis)
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
    "Nome da Universidade": "Syracuse",
    "Estado": "NY",
    "Cidade": "Syracuse",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 16,  # CORREÇÃO: oficial, trilha full-time padrão com estágio (não 12)
    "Custo/Crédito (Residente) USD": 2086,  # privada, tarifa flat (igual ao Intl)
    "Tuition Total (Residente) USD": 75096,  # privada, tarifa flat (igual ao Intl)
    "GMAT/GRE Obrigatório": "Não exigido no ciclo de admissão atual; envio incentivado (pode melhorar chances de admissão e de bolsa de mérito)",
    "Testes Inglês Aceitos": "TOEFL iBT mín. 100 (speaking pref. 23); IELTS Academic mín. 7.0 (speaking pref. 7.0); Duolingo mín. 130",
    "Deadline Aplicação": "Priority: 15 nov; Early: 15 jan; Regular: 15 mar; Final: 15 abr",
    "Principais Empregadores": "BAE Systems, Deloitte, EY, Goldman Sachs, Raytheon Missiles and Defense, T-Mobile",
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
        raise ValueError("Nenhuma linha encontrada para Syracuse (Syracuse, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Syracuse, Syracuse, NY)")

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
