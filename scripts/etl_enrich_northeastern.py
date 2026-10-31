"""
Data enrichment: atualiza APENAS a linha da 'Northeastern' (D'Amore-
McKim School of Business, Boston, MA) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (damore-mckim.northeastern.edu).

Foco: MS in Business Analytics, trilha CAMPUS-BASED em Boston
(distinta das trilhas Oakland e 100% Online, também oferecidas pela
D'Amore-McKim — não tocadas aqui), 30 créditos (já confirmado), STEM,
modelo de co-op (estágio remunerado opcional, o que estende a duração
full-time de 12 para até 16 meses).

TUITION: como a D'Amore-McKim é privada com tuition flat (já
documentado na planilha: 'Caminho Tuition Residente' = 'N/A (privada)'),
'Custo/Crédito (Intl)' e 'Custo/Crédito (Residente)' foram preenchidos
com o MESMO valor, derivado do total Intl já presente e não alterado
na planilha ($54.480 ÷ 30 créditos = $1.816,00). 'Tuition Total
(Residente)' preenchido com o mesmo total já confirmado para Intl
($54.480), por não haver distinção de residência em tuition privada.

'GMAT/GRE Obrigatório' preenchido: confirmado oficialmente como
test-optional -- GMAT/GRE NÃO são exigidos (envio opcional).

'Documentos Necessários' preenchido com a lista oficial: formulário de
aplicação, personal statement (≤500 palavras), resume, 2 cartas de
recomendação, transcripts acadêmicos de todas as instituições, GMAT/
GRE opcional, exame de proficiência em inglês (se aplicável), cópia de
passaporte/residência (se aplicável).

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com os 3 prazos
oficiais da trilha campus-based (Boston) para entrada em Fall 2026: 1
nov 2025; 1 jan 2026; 1 mai 2026 (esse último específico para
candidatos internacionais, com decisão liberada uma semana após cada
deadline).

'Bolsas Disponíveis' e 'Tipos de Bolsa' preenchidos com CAUTELA: o
único programa de bolsa com valor CONCRETO encontrado é o "Double
Husky Scholarship" (25% de desconto na tuition), mas ele é
especificamente para a trilha "stackable" que combina o MSBA com o
MBA (exige GPA 3.0+) -- não é uma bolsa geral para candidatos externos
ao MBA stacking. A página também menciona genericamente "full
complement of financial assistance" incluindo assistantships, mas sem
detalhamento de valores ou processo -- documentado como tal.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Duração (meses)' NÃO foi alterado -- a página oficial confirma uma
FAIXA de 12-16 meses (dependendo de o aluno fazer ou não o co-op
opcional), consistente com o valor de 12 já presente (base, sem
co-op); manter como estava evita forçar um único número numa faixa
válida. 'Entrevista': a página confirma que "virtual interview (by
invitation only)" é exigida para o MBA, mas NÃO menciona
explicitamente esse requisito para o MSBA -- não usado para não
assumir aplicabilidade indevida. 'Taxa Aplicação USD', 'Fees
Adicionais USD', 'Seguro Saúde (programa) USD', 'Tamanho Turma',
'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Ranking
US News', 'Ranking QS', 'Taxa Empregabilidade %', 'Salário Médio Pós
USD', 'Principais Empregadores' e todos os campos mensais de custo de
vida -- não encontrados nas páginas acessíveis via WebFetch (a página
de Student Finance com tuition detalhada retornou 404; a página do
Graduate Career Center não tem dados quantitativos por programa).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://damore-mckim.northeastern.edu/programs/ms-business-analytics/
  (formato, duração 12-16 meses, créditos, STEM, currículo, GMAT/GRE
  test-optional, deadlines campus-based, documentos exigidos, Double
  Husky Scholarship)
- https://damore-mckim.northeastern.edu/admissions/how-to-apply/
  (confirma que entrevista por convite é exigida para MBA, sem menção
  ao MSBA)
- https://damore-mckim.northeastern.edu/admissions/tuition-aid/
  (sem valores específicos, direciona a Student Finance — página essa
  que retornou 404)
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
    "Nome da Universidade": "Northeastern",
    "Estado": "MA",
    "Cidade": "Boston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1816.00,  # 54.480 / 30 créditos (total já presente, não alterado)
    "Custo/Crédito (Residente) USD": 1816.00,  # privada, tarifa flat (igual ao Intl)
    "Tuition Total (Residente) USD": 54480,  # privada, tarifa flat (igual ao Intl)
    "GMAT/GRE Obrigatório": "Test-optional — GMAT/GRE não são exigidos (envio opcional)",
    "Documentos Necessários": (
        "Formulário de aplicação; personal statement (≤500 "
        "palavras); resume; 2 cartas de recomendação; transcripts "
        "acadêmicos de todas as instituições; GMAT/GRE opcional; "
        "exame de proficiência em inglês (se aplicável); "
        "passaporte/residência (se aplicável)"
    ),
    "Deadline Aplicação": "Trilha campus-based (Boston), entrada Fall 2026: 1 nov 2025; 1 jan 2026; 1 mai 2026 (internacionais)",
    "Rounds Aplicação": "3 rounds: 1 nov 2025; 1 jan 2026; 1 mai 2026 (internacionais) — decisão ~1 semana após cada deadline",
    "Bolsas Disponíveis": "Double Husky Scholarship (25% de desconto) — específico da trilha stackable MSBA+MBA (GPA 3.0+); universidade também menciona 'full complement' de auxílio financeiro (incl. assistantships), sem detalhamento geral",
    "Tipos de Bolsa": "Double Husky Scholarship (25% desconto, apenas via trilha stackable para MBA)",
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
        raise ValueError("Nenhuma linha encontrada para Northeastern (Boston, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Northeastern, Boston, MA)")

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
