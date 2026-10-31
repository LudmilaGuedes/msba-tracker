"""
Data enrichment: atualiza APENAS a linha da 'Notre Dame' (Mendoza College of
Business, Notre Dame, IN) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM. Deep
research em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida local.

AUDITORIA: vários campos já cadastrados foram CONFIRMADOS corretos contra a
página oficial (sem alteração): Tamanho Turma (84, Classe de 2026), Tuition
Total Intl ($69.630, 2026-27), Fees Adicionais ($2.200), Taxa Empregabilidade
(92%, Classe de 2025), Salário Médio Pós ($81.313) e GPA Médio Admitidos
(3.62). Nenhum desses precisou de correção.

Notre Dame é universidade PRIVADA e cobra tarifa única de tuition (sem
distinção residente/não-residente) -- por isso 'Tuition Total (Residente)
USD', que estava "Verificar", foi preenchido com o mesmo valor da tuition
internacional ($69.630), consistente com 'Caminho Tuition Residente' já
cadastrado como "N/A (privada)".

Custo de vida: a página oficial de financial aid da Notre Dame publica um
orçamento anual (Fall/Spring, ano completo) com "Housing and Food" já
somados ($17.820), sem quebrar em aluguel vs. mercado -- por isso os 8
campos mensais individuais do schema (aluguel, mercado, seguro residencial,
utilities, celular, wifi, transporte, outros) foram mantidos "Verificar":
não há como decompor o valor agregado oficial nessas colunas sem inventar
uma proporção que a fonte não fornece. Ainda assim, o TOTAL agregado
("Housing and Food" $17.820 + "Personal Expenses" $6.720 + "Transportation"
$2.700 = $27.240, ano completo = duração do programa) é um número diretamente
publicado pela própria universidade e foi usado para preencher 'Custo Vida
Total (programa) USD'. Seguro-saúde não é item separado no orçamento oficial
("does not separately itemize health insurance"), então 'Seguro Saúde
(programa) USD' permanece "Verificar".

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://financialaid.nd.edu/apply-or-renew/graduate-students/graduate-program-costs/business-analytics-msba/
  (cost of attendance oficial 2026-27, opção Fall Start, ano completo)
- https://mendoza.nd.edu/graduate-programs/business-analytics-msba/
  (rounds de aplicação, class profile Classe de 2026, employment outcomes
  Classe de 2025, principais empregadores, bolsas)
- Busca por requisitos de aplicação MSBA: taxa de inscrição $50 (isenta p/
  Early Action), 2 cartas de recomendação, 1 ensaio (até 500 palavras),
  transcripts, GMAT/GRE (dispensável p/ egressos da própria Notre Dame),
  pré-requisitos de estatística/contabilidade/curso de negócios adicional
- TOEFL/IELTS: mínimos de 80/7.0 confirmados apenas para o MBA da Mendoza
  (não encontrado valor específico publicado para o MSBA -- registrado com
  essa ressalva no campo em vez de apresentado como certeza)
- GMAT médio ~690 / GRE ~320: vindo de agregador terceiro (mim-essay.com),
  não confirmado na página oficial da Mendoza -- registrado com ressalva
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
    "Nome da Universidade": "Notre Dame",
    "Estado": "IN",
    "Cidade": "Notre Dame",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Residente) USD": 69630,  # privada, tarifa única (sem distinção residente/intl) -- mesma da tuition Intl
    "Custo Vida Total (programa) USD": 27240,  # oficial ND: Housing+Food 17.820 + Personal Expenses 6.720 + Transportation 2.700 (ano completo, exclui seguro saúde não itemizado)
    "Principais Empregadores": (
        "Capital One, Deloitte, EY, Honeywell, P&G, Salesforce, The Walt "
        "Disney Company, United Airlines, entre outros em consultoria, "
        "serviços financeiros e tecnologia"
    ),
    "Matérias Pré-Requisito": (
        "Pelo menos 1 curso de estatística, 1 de contabilidade e 1 curso "
        "de negócios adicional (finanças, gestão ou marketing)"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório enviar GMAT ou GRE (sem nota mínima definida "
        "publicada); dispensado para egressos da graduação da própria "
        "Notre Dame"
    ),
    "GMAT Médio": 690,  # aproximado (agregador terceiro, não confirmado na página oficial da Mendoza); GRE equivalente ~320
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/PTE Academic -- mín. TOEFL 80 / IELTS 7.0 "
        "confirmados para o MBA da Mendoza (prática usual da escola); "
        "valor específico do MSBA não encontrado publicado separadamente"
    ),
    "Deadline Aplicação": (
        "Early Action: 10/set/2026 (decisão 16/out, confirmação 16/nov); "
        "Round 1: 20/out/2026; Round 2: 7/jan/2027; Round 3: 17/mar/2027 "
        "-- internacionais devem se inscrever até o Round 2 para dar "
        "tempo ao processo de visto"
    ),
    "Taxa Aplicação USD": 50,  # isenta para candidatos de Early Action
    "Rounds Aplicação": "4 rounds: Early Action + Rounds 1-3 (ver Deadline de Aplicação)",
    "Documentos Necessários": (
        "Transcripts oficiais de todas as instituições (com tradução "
        "para inglês se aplicável), 2 cartas de recomendação, 1 ensaio "
        "(até 500 palavras, tema: planos de carreira pós-mestrado), "
        "GMAT/GRE (dispensável p/ egressos de graduação da Notre Dame), "
        "TOEFL/IELTS/PTE (internacionais), taxa de $50 (isenta p/ Early "
        "Action)"
    ),
    "Entrevista": (
        "Candidatos de Early Action recebem agendamento prioritário de "
        "entrevista (ago-set, por ordem de chegada); não fica explícito "
        "se é obrigatória para os demais rounds"
    ),
    "Tipos de Bolsa": (
        "Mendoza Merit Scholarships (mérito) e Dean's Fellowships "
        "(candidatos excepcionais, pode cobrir parte substancial de "
        "tuition+fees); disponíveis a todos os candidatos qualificados "
        "independente de cidadania; consideração garantida de bolsa de "
        "mérito para candidatos de Early Action"
    ),  # detalhamento: planilha tinha só "Merit (Early Action = garantido)"
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
        raise ValueError("Nenhuma linha encontrada para Notre Dame (Notre Dame, IN). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Notre Dame, Notre Dame, IN)")

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
