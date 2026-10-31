"""
Data enrichment: atualiza APENAS a linha da 'UAH' (University of Alabama
in Huntsville, College of Business, Huntsville, AL) no
msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial (face-to-face,
evening), 30 créditos (+ prerequisito de Statistics se necessário),
STEM, 12 meses (tempo integral). Deep research em: custo total real de
tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida em Huntsville, AL.

TUITION: a calculadora oficial de tuition do College of Business
(uah.edu/business/grad/graduate-tuition, ano letivo 2026-27) fornece um
breakdown line-item para o formato campus/híbrido: Tuition $1.730/
crédito + Program Fee $87/crédito + College of Business Fee $27/crédito
+ Infrastructure Fee $22/crédito = $1.866/crédito TOTAL -- e, de forma
notável, esse valor é O MESMO para residentes e não-residentes no
formato campus (ao contrário da tabela geral de tuition da
universidade, que diferencia por residência só na tarifa institucional
base). PREENCHIDO: 'Custo/Crédito' com a tuition pura ($1.730, igual
para os dois) e 'Fees Adicionais' com a soma das 3 taxas por crédito
($136/crédito x 30 = $4.080) -- não inclui a International Fee
($150/semestre Fall/Spring, ~$25-50 no verão), que é um custo à parte,
específico de status migratório, não uma fee institucional geral.

Custo de vida: nenhuma fonte confiável com granularidade de "quarto
individual + banheiro privativo" foi encontrada para Huntsville, AL --
apenas médias de domicílio inteiro (ex.: aluguel médio $1.072/mês,
utilities $780/mês) sem uma forma segura de ratear para um único
quarto, e nenhuma evidência de transporte público gratuito para
estudantes. Por isso TODOS os campos de custo de vida mensal permanecem
'Verificar', e por consequência 'Custo Vida Total (programa) USD'
também.

'GPA Médio Admitidos' (3.36) e 'GMAT Médio' (537,5) vêm do perfil da
turma de 2018 divulgado no snapshot oficial do programa -- É O ÚNICO
DADO DE PERFIL DE TURMA ENCONTRADO, mas tem quase 8 anos; pode não
refletir o perfil atual. Preenchidos com essa ressalva documentada.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.uah.edu/business/grad/degrees/analytics/program-snapshot
  (30 créditos, 12 meses tempo integral, formato face-to-face evening,
  empregadores reais do programa, perfil da turma de 2018)
- https://www.uah.edu/business/grad/graduate-tuition (calculadora
  oficial de tuition 2026-27 do College of Business, breakdown completo
  por crédito)
- https://www.uah.edu/bursar/tuition/tuition-rates-2 (tarifas
  institucionais gerais 2025-26 e International Fee)
- https://www.uah.edu/admissions/graduate/apply-for-admission/
  international-students (TOEFL 79/IELTS 6.0/PTE 59/Duolingo 110,
  certificação financeira, taxa de aplicação $60, deadlines de I-20)
- https://www.uah.edu/admissions/graduate/apply-for-admission/
  test-score-requirements (GMAT mín. 500 / GRE mín. 300+3.0, critérios
  de waiver)
- https://www.uah.edu/graduate/faqs (Graduate Assistantship: tempo
  integral 9-12h/20h semana cobre tuition completa + seguro saúde
  individual + stipend + fees)
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
    "Nome da Universidade": "UAH",
    "Estado": "AL",
    "Cidade": "Huntsville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,
    "Total Créditos": 30,
    "Custo/Crédito (Intl) USD": 1730,  # oficial 2026-27, formato campus — igual ao residente
    "Custo/Crédito (Residente) USD": 1730,  # oficial 2026-27, formato campus — igual ao internacional
    "Tuition Total (Intl) USD": 51900,  # 30 créditos x $1.730
    "Tuition Total (Residente) USD": 51900,  # 30 créditos x $1.730
    "Fees Adicionais USD": 4080,  # (Program Fee $87 + College Fee $27 + Infrastructure $22)/crédito x 30 créditos
    "Caminho Tuition Residente": (
        "Flat — no formato campus/híbrido, a tuition é igual para "
        "residentes e não-residentes (calculadora oficial 2026-27 do "
        "College of Business)"
    ),
    "Processo TA/RA": (
        "Aberto a qualquer aluno admitido ao Graduate School; "
        "concedido por departamento/programa. Assistantship de tempo "
        "integral (9-12h matriculadas, 20h/semana de trabalho) cobre "
        "tipicamente tuition completa, seguro saúde (individual), "
        "stipend e fees associadas"
    ),
    "Graduate Assistantships": (
        "GTA/GRA disponíveis; tempo integral (20h/semana) cobre "
        "tuition completa + seguro saúde individual + stipend + fees"
    ),
    "GPA Mínimo": 3.0,
    "GPA Médio Admitidos": 3.36,  # turma de 2018 (único dado de perfil encontrado; pode estar desatualizado)
    "GMAT/GRE Obrigatório": (
        "GMAT (mín. 500) ou GRE (mín. 300 verbal+quant, 3.0 analytical "
        "writing); waiver possível com GPA 3.0 + 3 anos de experiência "
        "profissional ou pós-graduação já concluída"
    ),
    "GMAT Médio": 537.5,  # turma de 2018 (único dado de perfil encontrado; pode estar desatualizado)
    "Matérias Pré-Requisito": "Curso de Statistics como pré-requisito, se ainda não cursado na graduação",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 79 (escala antiga) ou 4 (nova escala) / IELTS "
        "mín. 6.0 / PTE Academic mín. 59 / Duolingo mín. 110 — "
        "dispensado com diploma dos EUA"
    ),
    "Deadline Aplicação": (
        "Doméstico: Fall 1º jun, Spring 30 nov, Summer 15 mar; "
        "Internacional (I-20): Fall 1º ago, Spring 1º dez — "
        "assistantship pode exigir prazo mais cedo"
    ),
    "Taxa Aplicação USD": 60,
    "Rounds Aplicação": "Sem rounds numerados; admissão por deadline de semestre (ver Deadline Aplicação)",
    "Documentos Necessários": (
        "Histórico acadêmico (oficial ou não-oficial via Charger "
        "Status), resume, GMAT/GRE (se exigido/não dispensado), "
        "certificação financeira (internacionais)"
    ),
    "Principais Empregadores": "ADTRAN, Aerojet Rocketdyne, Cigna, Deloitte, DIRECTV, Dynetics, NASA, SAIC",
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
        raise ValueError("Nenhuma linha encontrada para UAH (Huntsville, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UAH, Huntsville, AL)")

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
