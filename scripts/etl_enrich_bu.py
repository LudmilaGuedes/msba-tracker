"""
Data enrichment: atualiza APENAS a linha da 'BU' (Boston University,
Questrom School of Business, Boston, MA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM, 36
créditos, formato de 12 ou 16 meses (3 termos consecutivos ou com verão
intercalado para estágio). Deep research em: custo total real de
tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida em Boston, MA.

TUITION: CORRIGIDO. A planilha tinha $89.166 (Intl) sem origem
rastreável. A página oficial de admissões do programa (bu.edu/questrom)
confirma o custo TOTAL do programa (tuition + fees) em $88.539 (trilha
de 12 meses), com um breakdown explícito de $1.823 em fees obrigatórias
(student services, health/wellness, graduate program fee, MSBA program
fee) sobre uma tuition pura de $86.716. CORRIGIDO: 'Tuition Total' para
$86.716 (tuition pura) e 'Fees Adicionais' PREENCHIDO com $1.823. Como
é universidade PRIVADA, 'Custo/Crédito' e os campos '(Residente)'
(estavam 'Verificar') foram preenchidos com os MESMOS valores do Intl
($86.716 / 36 créditos = $2.409/crédito).

'Taxa Empregabilidade %' (95%) e 'Salário Médio Pós USD' ($93.000,
salário-base médio) vêm do MSBA Career Report da Turma de 2022
(questromworld.bu.edu) -- É O ÚNICO RELATÓRIO DE OUTCOMES ESPECÍFICO DO
MSBA ENCONTRADO, mas tem ~3 anos; pode não refletir a turma mais
recente. 'Principais Empregadores' preenchido só com Amazon
(confirmado especificamente para MSBA no mesmo relatório) -- outros
nomes encontrados (Deloitte, PwC, Accenture) eram do Questrom em geral
(todos os programas), não confirmados como especificos do MSBA, e por
isso NÃO foram incluídos, para não misturar outcomes de programas
diferentes.

Custo de vida: 'Aluguel Privativo/mês' preenchido com $1.100 (quarto
privativo mobiliado em Allston/Brighton, bairro popular entre alunos da
BU, na linha Green Line do T; múltiplos anúncios convergentes, muitos
com utilities incluídas). 'Transporte/mês' preenchido com $90 (passe
mensal padrão do MBTA/"the T") -- nenhuma evidência de transporte
gratuito para alunos da BU foi encontrada (diferente de outras
universidades já enriquecidas nesta planilha).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.bu.edu/questrom/graduate-programs/specialty-masters-
  programs/ms-in-business-analytics/admissions/ (tuition oficial
  detalhada por termo e fees, GMAT/GRE opcional, TOEFL 90/IELTS 6.5/
  Duolingo 120 -- com dispensa automática para quem tem 3+ anos de
  graduação em instituição de língua inglesa ou 2+ anos de emprego nos
  EUA, taxa de aplicação $125, documentos incl. 3 video essays via Kira
  Talent, entrevista só por convite, deadlines Round 1-5)
- https://www.bu.edu/academics/questrom/programs/ms-in-business-
  analytics/ (curriculo, formato 12/16 meses, concentrações opcionais)
- Busca web: MSBA Career Report Turma de 2022 (questromworld.bu.edu) --
  taxa de emprego 95%, salário-base médio $93K, compensação total média
  $101K, Amazon como empregador confirmado
- Busca web: aluguel de quarto privativo em Allston/Brighton, Boston MA
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
    "Nome da Universidade": "BU",
    "Estado": "MA",
    "Cidade": "Boston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2409,  # $86.716 (tuition pura, oficial) / 36 créditos
    "Custo/Crédito (Residente) USD": 2409,  # privada, mesma taxa do Intl
    "Tuition Total (Intl) USD": 86716,  # CORREÇÃO: planilha tinha $89.166 sem origem rastreável; oficial = tuition pura ($88.539 total - $1.823 fees)
    "Tuition Total (Residente) USD": 86716,  # privada, mesma taxa do Intl
    "Fees Adicionais USD": 1823,  # oficial: student services, health/wellness, graduate program fee, MSBA program fee
    "Aluguel Privativo/mês USD": 1100,  # quarto privativo mobiliado, Allston/Brighton (perto do campus, linha Green do T)
    "Transporte/mês USD": 90,  # passe mensal padrão do MBTA/"the T"
    "GMAT/GRE Obrigatório": "Opcional (pontuação competitiva reforça a candidatura, mas não é exigida)",
    "Testes Inglês Aceitos": (
        "TOEFL mín. 90 iBT (4.5 na nova escala) / IELTS mín. 6.5 por "
        "banda / Duolingo mín. 120 (110 recomendado) — dispensado com "
        "3+ anos de graduação em instituição de língua inglesa ou 2+ "
        "anos de emprego nos EUA"
    ),
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal específico listado além do diploma de bacharelado",
    "Deadline Aplicação": "Round 1: 13 out; Round 2: 17 nov; Round 3: 7 jan; Round 4: 16 fev; Round 5: 16 mar",
    "Taxa Aplicação USD": 125,
    "Rounds Aplicação": "5 rounds (Round 1 a Round 5, ver Deadline Aplicação); decisão em 6-8 semanas",
    "Documentos Necessários": (
        "Application online, resume (PDF), histórico acadêmico (não-"
        "oficial na submissão, oficial se admitido), 3 video essays "
        "via plataforma Kira Talent (1 tentativa por pergunta, 60s de "
        "resposta), GMAT/GRE opcional, ensaio opcional (máx. 250 "
        "palavras)"
    ),
    "Entrevista": "Não obrigatória — apenas por convite, a critério da admissão",
    "Taxa Empregabilidade %": 95,  # MSBA Career Report, Turma de 2022 (dado mais recente encontrado)
    "Salário Médio Pós USD": 93000,  # salário-base médio, Turma de 2022 (compensação total média: $101K)
    "Principais Empregadores": "Amazon (confirmado específico do MSBA)",
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
        raise ValueError("Nenhuma linha encontrada para BU (Boston, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (BU, Boston, MA)")

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
