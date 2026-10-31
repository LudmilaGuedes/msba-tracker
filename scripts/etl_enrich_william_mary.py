"""
Data enrichment: atualiza APENAS a linha da 'William & Mary (Raymond A.
Mason School of Business)' (Williamsburg, VA) no
msba_programs_database.xlsx.

Foco exclusivo: Residential MSBA (MS in Business Analytics presencial,
DIFERENTE do "Online MSBA" da mesma escola -- confundido em varias
fontes de busca, filtrado aqui), STEM, 10 meses, inicio em Fall
(agosto). Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
ingles, e custo de vida em Williamsburg, VA.

ACHADO DA AUDITORIA: 'Duracao (meses)' estava 12; a pagina oficial de
admissoes do RESIDENTIAL MSBA especifica claramente "Duration: 10
months". CORRIGIDO para 10.

TUITION: a pagina oficial de tuition do RESIDENTIAL MSBA (a mais
recente encontrada, ciclo 2024-2025 -- nenhuma atualizacao para 2026-27
foi localizada apesar de buscas direcionadas; pode estar desatualizado)
mostra: Residente VA $41.675 tuition + $7.483 fees = $49.158 total;
Nao-residente/Intl $49.432 tuition + $8.053 fees = $57.485 total, mais
$165 (taxa unica para novo aluno internacional) e $85/semestre (taxa de
servico para internacionais). PREENCHIDO 'Tuition Total' com a tuition
pura ($41.675 residente / $49.432 Intl) e 'Fees Adicionais' com o
cenario internacional completo: $8.053 (fees gerais) + $165 (taxa unica
intl) + $170 (2 semestres x $85 de service fee intl) = $8.388.
'Custo/Credito' permanece 'Verificar' -- o total de creditos do
programa nao foi confirmado em nenhuma fonte oficial encontrada.

'Seguro Saude (programa) USD' preenchido com $3.632/ano -- premio
oficial CORRENTE (2026-27) do W&M Student Insurance Plan, cobertura de
12 meses (ago-jul), obrigatorio para todos os alunos em tempo integral
e para internacionais com visto F-1/J-1 (sem opcao de dispensa sem
comprovar cobertura equivalente).

'Transporte/mes' preenchido com $0 -- o WATA (sistema de onibus de
Williamsburg) e GRATUITO para alunos da W&M mediante apresentacao do ID
estudantil (12 rotas pela regiao).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://mason.wm.edu/graduate/msba/residential/admissions/
  tuition-financial-aid/ (tuition oficial do Residential MSBA, fees
  especificas para internacionais, Merit Scholarship, Graduate
  Assistantship 5-10h/semana, Alumni Referral Scholarship $5.000)
- https://mason.wm.edu/graduate/msba/residential/admissions/ (GPA medio
  3.5, waiver de GRE/GMAT com GPA>=3.2, TOEFL 100/IELTS 7.0, pre-
  requisitos de Estatistica + fundamentos de negocios, documentos,
  entrevista obrigatoria, deadlines de 4 rounds, duracao 10 meses,
  inicio Fall)
- Busca web: tamanho de turma 85 alunos, taxa de emprego full-time 83%,
  salario medio $85.000, empregadores (Deloitte, IBM, Accenture, KPMG)
- Busca web: WATA gratuito para alunos W&M; seguro saude $3.632/ano
  (oficial 2026-27); aluguel de quarto privativo em Williamsburg, VA
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
    "Nome da Universidade": "William & Mary (Raymond A. Mason School of Business)",
    "Estado": "VA",
    "Cidade": "Williamsburg",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 10,  # CORREÇÃO: planilha tinha 12; oficial confirma 10 meses
    "Início": "Fall (agosto)",
    "Tamanho Turma": 85,
    "Tuition Total (Residente) USD": 41675,  # oficial (ciclo mais recente encontrado, 2024-25 — sem atualização 2026-27 localizada)
    "Tuition Total (Intl) USD": 49432,  # oficial (mesma ressalva de ciclo)
    "Fees Adicionais USD": 8388,  # cenário internacional: $8.053 (fees gerais) + $165 (taxa única intl) + $170 (2x $85 service fee intl)
    "Seguro Saúde (programa) USD": 3632,  # oficial 2026-27, cobertura 12 meses, obrigatório
    "Aluguel Privativo/mês USD": 700,  # quarto privativo, Williamsburg VA (faixa de mercado encontrada)
    "Transporte/mês USD": 0,  # WATA gratuito para alunos W&M com ID estudantil
    "Tipos de Bolsa": (
        "Mérito; Graduate Assistantships (5–10h/semana, pagas por "
        "hora); Alumni Referral Scholarship ($5.000 na aceitação, para "
        "candidatos indicados por ex-alunos, limitado a 3 indicações/"
        "ano por ex-aluno)"
    ),
    "Processo TA/RA": (
        "Aberto a novos alunos; 5-10h/semana, pagamento por hora "
        "(valor não divulgado publicamente); processo específico de "
        "candidatura não detalhado"
    ),
    "GPA Médio Admitidos": 3.5,
    "GMAT/GRE Obrigatório": "Dispensável para candidatos com GPA ≥ 3.2; caso contrário, recomendado (sem nota mínima fixa)",
    "Testes Inglês Aceitos": (
        "TOEFL mín. 100 (iBT) / IELTS mín. 7.0 — dispensa para quem "
        "tem 4+ anos de graduação em país/instituição de língua "
        "inglesa (código TOEFL: B308)"
    ),
    "Matérias Pré-Requisito": (
        "Estatística em nível universitário (últimos 5 anos) + "
        "fundamentos de negócios (diploma/major/minor em business OU "
        "4 cursos Coursera: Finance, Accounting, Marketing, Operations "
        "Management)"
    ),
    "Experiência Profissional": (
        "Não obrigatória, mas experiência prévia em data analytics "
        "reforça a candidatura; turma atual tem média de 2 anos de "
        "experiência profissional"
    ),
    "Documentos Necessários": (
        "Application online, histórico acadêmico não-oficial (oficial "
        "após admissão), resume (2 páginas preferencialmente), ensaio "
        "(~800 palavras), 2 cartas de recomendação (acadêmica/"
        "profissional), entrevista"
    ),
    "Entrevista": "Obrigatória (presencial ou virtual, por convite)",
    "Deadline Aplicação": (
        "Round 1: 1º set (decisão 20 set); Round 2: 1º dez (decisão 20 "
        "dez); Round 3: 1º fev (decisão 20 fev); Round 4: 15 abr "
        "(decisão 5 mai)"
    ),
    "Rounds Aplicação": "4 rounds (admissão contínua/rolling); ver Deadline Aplicação",
    "Taxa Empregabilidade %": 83,  # taxa de oferta de emprego full-time
    "Salário Médio Pós USD": 85000,
    "Principais Empregadores": "Deloitte, IBM, Accenture, KPMG",
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
        raise ValueError("Nenhuma linha encontrada para William & Mary (Williamsburg, VA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (William & Mary, Williamsburg, VA)")

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
