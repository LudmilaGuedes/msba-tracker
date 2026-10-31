"""
Data enrichment: atualiza APENAS a linha da 'University of Dallas' (Satish
& Yasmin Gupta College of Business, Irving, TX) no msba_programs_database.
xlsx.

Foco exclusivo: Master of Science in Business Analytics, presencial, STEM,
30 creditos (13 disciplinas). Deep research em: custo total real de
tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/ingles, e custo de vida em Irving, TX. ATENCAO: "University of
Dallas" (privada, Irving) e uma instituicao DIFERENTE de "University of
Texas at Dallas / UT Dallas" (publica, Richardson, ja cadastrada em outra
linha desta planilha) -- varias buscas trazem resultados misturados dos
dois nomes; todos os dados usados aqui foram verificados como vindos
especificamente de paginas oficiais em udallas.edu.

AUDITORIA DE TUITION (ja estava correto, apenas verificado): $1.300/
credito (igual para todos, sem distincao de residencia -- universidade
privada), $39.000 total para 30 creditos, $385 de fees adicionais (=
matriculation fee $160 + graduation fee $100 + parking fee $125,
confirmado oficialmente). Nenhuma correcao necessaria nesses campos.

'Caminho Tuition Residente': preenchido com "Nao aplicavel" -- sendo
universidade PRIVADA, nao ha distincao de tuition por residencia (todos
pagam a mesma taxa), entao o conceito de "caminho para tuition de
residente" nao se aplica aqui.

'Duracao (meses)' permanece 'Verificar': a pagina oficial descreve o
programa como "1-2 anos, dependendo do calendario e carga de materias",
sem indicar um numero especifico de meses para o formato presencial em
tempo integral (o programa e desenhado com flexibilidade evening/online/
hibrido para profissionais, entao a duracao real varia bastante por
aluno). Nao ha uma fonte que decida um numero unico com confianca.

'Materias Pre-Requisito' permanece 'Verificar': a pagina oficial de
requisitos lista pre-requisitos especificos para as trilhas de
Accounting/Finance/Financial Planning (ex.: Intermediate Financial
Accounting, Python para Finance), mas NAO confirma se a trilha de
Business Analytics tem pre-requisito proprio -- nao assumido sem
confirmacao direta.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://udallas.edu/academics/programs/business-analytics/index.php
  (formato, STEM/OPT 36 meses, 30 creditos/13 materias, GMAT/GRE nao
  exigido)
- https://udallas.edu/admissions-aid/graduate/college-of-business/
  ms-mba-admission-requirements.php e .../ms-mba-application-process.php
  (documentos, taxa de aplicacao $50, GMAT/GRE opcional, personal
  statement sobre "virtuous leadership")
- Busca web confirmando TOEFL 80 / IELTS 6.5 / Duolingo 110 / PTE 67 /
  MET 55 / iTEP 3.9 para vistos sem diploma dos EUA (paginas udallas.edu
  de international admissions)
- https://udallas.edu/academics/colleges-schools/college-of-business/
  students/scholarships/index.php (Gupta College Graduate Scholarship:
  20% off, $1.300->$1.040/credito, consideracao automatica)
- https://udallas.edu/academics/colleges-schools/college-of-business/
  students/studentworker.php (Graduate Assistant: $13/h, ate 20h/semana,
  termo de 12 semanas, GPA minimo 3.0, elegibilidade F-1 com restricoes)
- Busca web: aluguel de quarto privativo com banheiro privativo em
  Irving/Las Colinas, TX -- multiplos anuncios convergindo em ~$700-800/
  mes
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
    "Nome da Universidade": "University of Dallas",
    "Estado": "TX",
    "Cidade": "Irving",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Início": "Fall, Spring, Summer",  # admissão contínua, com deadlines de bolsa alinhados aos 3 semestres
    "Aluguel Privativo/mês USD": 800,  # quarto privativo com banheiro privativo, Irving/Las Colinas TX (múltiplos anúncios convergentes)
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Gupta College of Business Graduate Scholarship (20% de desconto "
        "por crédito, de $1.300 para $1.040/crédito; consideração "
        "automática para todo admitido, sem GPA/teste mínimo divulgado, "
        "mas seletiva); demais bolsas nomeadas do College majoritariamente "
        "restritas a MBA e/ou cidadãos dos EUA"
    ),
    "Quão Realista (Intl)": (
        "Moderada — a bolsa principal (Gupta Scholarship, 20% off) "
        "considera automaticamente todo admitido sem restrição a "
        "internacionais, mas é seletiva ('muitos são considerados, "
        "poucos selecionados'); as demais bolsas nomeadas exigem "
        "cidadania dos EUA"
    ),
    "TA/RA Disponível": "Sim (Graduate Assistant, vagas limitadas por departamento)",
    "Processo TA/RA": (
        "Candidatura por e-mail com currículo (sem processo formal "
        "centralizado); preenchimento conforme necessidade "
        "departamental; exige mínimo de 6h já cursadas na UD e GPA ≥ 3.0"
    ),
    "Caminho Tuition Residente": "Não aplicável — universidade privada, tuition igual para todos independente de residência",
    "Graduate Assistantships": (
        "Pagamento de $13/hora, até 20h/semana, termo de 12 semanas; sem "
        "menção de waiver de tuition; alunos F-1 podem participar, com "
        "restrições (não podem trabalhar em trimestres de férias sem "
        "matrícula)"
    ),
    "Experiência Profissional": "Não obrigatória (Catalyst Application disponível para profissionais experientes sem diploma prévio)",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 80 / IELTS Academic mín. 6.5 / Duolingo mín. 110 "
        "/ PTE Academic mín. 67 / MET mín. 55 / iTEP mín. 3.9 — exigido "
        "apenas para portadores de visto sem diploma dos EUA"
    ),
    "Taxa Aplicação USD": 50,
    "Rounds Aplicação": "Admissão contínua (rolling); sem rounds numerados",
    "Deadline Aplicação": "Rolling; prioridade de bolsa: 1º jul (Fall), 1º dez (Spring), 1º abr (Summer)",
    "Documentos Necessários": (
        "Resume/CV, histórico acadêmico (unofficial na aplicação; oficial "
        "se admitido), passaporte (internacional); opcionais para "
        "reforçar a candidatura: personal statement (500–1.000 palavras, "
        "tema 'virtuous leadership'), carta de referência, GMAT/GRE"
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
        raise ValueError("Nenhuma linha encontrada para University of Dallas (Irving, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Dallas, Irving, TX)")

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
