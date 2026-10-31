"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'Dakota State'
(College of Business & IS, Madison, SD) no msba_programs_database.xlsx.

Programa: MS in Analytics & Applied Artificial Intelligence (MSAA),
presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-03):
  - https://dsu.edu/admissions/graduate/cost-aid/tuition-fees.html
  - https://dsu.edu/admissions/cost-of-attendance.html
  - https://dsu.edu/programs/msaa/index.html
  - https://dsu.edu/programs/msaa/msaa-timelines-requirements.html
  - https://dsu.edu/admissions/graduate/graduate-assistantships.html
  - https://dsu.edu/admissions/graduate/graduate-scholarships-fellowships.html
  - https://dsu.edu/admissions/international/graduate/index.html
  - https://myahpcare.com/wp-content/uploads/SDBOR_PC-26-27.pdf
    (tabela de prêmios do seguro internacional do SD Board of Regents)

ACHADOS / CORREÇÕES:
- Fees Adicionais: a tabela oficial 2026-27 traz "Est. Masters Tuition
  and Fees" por crédito -- os US$ 698,10 (não-residente) e US$ 393,70
  (residente) JÁ INCLUEM as fees obrigatórias (confirmado pelo Cost of
  Attendance: 12 créditos de mestrado = US$ 8.378 "Tuition & Fees").
  Fees Adicionais = 0 (só taxas de curso/laboratório específicas,
  "additional credit fees may apply to certain courses").
- Seguro Saúde: obrigatório para internacionais (BOR Policy 3.14), plano
  AHP/UnitedHealthcare do SDBOR; prêmio 2026-27 do aluno = US$ 1.422/ano
  (Fall US$ 592,50 + Spring/Summer US$ 829,50).
- Duração: a FAQ oficial diz 1-1,5 ano em full-time (2-3 disciplinas por
  semestre). 30 créditos a 9 créditos/semestre (carga full-time de F-1)
  levam ~4 períodos; com entrada no Fall a formatura cai por volta de
  dezembro do ano seguinte = ~16 meses. O valor anterior (12) era o
  cenário mínimo.
- Formato: o programa é oferecido presencial E online; F-1 e
  assistantships exigem presencial no campus de Madison.
- Início: internacionais só Fall e Spring (Summer só doméstico).
- Currículo: 4 trilhas (Business, Healthcare analytics, Information
  systems, General) e exame abrangente (comprehensive exam) obrigatório.
- Bolsas: não há bolsa institucional de mérito para o mestrado; a página
  de Graduate Scholarships só lista fontes externas (NSF, bases de
  bolsas etc.). O único funding interno é a GA, e a própria DSU avisa:
  "More students apply than there are assistantships available. Please
  make preparation to fund your graduate study that does not rely on
  graduate assistantships alone."
- Documentos: internacionais também precisam do Certificate of
  Financial Responsibility (Masters CFR) e de prova do seguro antes da
  matrícula; a DSU pode pedir avaliação de transcripts por agência.
- Empregadores: a página também cita T-Mobile e Verizon.

NÃO ALTERADOS: Salário Médio Pós (a página só mostra US$ 102.040 como
salário médio de "computer systems analysts" em SD, dado genérico de
ocupação, não dos formados -- permanece 'Verificar'); Taxa
Empregabilidade 92,9% (dado oficial, mas inclui "further education");
Rankings, Tamanho Turma, GPA Médio, Taxa Aceitação.
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
    "Nome da Universidade": "Dakota State",
    "Estado": "SD",
    "Cidade": "Madison",
}

UPDATES = {
    "Formato": (
        "Presencial (campus de Madison); o programa também existe online, "
        "mas visto F-1 e assistantships exigem o presencial"
    ),
    "Duração (meses)": 16,  # 1-1,5 ano oficial; 30 cr. a ~9 cr./semestre (Fall -> dez do ano seguinte)
    "Início": "Fall/Spring para internacionais (Summer só para domésticos)",
    "Destaques Currículo": (
        "30 créditos (18 core + 6 obrigatórios + 6 eletivas) com 4 trilhas "
        "(Business, Healthcare analytics, Information systems, General); "
        "big data com Hadoop, Spark e Kafka, BI e visualização, AI "
        "aplicada; exame abrangente (comprehensive exam) obrigatório; "
        "aulas ao vivo ou gravadas; full-time em 1-1,5 ano"
    ),
    "Fees Adicionais USD": 0,  # fees obrigatórias já incluídas no valor por crédito ("Tuition and Fees")
    "Seguro Saúde (programa) USD": 1422,  # plano AHP/UHC do SDBOR, obrigatório, 2026-27
    "Tipos de Bolsa": (
        "Só Graduate Assistantships (GTA, GRA, GAA), para full-time "
        "presencial: tuition reduzida a 1/3 da tarifa de residente + "
        "stipend mensal; número limitado; exige GRE. Não há bolsa "
        "institucional de mérito para o mestrado -- a página de Graduate "
        "Scholarships só lista fontes externas"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- internacionais F-1 são elegíveis a GA (tuition a 1/3 "
        "da tarifa de residente + stipend), mas as vagas são limitadas, "
        "exigem GRE (~320+ preferido) e a própria DSU avisa que há mais "
        "candidatos que vagas e que não se deve depender só da GA. Sem GA, "
        "o custo já é baixo (~US$ 21k de tuition com fees incluídas)"
    ),
    "Principais Empregadores": (
        "Amazon, eBay, Kroger, Disney, T-Mobile, Verizon, Bancorp, "
        "Citibank, Chase, Wells Fargo, Sanford Health, Mayo Clinic, Avera "
        "Health (parceiros citados na página do MSAA)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 76 (antes de 21 jan 2026) ou 4.0 na nova escala "
        "(sub-notas ≥4.0) / TOEFL PBT 550 / IELTS 6.0 (sub-notas ≥5.5) / "
        "PTE 60 / Duolingo 110 (sub-nota integrada ≥105) / ELS nível 112; "
        "válidos por 2 anos; TOEFL e PTE oficiais enviados pela agência "
        "(código TOEFL da DSU: 6247)"
    ),
    "Documentos Necessários": (
        "Aplicação online (taxa US$40); purpose statement (200-300 "
        "palavras); transcripts oficiais com grau conferido (internacionais: "
        "tradução para inglês; a DSU pode pedir avaliação por agência); "
        "comprovante de inglês; nenhuma carta de recomendação exigida; "
        "Certificate of Financial Responsibility (Masters CFR) e prova de "
        "seguro saúde para internacionais; para GA: Assistantship "
        "Application, resume e GRE"
    ),
    "Deadline Aplicação": (
        "15 abr (Fall, internacional; assistantship também 15 abr); 15 ago "
        "(Fall, doméstico); Spring: 1 nov (intl) / 30 nov (doméstico); "
        "Summer: 30 abr (só doméstico). Para ingresso (Fall 2028): ~15 "
        "abr 2028"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para Dakota State (Madison, SD). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Dakota State, Madison, SD)")

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
