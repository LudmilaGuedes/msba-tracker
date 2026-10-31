"""
Data enrichment: atualiza APENAS a linha da 'Rutgers University–Camden'
(Rutgers School of Business–Camden, Camden, NJ) no
msba_programs_database.xlsx.

Foco: MS in Business Analytics, trilha ON-CAMPUS (distinta da trilha
100% online, que a própria Rutgers confirma NÃO oferecer visto de
estudante -- "US visas are not offered for this program" -- consistente
com a nota já presente na planilha 'STEM Designated' = 'Sim (apenas
versão on-campus)', validando que a linha já estava corretamente
identificando a trilha certa para o perfil de visto F-1/J-1), 30
créditos (já confirmado), STEM, 3 ciclos de admissão por ano (Fall/
Spring/Summer).

TUITION: PREENCHIDO (estava 'Verificar' em ambos os campos). A página
oficial de tuition 2026-27 confirma, especificamente para a trilha
On-Campus MSBA (distinta da Online, que tem tarifa única de $1.050/
crédito): In-State $1.149,00/crédito + fees; Out-of-State $1.944,00/
crédito + fees. PREENCHIDO 'Custo/Crédito (Residente)' = $1.149 e
'(Intl)' = $1.944 (usando a tarifa out-of-state). 'Tuition Total'
(30 créditos): Residente = $34.470; Intl = $58.320. 'Fees Adicionais
USD' PERMANECE 'Verificar' -- a página confirma que fees existem
("plus Fees") mas não divulga o valor exato, e o link para o
detalhamento completo (finance.rutgers.edu) não trouxe a quebra
específica de Camden nesta rodada.

'GPA Mínimo' preenchido: 3.0 geral é o mínimo declarado no overview do
programa; candidatos com GPA 3.0-3.49 precisam de 2 cartas de
recomendação (vs. o padrão, não especificado). Uma fonte separada (FAQ)
menciona 2.75 como limiar para RECOMENDAR envio de GMAT/GRE -- ambos os
números documentados para não perder nenhuma informação, já que vêm de
páginas oficiais distintas.

'GMAT/GRE Obrigatório' preenchido com achado IMPORTANTE de duas
camadas: (1) NÃO exigido para admissão geral; recomendado apenas se
GPA <2.75 ou menos de 3 anos de experiência profissional; MAS (2) um
GMAT/GRE oficial é OBRIGATÓRIO para qualquer candidato que queira ser
CONSIDERADO PARA BOLSA -- inclusive internacionais em F-1/J-1 (a
página confirma explicitamente que bolsas estão disponíveis para
titulares de F-1/J-1, mas condicionadas ao envio do teste). Essa
distinção é relevante para o perfil de busca, que provavelmente
buscará bolsas.

'Testes Inglês Aceitos' preenchido com os mínimos oficiais em dois
patamares de TOEFL: até 21/jan/26 (Writing 22/Speaking 23/Reading 21/
Listening 17); a partir de 22/jan/26 (Writing 4.5/Speaking 4.5/
Reading 4/Listening 4); IELTS mín. 6.5; Duolingo mín. 120 (efetivo a
partir do Fall 2026).

'Taxa Aplicação USD' preenchido: $70. 'Documentos Necessários'
preenchido: aplicação de admissão; transcripts oficiais (equivalentes
a 120 créditos de bacharelado de 4 anos nos EUA); avaliação de
credencial curso-a-curso de agência credenciada NACES, se necessário;
comprovante de proficiência em inglês; 2 cartas de recomendação;
personal statement; CV/resume.

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com os prazos
oficiais, diferenciados por status: internacionais -- Fall (início
setembro) 1 jun; Spring (início janeiro) 1 out; domésticos -- Fall 1
ago; Spring 1 dez; Summer 1 mai. 3 ciclos de admissão por ano
confirmados (Fall/Spring/Summer).

'Bolsas Disponíveis' e 'Tipos de Bolsa' preenchidos com a mesma
ressalva de GMAT/GRE: bolsas disponíveis para titulares de F-1/J-1,
mas exigem envio de score oficial de GMAT ou GRE para consideração
(diferente da admissão geral, que dispensa o teste).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Fees Adicionais USD' (ver nota acima), 'Seguro Saúde (programa) USD'
(nenhuma menção encontrada em nenhuma página consultada), 'Início'
(o programa aceita 3 ciclos, sem um único "início" fixo -- mantido
'Verificar' por não haver um valor único correto para preencher),
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação
%', 'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade %',
'Salário Médio Pós USD', 'Principais Empregadores' (a página de
career outcomes só lista cargos-alvo genéricos, sem estatísticas nem
empregadores nomeados), 'Processo TA/RA', 'TA/RA Disponível',
'Caminho Tuition Residente', 'Experiência Profissional' (mencionada
apenas como critério de dispensa de GMAT/GRE, não como requisito em
si), 'Entrevista', e todos os campos mensais de custo de vida em
Camden, NJ.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://business.camden.rutgers.edu/graduate/ms-in-business-analytics/
  (formato on-campus vs online, duração, créditos, currículo, GPA,
  GMAT/GRE não exigido, 3 ciclos de admissão)
- https://business.camden.rutgers.edu/graduate/international_msba_applicants/
  (confirmação de que apenas a trilha on-campus é elegível para visto;
  deadlines internacionais, taxa de aplicação $70, documentos, mínimos
  de inglês em 2 patamares de TOEFL, bolsas condicionadas a GMAT/GRE)
- .../master-of-science-in-business-analytics-msba-faqs/ (deadlines
  domésticos, GMAT/GRE com critério de dispensa por GPA/experiência,
  cargos-alvo de carreira sem estatísticas)
- https://business.camden.rutgers.edu/prospective-students/tuition_fees/
  (tuition oficial 2026-27, discriminada entre On-Campus e Online, por
  residência)
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
    "Nome da Universidade": "Rutgers University–Camden",
    "Estado": "NJ",
    "Cidade": "Camden",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 1149,  # oficial 2026-27, On-Campus MSBA
    "Custo/Crédito (Intl) USD": 1944,  # oficial 2026-27, On-Campus MSBA (out-of-state)
    "Tuition Total (Residente) USD": 34470,  # 30 x $1.149
    "Tuition Total (Intl) USD": 58320,  # 30 x $1.944
    "GPA Mínimo": (
        "3.0 geral (mínimo declarado no overview do programa); GPA "
        "3.0-3.49 exige 2 cartas de recomendação; uma fonte separada "
        "(FAQ) cita 2.75 como limiar alternativo para recomendar "
        "envio de GMAT/GRE"
    ),
    "GMAT/GRE Obrigatório": (
        "NÃO exigido para admissão geral; recomendado se GPA <2.75 "
        "ou <3 anos de experiência profissional; MAS OBRIGATÓRIO "
        "(mesmo para internacionais em F-1/J-1) para ser considerado "
        "para bolsa"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL até 21/jan/26 (Writing 22/Speaking 23/Reading 21/"
        "Listening 17) ou a partir de 22/jan/26 (Writing 4.5/"
        "Speaking 4.5/Reading 4/Listening 4); IELTS mín. 6.5; "
        "Duolingo mín. 120 (efetivo Fall 2026)"
    ),
    "Taxa Aplicação USD": 70,
    "Documentos Necessários": (
        "Aplicação de admissão; transcripts oficiais (equivalente a "
        "120 créditos de bacharelado de 4 anos dos EUA); avaliação "
        "curso-a-curso de agência NACES (se necessário); comprovante "
        "de proficiência em inglês; 2 cartas de recomendação; "
        "personal statement; CV/resume"
    ),
    "Deadline Aplicação": "Internacionais: Fall (set) 1 jun; Spring (jan) 1 out. Domésticos: Fall 1 ago; Spring 1 dez; Summer 1 mai",
    "Rounds Aplicação": "3 ciclos de admissão por ano: Fall, Spring, Summer (domésticos); Fall e Spring (internacionais)",
    "Bolsas Disponíveis": "Sim, disponíveis para titulares de F-1/J-1, mas exige envio de score oficial de GMAT ou GRE para consideração",
    "Tipos de Bolsa": "Bolsas para F-1/J-1 condicionadas a envio de GMAT/GRE oficial (diferente da admissão geral, que dispensa o teste)",
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
        raise ValueError("Nenhuma linha encontrada para Rutgers University–Camden (Camden, NJ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Rutgers University–Camden, Camden, NJ)")

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
