"""
Data enrichment: atualiza APENAS a linha da 'University of Pittsburgh
(Katz)' (Pittsburgh, PA) no msba_programs_database.xlsx.

NOTA IMPORTANTE (orçamento de busca web da sessão esgotado: 200/200
WebSearch); toda a pesquisa deste enrichment foi feita via WebFetch,
navegando diretamente pelo site oficial da Katz Graduate School of
Business (business.pitt.edu) e do sistema de tuition da universidade
(tuition.pitt.edu). Isso limitou o alcance de alguns campos (custo de
vida detalhado, seguro saúde) que normalmente seriam cruzados com
buscas adicionais -- documentado campo a campo abaixo.

CORREÇÃO DE NOME DE PROGRAMA (achado importante): a linha trazia
'Nome Oficial do Programa': 'MS in Management and Business Analytics
(+ variantes: Finance & BA, Marketing Science & BA, Accounting & BA)'
-- esse nome está DESATUALIZADO/impreciso. A pesquisa oficial mostra
que o programa standalone mais próximo de um MSBA hoje é o "Master of
Science in Analytics and Applications of AI for Business" (MS
Analytics & AI for Business), full-time, presencial, 9 meses, 30
créditos, STEM-designated -- um programa AUTÔNOMO, não uma variante do
MS Management. O antigo nome parece ter misturado esse programa com o
"MS Management" (que tem uma trilha dupla-degree combinando com o MS
Analytics & AI em ~3 semestres, mas é um produto diferente).
CORRIGIDO para o nome correto e atual. Sem exclusão de F-1/
internacionais identificada -- prosseguido com enrichment normal, sem
necessidade de escalar ao usuário (mesmo critério usado no caso
Rutgers-Newark nesta mesma sessão).

'Duração (meses)' CORRIGIDO de 17 para 9 -- a duração de 17 meses não
corresponde a nenhuma variante oficial encontrada do programa
standalone full-time (9 meses) nem part-time (~18 meses); provavelmente
resquício da confusão de nome de programa acima. 9 meses é o valor
oficial confirmado para o formato full-time presencial, foco desta
pesquisa.

TUITION: CORRIGIDO. O valor anterior era uma string de faixa
("45.800–68.700 conforme variante (Verificar); especialidade
~US$51.204/ano (US News)") sem base clara em uma fonte única
confiável. A página oficial de tuition da universidade (tuition.pitt.edu)
confirma, para 2026-27, MS full-time: $17.550/termo (residente PA) /
$26.626/termo (não-residente/internacional). Assumindo 2 termos
full-time para completar os 30 créditos em 9 meses (Fall + Spring):
Tuition Total = 2 × $17.550 = $35.100 (residente) / 2 × $26.626 =
$53.252 (internacional). 'Custo/Crédito' derivado: $1.170,00
(residente) / $1.775,07 (internacional).

'Fees Adicionais USD' preenchido com $2.006 -- soma oficial das taxas
obrigatórias para todos os graduate students em regime full-time/2
termos (Student Activity $80 + Wellness & Recreation $1.060 +
Computing & Network $550 + Transportation $316).

'Transporte/mês USD' preenchido com $0 -- confirmado oficialmente pela
página de FAQ para candidatos internacionais: transporte público é
GRATUITO para alunos mediante apresentação do ID estudantil da Pitt.

'GMAT/GRE Obrigatório' CONFIRMADO/refinado: opcional para os programas
de MS (já estava correto na planilha), com nota adicional de que a
Katz não aceita mais o exame GRE At-Home desde 10/dez/2024.

'Testes Inglês Aceitos' preenchido com os mínimos oficiais (dois
regimes, pois o TOEFL mudou de formato em 21/jan/2026): TOEFL mín. 80
(exames antes de 21/jan/26) ou 4.0 com todos os sub-scores 4.0+ (exames
a partir de 21/jan/26); IELTS mín. 6.5 (recomendado 7.0); Duolingo mín.
125 (recomendado 130); validade de 2 anos.

'Entrevista' preenchido com 'Sim — obrigatória' (listada como
requisito no processo geral de admissão de MS da Katz). 'Documentos
Necessários' e 'Deadline Aplicação'/'Rounds Aplicação' preenchidos com
os 4 prazos oficiais (Early Advantage 15 out, decisão até 15 nov;
Round One 15 jan; International Final 15 abr; Domestic Final 15 jul) e
a lista de documentos (2 cartas de recomendação, 2 essays, resume,
transcript, entrevista; para internacionais: traduções certificadas se
aplicável, certificados de conclusão de curso).

'Taxa Aplicação USD' preenchido com $50 -- confirmado oficialmente na
página de FAQ para candidatos internacionais.

'Bolsas Disponíveis' e 'Tipos de Bolsa' preenchidos com CAUTELA: (1)
bolsa de mérito Katz com consideração automática para todos os
admitidos em programa full-time (valor variável, não divulgado
publicamente); (2) bolsa "Master Your Path" (~50% de desconto na
tuition) -- mas essa é EXCLUSIVA para ex-alunos de GRADUAÇÃO da
própria Pitt (graduating seniors e alumni), então NÃO se aplica
tipicamente a uma candidata internacional externa sem bacharelado na
Pitt; documentado como tal para não sugerir elegibilidade indevida.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Taxa Empregabilidade %', 'Salário Médio Pós USD', 'Principais
Empregadores', 'Ranking US News', 'Ranking QS', 'Tamanho Turma',
'Início', 'GPA Mínimo', 'GPA Médio Admitidos', 'Taxa Aceitação %',
'Matérias Pré-Requisito', 'TA/RA Disponível', 'Processo TA/RA',
'Graduate Assistantships' -- não encontrados nas páginas oficiais
consultadas via WebFetch (a página do programa lista apenas
possibilidades de carreira genéricas, sem estatísticas). 'Seguro Saúde
(programa) USD' e todos os campos mensais de custo de vida (aluguel,
mercado, utilities, celular, wifi, outros gastos, custo vida total)
também permanecem 'Verificar' -- a página oficial de seguro saúde da
Pitt retornou erro 403 ao ser acessada via WebFetch, e sem WebSearch
disponível (orçamento esgotado) não foi possível cruzar fontes
alternativas de custo de vida em Pittsburgh com confiança suficiente
para preencher esses campos nesta rodada.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.business.pitt.edu/ms-analytics-and-ai/ (nome oficial,
  formato, duração, créditos, currículo, GMAT/GRE opcional, deadlines,
  carreiras possíveis — sem estatísticas de outcomes)
- https://www.tuition.pitt.edu/graduate/pa-resident e
  https://www.tuition.pitt.edu/graduate/out-of-state (tuition oficial
  2026-27, residente e não-residente, taxas obrigatórias)
- https://business.pitt.edu/katz-admissions/get-admissions-information/
  (requisitos gerais de MS: teste opcional, mínimos de inglês,
  documentos, entrevista obrigatória, deadlines)
- https://business.pitt.edu/katz-admissions/katz-tuition-and-aid/
  (bolsas de mérito, Master Your Path, empréstimos)
- https://www.business.pitt.edu/ms-international-applicant-admission-and-faq/
  (taxa de aplicação $50, transporte público gratuito com ID
  estudantil, documentos para internacionais)
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
    "Nome da Universidade": "University of Pittsburgh (Katz)",
    "Estado": "PA",
    "Cidade": "Pittsburgh",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "Master of Science in Analytics and Applications of AI for Business (MS Analytics & AI for Business)",
    "Duração (meses)": 9,  # CORREÇÃO: oficial full-time, não 17 (resquício de nome de programa desatualizado)
    "Destaques Currículo": (
        "Core: análise estatística, programação de dados, decision "
        "technologies, data mining; Technical Foundation of "
        "Generative AI; Deploying AI for Business: Strategy, Design, "
        "and Execution; eletivas com foco em Responsible AI"
    ),
    "Custo/Crédito (Intl) USD": 1775.07,  # 53.252 / 30, oficial 2026-27
    "Custo/Crédito (Residente) USD": 1170.00,  # 35.100 / 30, oficial 2026-27
    "Tuition Total (Intl) USD": 53252,  # CORREÇÃO: 2 termos x $26.626 (não-residente, oficial 2026-27)
    "Tuition Total (Residente) USD": 35100,  # 2 termos x $17.550 (residente PA, oficial 2026-27)
    "Caminho Tuition Residente": "Não aplicável na prática para internacionais em programa curto (9 meses); tarifa não-residente é a relevante",
    "Fees Adicionais USD": 2006,  # taxas obrigatórias oficiais, full-time 2 termos
    "Transporte/mês USD": 0,  # transporte público gratuito com ID estudantil da Pitt, confirmado oficialmente
    "GMAT/GRE Obrigatório": "Opcional para programas de MS; Katz não aceita mais GRE At-Home desde 10/dez/2024",
    "Testes Inglês Aceitos": (
        "TOEFL mín. 80 (exames antes de 21/jan/26) ou 4.0 com todos "
        "os sub-scores 4.0+ (exames a partir de 21/jan/26); IELTS "
        "mín. 6.5 (recomendado 7.0); Duolingo mín. 125 (recomendado "
        "130); validade de 2 anos"
    ),
    "Entrevista": "Sim — obrigatória (listada como requisito no processo geral de admissão de MS da Katz)",
    "Deadline Aplicação": "Early Advantage: 15 out (decisão até 15 nov); Round One: 15 jan; International Final: 15 abr; Domestic Final: 15 jul",
    "Rounds Aplicação": "4 rounds: Early Advantage, Round One, International Final, Domestic Final",
    "Documentos Necessários": (
        "2 cartas de recomendação; 2 essays; resume; transcript; "
        "entrevista obrigatória; para internacionais: traduções "
        "certificadas de documentos em outro idioma, certificados/"
        "diplomas de conclusão de curso"
    ),
    "Taxa Aplicação USD": 50,  # confirmado oficialmente para candidatos internacionais
    "Bolsas Disponíveis": "Sim — bolsa de mérito Katz com consideração automática para todos admitidos em programa full-time (valor variável, não divulgado)",
    "Tipos de Bolsa": (
        "Bolsa de mérito Katz (automática, valor variável); "
        "\"Master Your Path\" (~50% de desconto) — EXCLUSIVA para "
        "ex-alunos de graduação da própria Pitt, não aplicável a "
        "candidatos externos/internacionais sem bacharelado na Pitt"
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
        raise ValueError("Nenhuma linha encontrada para University of Pittsburgh (Katz) (Pittsburgh, PA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Pittsburgh (Katz), Pittsburgh, PA)")

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
