"""
Data enrichment: atualiza APENAS a linha da 'Rutgers Business School
(Newark)' (Newark, NJ) no msba_programs_database.xlsx.

RESSALVA IMPORTANTE SOBRE O PROGRAMA (decisão documentada, sem
escalar ao usuário): o campus de Newark da Rutgers Business School NÃO
possui um programa chamado literalmente "Master of Science in Business
Analytics" -- a linha já trazia (de pesquisa anterior) o nome oficial
"Master of Information Technology and Analytics (MITA)", um programa
mais amplo de 30 créditos com 3 áreas de concentração (Cyber Security;
AI/ML/Data Science; e Operations Research and Business Analytics).
Diferente do caso Vanderbilt (removido nesta mesma planilha), aqui NÃO
há exclusão estrutural de alunos F-1/internacionais -- o programa é
STEM-designated, aceita internacionais normalmente (apenas exige
regime full-time), e a concentração "Operations Research and Business
Analytics" é diretamente relevante à busca. Por isso, a linha foi
mantida e enriquecida como "equivalente mais próximo" em Newark, com
esta ressalva registrada aqui em vez de removida ou sinalizada ao
usuário. (Nota: já existe uma linha separada e correta para "Rutgers
University–Camden", com um programa chamado oficialmente "MS in
Business Analytics (On-Campus)" -- campus e programa diferentes, não
tocados por este script.)

Foco: MITA, formato presencial (majoritariamente Newark), STEM, 30
créditos, 1 a 1,5 anos (full-time). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de
GPA/GMAT/GRE/inglês, e custo de vida em Newark, NJ.

TUITION: PREENCHIDO (estava 'Verificar'). A página oficial de tuition
do MITA (2026-27) lista tarifas por SEMESTRE para regime full-time
(12+ créditos, tarifa fixa independente do número exato de créditos
acima de 12): $15.444,00/semestre (residente NJ) / $22.020,00/semestre
(fora do estado/internacional). Assumindo 2 semestres full-time para
completar os 30 créditos (padrão do programa, 1 ano): Tuition Total =
2 × $15.444 = $30.888 (residente) / 2 × $22.020 = $44.040
(internacional). 'Custo/Crédito' derivado dividindo por 30: $1.029,60
(residente) / $1.468,00 (internacional). 'Total Créditos' preenchido
com 30 (confirmado oficialmente).

'Fees Adicionais USD' preenchido com $5.829 -- soma das taxas
obrigatórias oficiais por semestre em regime full-time (Student Fee
$979,50 + Technology Fee $221 + Off-Site Fee $1.119,50 + School Fee
$594,50 = $2.914,50/semestre) × 2 semestres.

'GMAT/GRE Obrigatório' PREENCHIDO: teste é OPCIONAL (test-optional),
mas aceito se enviado -- médias de admitidos recentes: GMAT ~618, GRE
~307 (148 verbal / 159 quant). 'Testes Inglês Aceitos' preenchido:
TOEFL (iBT, média de admitidos ~95), IELTS mín. 6.5, PTE ou Duolingo
também aceitos; dispensado para diploma de instituição em país de
língua inglesa. 'Matérias Pré-Requisito' preenchido com o requisito
oficial de conhecimento básico de cálculo e estatística.

'Taxa Aplicação USD' preenchido com $0 -- confirmado oficialmente que
a taxa de aplicação é DISPENSADA para todos os candidatos deste
programa. 'Deadline Aplicação' preenchido com prazos DIFERENTES para
domésticos (15 jul Fall / 15 dez Spring) e internacionais (15 mai Fall
/ 15 out Spring) -- prazos internacionais são mais cedo.

'Bolsas Disponíveis' preenchido: bolsa por mérito de $5.000 para 4
alunos da turma de outono e 2 alunos da turma de primavera (número
limitado de vagas). 'TA/RA Disponível' e 'Graduate Assistantships'
preenchidos com 'Não disponível atualmente' -- confirmado
explicitamente pela FAQ oficial do programa (não é uma lacuna de
pesquisa, é uma confirmação negativa).

'Seguro Saúde (programa) USD' preenchido com $3.286/ano -- prêmio
oficial do Rutgers SHIP (Student Health Insurance Plan) 2026-27,
obrigatório para a maioria dos alunos (incluindo internacionais F-1/
J-1), com opção de waiver mediante seguro comparável.

CUSTO DE VIDA (Newark, NJ): preenchido com ESTIMATIVAS GERAIS de
confiança moderada (não são valores oficiais específicos do programa
-- mesmo padrão de cautela usado em preenchimentos anteriores de custo
de vida "geral da cidade", ex.: JHU/Georgetown em DC): 'Aluguel
Privativo/mês' $1.000 (quarto privativo em moradia compartilhada perto
do campus, faixa $950-1.030 encontrada); 'Mercado/mês' $350 (Newark
~3% acima da média nacional de alimentação); 'Utilities/mês' $175
(faixa $140-210 encontrada); 'Celular/mês' $50; 'WiFi/mês' $65 (média
nacional ~$66); 'Transporte/mês' $100 (NJ Transit/PATH -- sem
confirmação de passe gratuito para alunos deste programa, diferente
de casos como CMU/Oakland onde isso foi confirmado). 'Custo Vida Total
(programa) USD' permanece 'Verificar' -- não há um agregado oficial, e
somar estimativas próprias empilharia incerteza sobre incerteza (mesmo
critério de cautela usado em outras linhas desta planilha).

CUIDADO APLICADO (dados NÃO encontrados/usados): 'Taxa Empregabilidade
%', 'Salário Médio Pós USD', 'Ranking US News', 'Ranking QS', 'GPA
Mínimo', 'GPA Médio Admitidos', 'Taxa Aceitação %', 'Tamanho Turma',
'Início' e 'Entrevista' não foram encontrados em nenhuma fonte oficial
do MITA -- permanecem 'Verificar'. O orçamento de busca web desta
sessão se esgotou antes de uma última tentativa de encontrar dados de
emprego/salário; esses campos ficam como estavam.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.business.rutgers.edu/masters-information-technology-analytics/tuition
  (tuition e fees oficiais 2026-27, por semestre, full-time e part-time)
- https://www.business.rutgers.edu/masters-information-technology-analytics/frequently-asked-questions
  (médias de GMAT/GRE/TOEFL de admitidos, bolsas de $5.000, ausência de
  TA/RA, taxa de aplicação dispensada)
- https://www.business.rutgers.edu/masters-information-technology-analytics/admissions
  (deadlines domésticos vs. internacionais, pré-requisito de cálculo/
  estatística)
- https://www.business.rutgers.edu/masters-information-technology-analytics/application-checklist
  (lista de documentos exigidos, sem menção a entrevista)
- https://www.business.rutgers.edu/masters-information-technology-analytics
  (formato, concentrações, duração, STEM designation)
- Busca web: seguro saúde SHIP Rutgers $3.286/ano (oficial 2026-27,
  waiver disponível); estimativas gerais de custo de vida em Newark, NJ
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
    "Nome da Universidade": "Rutgers Business School (Newark)",
    "Estado": "NJ",
    "Cidade": "Newark",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 30,  # oficial
    "Custo/Crédito (Intl) USD": 1468.00,  # 44.040 / 30
    "Custo/Crédito (Residente) USD": 1029.60,  # 30.888 / 30
    "Tuition Total (Intl) USD": 44040,  # 2 semestres x $22.020 (out-of-state/intl, full-time, oficial 2026-27)
    "Tuition Total (Residente) USD": 30888,  # 2 semestres x $15.444 (NJ resident, full-time, oficial 2026-27)
    "Caminho Tuition Residente": "Não aplicável na prática para internacionais em programa de 1 ano; tarifa não-residente/internacional é a relevante",
    "Fees Adicionais USD": 5829,  # 2 semestres x $2.914,50 (Student+Technology+Off-Site+School Fee, full-time, oficial)
    "GMAT/GRE Obrigatório": (
        "Opcional (test-optional), mas aceito se enviado — médias de "
        "admitidos recentes: GMAT ~618, GRE ~307 (148 verbal / 159 "
        "quant)"
    ),
    "GMAT Médio": 618,
    "Testes Inglês Aceitos": (
        "TOEFL iBT (média de admitidos ~95), IELTS mín. 6.5, PTE ou "
        "Duolingo também aceitos; dispensado para diploma de "
        "instituição em país de língua inglesa"
    ),
    "Matérias Pré-Requisito": "Bacharelado + conhecimento básico de cálculo e estatística",
    "Taxa Aplicação USD": 0,  # confirmado oficialmente: taxa dispensada para todos os candidatos
    "Deadline Aplicação": "Internacionais: 15 mai (Fall) / 15 out (Spring); domésticos: 15 jul (Fall) / 15 dez (Spring)",
    "Rounds Aplicação": "Sem rounds numerados; dois ciclos por ano (Fall/Spring), prazos diferentes para domésticos e internacionais",
    "Documentos Necessários": (
        "Online application (taxa dispensada); históricos acadêmicos "
        "com conferral de grau; GMAT/GRE opcional; teste de inglês "
        "(não-nativos); 1 carta de recomendação; resume; personal "
        "statement (300-500 palavras); requisitos adicionais para "
        "candidatos internacionais"
    ),
    "Bolsas Disponíveis": "Sim — bolsa por mérito de $5.000 para 4 alunos (turma de outono) e 2 alunos (turma de primavera)",
    "Tipos de Bolsa": "Bolsa por mérito de $5.000 (número limitado de vagas por turma)",
    "TA/RA Disponível": "Não disponível atualmente (confirmado oficialmente pela FAQ do programa)",
    "Graduate Assistantships": "Não disponível atualmente (mesma confirmação oficial de TA/RA)",
    "Seguro Saúde (programa) USD": 3286,  # SHIP oficial Rutgers 2026-27, waiver disponível
    "Aluguel Privativo/mês USD": 1000,  # estimativa, quarto privativo perto do campus Newark (faixa $950-1.030)
    "Mercado/mês USD": 350,  # estimativa geral, Newark ~3% acima da média nacional de alimentação
    "Utilities/mês USD": 175,  # estimativa geral (faixa $140-210)
    "Celular/mês USD": 50,  # estimativa geral, plano individual
    "WiFi/mês USD": 65,  # estimativa geral (média nacional ~$66)
    "Transporte/mês USD": 100,  # estimativa geral, NJ Transit/PATH (sem passe gratuito confirmado)
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
        raise ValueError("Nenhuma linha encontrada para Rutgers Business School (Newark) (Newark, NJ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Rutgers Business School (Newark), Newark, NJ)")

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
