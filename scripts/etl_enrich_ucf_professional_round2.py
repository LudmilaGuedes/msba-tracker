"""
Data enrichment (round 2): atualiza APENAS a linha 'UCF (Professional)'
(Professional MSM / Business Analytics track, Orlando, FL) no
msba_programs_database.xlsx.

ATENCAO (repetida do round 1): distinta da linha 'UCF' (MS in Business
Analytics tradicional, ja enriquecida em rodadas separadas). O matching
exige Nome da Universidade == 'UCF (Professional)' E Estado == 'FL' E
Cidade == 'Orlando'.

Foco exclusivo: trilha Business Analytics do Professional MSM, formato
cohort presencial noturno p/ profissionais que já trabalham. Deep research
em: custo de tuition p/ internacionais, bolsas/assistantships,
GPA/GMAT/GRE/inglês, custo de vida -- a pedido da usuaria, mesmo processo
das rodadas anteriores (FSU, USF, UCF tradicional).

VERIFICACAO CRITICA -- A TAXA FLAT DE $30.000 E TOTAL, NAO ANUAL:
o round 1 ja tinha essa cifra cadastrada como 'Tuition Total', mas o texto
da fonte ('$30.000/ano') deixava ambiguidade se seria cobrado por ano
(preocupante, dado que o programa dura ~20 meses, ou seja, mais de 1 ano
-- poderia significar um total real de ~$50k, nao $30k). Reconfirmado
nesta rodada, com prompt especifico, direto na fonte oficial
(studentaccounts.ucf.edu/tf-profmsmgmt-ba/): a nota oficial diz "The
Program Tuition includes all credit hours..." -- linguagem de valor
UNIFICADO/TOTAL do programa (todos os creditos), NAO uma cobranca anual
recorrente. Tuition Total permanece $30.000 (sem alteracao), agora com
mais confianca.

CUSTO/CREDITO (novo, antes "Verificar"): preenchido como valor DERIVADO
($30.000 / 30 créditos = $1.000/crédito) -- não é uma cobrança real por
crédito (é taxa flat), mas um número derivado para permitir comparação
com as outras linhas do dataset, mesmo criterio ja usado para outras
escolas com tarifa flat (ex.: UC Cincinnati).

ADMISSAO (novo/refinado):
- Taxa de Aplicação: $30 (confirmado, fonte de busca agregada a partir
  do texto oficial do programa -- mesma ordem de grandeza da trilha
  tradicional, que cobra $31).
- Documentos Necessários: resume, goal statement (500-1.000 palavras),
  3 cartas de recomendação (professores e/ou supervisores atuais/
  anteriores); candidatos com graduação fora dos EUA precisam de 1
  transcript oficial + certificação oficial de conclusão de curso com
  data.
- Testes Inglês Aceitos: confirmado que TOEFL/IELTS/Duolingo são aceitos
  e exigidos de todo candidato não-cidadão americano, MAS não foi
  encontrado o score mínimo específico desta trilha (distinto da trilha
  tradicional 'UCF', que exige TOEFL 91/IELTS 7.0/Duolingo 125 --
  não presumido para esta trilha sem confirmação).
- GMAT/GRE: reconfirmado "não obrigatório" (sem mudança).
- STEM Designated: PERMANECE "Verificar" -- nova rodada de busca (catálogo
  oficial UCF, listas de programas STEM do Miller College of Business)
  novamente NÃO conseguiu confirmar CIP code ou designação STEM
  especificamente para esta trilha executiva (só a trilha tradicional
  'MS in Business Analytics' foi confirmada como STEM em fontes
  encontradas). Mantido o mesmo criterio conservador do round 1.
- Deadline/Rounds Aplicação: NAO encontrado deadline especifico desta
  trilha (a pagina oficial de deadlines do College of Graduate Studies
  so lista a trilha tradicional; buscas confirmam apenas "varia por
  trilha, contate o college diretamente") -- mantido "Verificar" em vez
  de escrever uma nao-resposta na celula.

CUSTO DE VIDA -- SEGURO SAUDE (novo, antes "Verificar"): reaproveitado o
valor oficial já confirmado para a linha 'UCF' nesta mesma sessão de
trabalho (https://studenthealth.ucf.edu/payments/intl/, UCF SHIP
2026-27 = $2.517/ano) -- é um requisito universitário (FL BOG regulation
6.009), não específico de programa, então válido para qualquer aluno F/J
da UCF, incluindo esta trilha. Prorateado para a duração estimada de ~20
meses: 2.517 x (20/12) = $4.195.

NAO ENCONTRADO (mantido "Verificar"): Bolsas Disponíveis, Tipos de Bolsa,
TA/RA Disponível, Processo TA/RA, Graduate Assistantships, Quão Realista
(Intl). A política geral de GA da UCF (0,5/0,25 FTE) existe, mas nenhuma
fonte confirma se esta trilha específica -- desenhada para profissionais
que JÁ TRABALHAM em tempo integral e cursam aulas à noite (seg/qua) --
efetivamente oferece essas vagas (estruturalmente improvável dado o
público-alvo, mas isso é inferência, não fonte confirmada, então não foi
escrito na célula). Ranking US News/QS, Taxa Empregabilidade %, Salário
Médio Pós, Principais Empregadores, GPA Médio Admitidos, GMAT Médio, Taxa
Aceitação %, Tamanho Turma, Entrevista: sem fonte encontrada.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UCF (Professional)",
    "Estado": "FL",
    "Cidade": "Orlando",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 1000.0,  # derivado: 30000 / 30 créditos (tarifa flat, não cobrança real por crédito)
    "Custo/Crédito (Intl) USD": 1000.0,  # mesma tarifa flat, sem premium intl
    "Taxa Aplicação USD": 30,
    "Documentos Necessários": (
        "Resume, goal statement (500-1.000 palavras), 3 cartas de "
        "recomendação (professores e/ou supervisores atuais/anteriores); "
        "candidatos com graduação fora dos EUA: 1 transcript oficial + "
        "certificação oficial de conclusão de curso com data"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/Duolingo aceitos, exigidos de candidatos não-"
        "cidadãos americanos; score mínimo específico desta trilha não "
        "encontrado (distinto da trilha tradicional 'UCF', que exige "
        "TOEFL 91/IELTS 7.0/Duolingo 125 -- não presumido aqui sem "
        "confirmação)"
    ),
    "Seguro Saúde (programa) USD": 4195,  # UCF SHIP anual (2517) x 20/12 meses (duração estimada da trilha)
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
        raise ValueError("Nenhuma linha encontrada para UCF (Professional) (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF (Professional), Orlando, FL)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
