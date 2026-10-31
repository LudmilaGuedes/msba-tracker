"""
Data enrichment: atualiza APENAS a linha da 'Georgia Tech' (Scheller
College of Business / interdisciplinar com Computing e Engineering,
Atlanta, GA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Analytics (MSA), campus Atlanta, presencial, STEM.
Deep research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida
local. Esta linha tinha o bloco inteiro de bolsas/TA-RA e a tuition
inteiramente como "Verificar".

ACHADO DA AUDITORIA: 'GMAT/GRE Obrigatório' estava como "Aceito" (sugere
opcional), mas para o campus presencial de Atlanta o GRE/GMAT é OBRIGATÓRIO,
sem dispensa (a dispensa só existe para o programa online/OMSA) --
CORRIGIDO.

Tuition: a página oficial de tuition-and-funding (2026-27) dá 3 tarifas
(residente GA, não-residente doméstico, internacional) para 2 durações
(1 ano/3 semestres ou 1.5 ano/4 semestres). Usada a opção de 1 ano/3
semestres, que bate com 'Duração (meses)' já cadastrada (12). O schema não
tem coluna para "não-residente doméstico", então essa tarifa intermediária
foi descartada (mesmo critério do Purdue). Os valores oficiais já incluem
uma taxa única de programa de $2.000 -- por isso 'Fees Adicionais USD' foi
MANTIDO "Verificar" em vez de preenchido com esse valor: somá-lo de novo
separadamente contaria a mesma taxa duas vezes no custo total.

Custo de vida: 'Seguro Saúde (programa)', 'Utilities/mês' e 'Transporte/mês'
continuam "Verificar" (não encontrados com confiança) -- por isso 'Custo
Vida Total (programa) USD' também permanece "Verificar": com 3 dos 9
componentes ainda faltando, somar os demais daria um total incompleto.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://www.analytics.gatech.edu/admissions/tuition-and-funding
  (tuition 2026-27 por residência/duração; GTA/GRA; fellowships; taxa de
  $2.000 já embutida nos totais)
- https://www.analytics.gatech.edu/admissions
  (deadlines: prioritário 1/jan, final 1/mar; documentos exigidos: GRE/GMAT,
  transcripts, resume, Statement of Purpose até 500 palavras, 3 cartas de
  recomendação; taxa de inscrição $95 doméstico / $105 internacional --
  reconciliado com um valor mais antigo de $75/$85 encontrado em outra
  página do mesmo domínio, claramente desatualizada)
- https://www.analytics.gatech.edu/prospective-students/common-questions
  (GRE/GMAT obrigatório p/ Atlanta sem dispensa; medianas 92-94 percentil
  quant / 85-90 percentil verbal; TOEFL 100 mín./seção 19; IELTS 7.5 geral
  com mínimos por seção; pré-requisitos de matemática e programação;
  placement rate histórico 95-100%)
- Salário médio ($109.000): fonte terceira (mim-essay.com), NÃO confirmado
  diretamente no relatório oficial "Georgia Tech Career Survey Report AY
  2024-2025", cujo detalhamento por programa não pôde ser acessado --
  registrado com essa ressalva
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
    "Nome da Universidade": "Georgia Tech",
    "Estado": "GA",
    "Cidade": "Atlanta",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Intl) USD": 61774,  # opção 1 ano / 3 semestres, 36 créditos, já inclui taxa única de programa de $2.000
    "Tuition Total (Residente) USD": 42064,  # residente da Geórgia, mesma opção de 1 ano / 3 semestres
    "GMAT/GRE Obrigatório": (
        "Obrigatório (sem dispensa) para o campus presencial de Atlanta "
        "-- GRE ou GMAT (escores antigos/vencidos aceitos para quem já "
        "tem pós-graduação); medianas dos admitidos: percentil 92-94 em "
        "quantitativo, 85-90 em verbal"
    ),  # CORRECAO: planilha tinha "Aceito" (sugeria opcional)
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 (mín. 19 por seção) / IELTS Academic mín. "
        "7.5 (mín. 6.5 em Reading/Listening/Speaking, mín. 5.5 em "
        "Writing) / GT Language Institute Test; dispensado para "
        "graduação cursada integralmente em instituição de ensino em "
        "inglês"
    ),  # CORRECAO/detalhamento: planilha tinha só "TOEFL/IELTS" vago
    "Matérias Pré-Requisito": (
        "Background em matemática (cálculo, probabilidade/estatística, "
        "alguma álgebra linear) e computação (pelo menos 1 curso de "
        "programação, preferencialmente Python)"
    ),
    "Taxa Aplicação USD": 105,  # internacional; doméstico é $95
    "Deadline Aplicação": "Prioritário: 1/jan; Final (late): 1/mar -- mesmas datas para candidatos domésticos e internacionais",
    "Rounds Aplicação": "2: Prioritário (1/jan) e Final/Late (1/mar)",
    "Documentos Necessários": (
        "GRE ou GMAT (escores vencidos aceitos), transcripts oficiais de "
        "todas as instituições, resume, Statement of Purpose (até 500 "
        "palavras), 3 cartas de recomendação, TOEFL/IELTS/GT Language "
        "Institute Test (internacionais), taxa de inscrição"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Fellowships (número limitado, para os melhores candidatos; "
        "consideração automática, sem inscrição separada)"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- dezenas de vagas de GTA/ano abertas a internacionais "
        "em tempo integral, mas fellowships são limitadas"
    ),
    "TA/RA Disponível": (
        "Sim (GTA/GRA, dezenas de vagas por ano; consideração automática, "
        "sem inscrição separada; disponível para internacionais em tempo "
        "integral)"
    ),
    "Processo TA/RA": (
        "Automático -- todos os candidatos são considerados, sem "
        "inscrição separada; não garantido além do termo inicial, "
        "contingente a desempenho e verba disponível"
    ),
    "Graduate Assistantships": (
        "Sim -- GTA (~33% FTE, ~US$1.184/mês + isenção de tuition) e GRA "
        "(~33% FTE, ~US$1.895/mês + isenção de tuition), valores "
        "aproximados por semestre"
    ),
    "Taxa Empregabilidade %": 97,  # aprox., ponto médio da faixa oficial divulgada (95-100% recebem ao menos 1 oferta)
    "Salário Médio Pós USD": 109000,  # fonte terceira (mim-essay.com), não confirmado no relatório oficial de carreira da GT
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
        raise ValueError("Nenhuma linha encontrada para Georgia Tech (Atlanta, GA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Georgia Tech, Atlanta, GA)")

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
