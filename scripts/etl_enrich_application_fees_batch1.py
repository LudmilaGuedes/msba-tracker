"""
Data enrichment (lote único): atualiza a 'Taxa Aplicação USD' (e, em
alguns casos, uma nota em 'Rounds Aplicação') de MÚLTIPLAS linhas do
msba_programs_database.xlsx, a partir de uma varredura ampla feita a
pedido do usuário: confirmar a taxa de aplicação de todas as ~109
faculdades do banco (exceto Baruch CUNY e Fordham, já tratadas em
rounds anteriores) e identificar se existe algum mecanismo de
isenção/desconto vinculado a round ou período de aplicação (pergunta
original: "existe outro programa em que se candidatar no primeiro
round é de graça, como a Fordham?").

METODOLOGIA: pesquisa via WebSearch/WebFetch em 9 lotes paralelos
(2026-09-08), cobrindo as 109 faculdades restantes. Cada escola foi
pesquisada priorizando a página oficial de admissões do programa
específico de MS Business Analytics (evitando confundir com taxa de
MBA, erro já cometido antes nesta base).

CRITÉRIO DE ATUALIZAÇÃO (conservador, mesmo princípio já usado nos
rounds anteriores): só escrevo no banco os valores que a pesquisa
CONFIRMOU (fonte oficial ou, na falta dela, fonte agregada consistente
e sem contradição forte). Onde a pesquisa encontrou informação
CONFLITANTE ou de baixa confiança, mantive o valor já existente e
documento a divergência aqui (e reporto ao usuário) em vez de
sobrescrever às cegas.

RESULTADO GERAL SOBRE A PERGUNTA ORIGINAL: nenhuma das 109 faculdades
pesquisadas tem uma estrutura de "Round 1 = taxa 100% isenta por
padrão para todo mundo" idêntica à da Fordham. O que existe, em vários
casos, são mecanismos DIFERENTES:
- Notre Dame: Early Action tem taxa isenta + bolsa de mérito garantida
  automaticamente (mecanismo mais parecido com o da Fordham
  encontrado nesta varredura) -- ADICIONADO abaixo.
- Wayne State: taxa está ATUALMENTE (checado em set/2026) totalmente
  isenta para os ciclos Spring/Summer/Fall 2026 e Winter 2027 -- mas é
  uma promoção institucional temporária, não uma regra permanente por
  round -- ADICIONADO abaixo com essa ressalva.
- Illinois Tech: NUNCA cobra taxa de aplicação (política permanente,
  não é "grátis só no primeiro round") -- CORRIGIDO abaixo (era
  'Verificar').
- Seattle U, Baylor, SMU, CU Boulder (banco já tinha $0): mesma
  lógica -- gratuito sempre, não por round. Não alterados (já
  corretos, exceto CU Boulder -- ver nota de divergência abaixo).
- U of Alabama: "Free App Week" -- semana promocional anual com taxa
  zerada (datas variam a cada ano) + código de isenção específico para
  internacionais com GPA >= 3.5, solicitado até um deadline fixo (10
  out). NÃO adicionado ao banco por ser uma promoção pontual/anual sem
  data fixa confirmável agora -- reportado ao usuário para
  acompanhamento manual se for do interesse dela.
- UMass Amherst: taxa não é isenta, mas é REEMBOLSADA na primeira
  fatura de tuition se aplicar até o priority deadline -- mecanismo
  parecido mas não idêntico (é reembolso condicional, não isenção
  antecipada). Não adicionado ao banco (dado já correto, mecanismo é
  mais uma nota de cor do que uma taxa a corrigir).

DIVERGÊNCIAS ENCONTRADAS E NÃO CORRIGIDAS (confiança insuficiente --
reportadas ao usuário, não escritas no banco):
- U of South Carolina: pesquisa sugere taxa normal de $50 (banco tem
  $108) e alega isenção total vigente "até novo aviso" para todos os
  candidatos -- fonte de confiança moderada, não é official .edu
  totalmente verificado nesta rodada. NÃO alterado.
- CU Boulder: banco tem $0 (gratuito), mas a pesquisa encontrou fontes
  conflitantes sugerindo que pode NÃO ser gratuito ($50-75) -- fetch
  oficial não confirmou nenhum valor. NÃO alterado, mas sinalizado
  como precisando de reverificação -- se o $0 atual estiver errado,
  isso é uma correção pendente importante.
- Johns Hopkins (Carey MSBA/AI): possível isenção TOTAL da taxa
  especificamente para esse programa, mas fonte primária bloqueou o
  fetch (HTTP 403) -- confiança média-baixa. NÃO alterado.
- Florida Tech: pesquisa encontrou $50 (presencial) vs. $60 já no
  banco, e menção de isenção total para o formato online (não
  confirmado se aplicável ao MSBA) -- conflito não resolvido. NÃO
  alterado.
- Georgia Tech, U of Arkansas, U of Cincinnati, UIC, U of Iowa (só
  para caso doméstico): pequenas divergências de fonte secundária vs.
  valor já no banco, mas sem confirmação oficial forte o bastante para
  substituir o dado existente. NÃO alterados.

CAMPOS ATUALIZADOS NESTA RODADA (confirmados/corrigidos com confiança
razoável):
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

# Cada item: (Nome da Universidade, Estado, Cidade) -> dict de updates
SCHOOL_UPDATES = [
    (
        {"Nome da Universidade": "Notre Dame", "Estado": "IN", "Cidade": "Notre Dame"},
        {
            "Taxa Aplicação USD": 100,  # oficial (era 50, desatualizado)
            "Rounds Aplicação": (
                "4 rounds: Early Action + Rounds 1-3 (ver Deadline de "
                "Aplicação). Early Action: taxa de aplicação isenta + "
                "bolsa de mérito garantida automaticamente."
            ),
        },
    ),
    (
        {"Nome da Universidade": "USF", "Estado": "FL", "Cidade": "Tampa"},
        {"Taxa Aplicação USD": 30},
    ),
    (
        {"Nome da Universidade": "VT Global BA", "Estado": "VA", "Cidade": "Blacksburg"},
        {"Taxa Aplicação USD": 75},
    ),
    (
        {"Nome da Universidade": "Binghamton SUNY", "Estado": "NY", "Cidade": "Binghamton"},
        {"Taxa Aplicação USD": 75},
    ),
    (
        {"Nome da Universidade": "Syracuse", "Estado": "NY", "Cidade": "Syracuse"},
        {"Taxa Aplicação USD": 75},
    ),
    (
        {"Nome da Universidade": "GWU", "Estado": "DC", "Cidade": "Washington"},
        {"Taxa Aplicação USD": 75},
    ),
    (
        {"Nome da Universidade": "Miami U Ohio", "Estado": "OH", "Cidade": "Oxford"},
        {"Taxa Aplicação USD": 50},
    ),
    (
        {"Nome da Universidade": "WashU", "Estado": "MO", "Cidade": "St. Louis"},
        {"Taxa Aplicação USD": 100},
    ),
    (
        {"Nome da Universidade": "SF State", "Estado": "CA", "Cidade": "San Francisco"},
        {"Taxa Aplicação USD": 70},
    ),
    (
        {"Nome da Universidade": "Oklahoma State", "Estado": "OK", "Cidade": "Stillwater"},
        {"Taxa Aplicação USD": 75},  # tarifa internacional confirmada (doméstico é 50)
    ),
    (
        {"Nome da Universidade": "Kent State", "Estado": "OH", "Cidade": "Kent"},
        {"Taxa Aplicação USD": 70},
    ),
    (
        {"Nome da Universidade": "Illinois Tech", "Estado": "IL", "Cidade": "Chicago"},
        {
            "Taxa Aplicação USD": 0,  # NUNCA cobra taxa (política permanente, não é "por round")
            "Rounds Aplicação": "Programa não cobra taxa de aplicação (política permanente, não vinculada a round).",
        },
    ),
    (
        {"Nome da Universidade": "Santa Clara University", "Estado": "CA", "Cidade": "Santa Clara"},
        {"Taxa Aplicação USD": 70},
    ),
    (
        {"Nome da Universidade": "Seton Hall University", "Estado": "NJ", "Cidade": "South Orange"},
        {"Taxa Aplicação USD": 75},  # taxa geral de pós-graduação (não achada versão específica do MSBA)
    ),
    (
        {"Nome da Universidade": "George Mason University", "Estado": "VA", "Cidade": "Fairfax"},
        {"Taxa Aplicação USD": 75},
    ),
    (
        {"Nome da Universidade": "Wayne State University", "Estado": "MI", "Cidade": "Detroit"},
        {
            "Taxa Aplicação USD": 50,  # taxa normal
            "Rounds Aplicação": (
                "ATUALMENTE (checado set/2026): taxa de aplicação "
                "totalmente isenta para os ciclos Spring/Summer/Fall "
                "2026 e Winter 2027 -- promoção institucional "
                "temporária, não permanente; confirmar vigência antes "
                "de aplicar."
            ),
        },
    ),
    (
        {
            "Nome da Universidade": "East Texas A&M University (ETAMU)",
            "Estado": "TX",
            "Cidade": "Commerce",
        },
        {"Taxa Aplicação USD": 50},  # oficial (era 75, desatualizado)
    ),
    (
        {
            "Nome da Universidade": "Texas A&M University-Texarkana (A&M-Texarkana)",
            "Estado": "TX",
            "Cidade": "Texarkana",
        },
        {"Taxa Aplicação USD": 50},
    ),
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for match, updates in SCHOOL_UPDATES:
        unknown_cols = [h for h in updates if h not in header_to_col]
        if unknown_cols:
            raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

        col_uni = header_to_col["Nome da Universidade"]
        col_estado = header_to_col["Estado"]
        col_cidade = header_to_col["Cidade"]

        target_rows = []
        for row in ws.iter_rows(min_row=2):
            if (
                row[col_uni - 1].value == match["Nome da Universidade"]
                and row[col_estado - 1].value == match["Estado"]
                and row[col_cidade - 1].value == match["Cidade"]
            ):
                target_rows.append(row[0].row)

        if len(target_rows) == 0:
            raise ValueError(f"Nenhuma linha encontrada para {match}. Abortando.")
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha corresponde a {match} (linhas {target_rows}). Abortando por seguranca.")

        row_idx = target_rows[0]
        print(f"\n{match['Nome da Universidade']} (linha {row_idx}):")

        for header, new_value in updates.items():
            col_idx = header_to_col[header]
            cell = ws.cell(row=row_idx, column=col_idx)
            before = cell.value
            cell.value = new_value
            cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
            print(f"  - {header}: {before!r} -> {new_value!r}")

    wb.save(DB_PATH)
    print("\nSalvo com sucesso.")


if __name__ == "__main__":
    main()
