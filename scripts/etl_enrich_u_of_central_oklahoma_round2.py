"""
Data enrichment (round 2): atualiza APENAS a linha 'U of Central
Oklahoma' (College of Business, Edmond, OK) no
msba_programs_database.xlsx.

Pedido da usuaria: confirmar tuition para o perfil dela (internacional),
bolsas/descontos (se existem de verdade e quão realistas são), e buscar
rankings/empregabilidade/salário médio inicial.

=========================================================================
1) TUITION -- JA CONFIRMADA EM AUDITORIA ANTERIOR DESTA SESSAO
=========================================================================
$417/crédito internacional (oficial), 32 créditos = $13.344 -- ja
reconfirmado no lote das "20 mais baratas". Sem alteracao.

=========================================================================
2) BOLSAS -- CONFIRMADAS, COM DETALHE NOVO IMPORTANTE
=========================================================================
As duas bolsas ja cadastradas foram CONFIRMADAS por fonte independente
(agregador especializado em bolsas, WeMakeScholars/scholarship-
positions, cruzado com fonte oficial UCO sobre o processo):
  - Distinguished Graduate Scholar: ~$3.000/ano -- confirmado.
  - Graduate Merit Global Scholar: ~$2.000/ano, até 2 anos -- confirmado.
    Elegibilidade: candidatos de QUALQUER país; GPA mínimo de graduação
    3.25 (bolsa inicial); matrícula em tempo integral; RENOVÁVEL com GPA
    ≥ 3.00 -- e o mais importante: "no application required" -- ou
    seja, é avaliação automática, não uma candidatura separada.
CONFIRMADO tambem (fonte: pagina oficial UCO sobre financas p/
internacionais, ja indexada): candidatos internacionais sao revisados
automaticamente pelo Office of Global Affairs para elegibilidade de
bolsa apos a admissao -- reforca que NAO exige candidatura extra, e que
os ~$700 mil/ano ja cadastrados (distribuidos a internacionais na UCO
como um todo) sao reais.
'International Achievement Award' (ja cadastrada, citada pelo nome em
multiplas fontes) segue SEM VALOR EXATO encontrado -- mantido no texto,
mas sem numero, como ja estava.

CALCULO DE REALISMO (novo, adicionado a 'Quão Realista'): as duas bolsas
com valor confirmado, se o perfil da candidata atender aos criterios
(GPA >= 3.25 na graduacao), podem somar ATE ~$5.000/ano -- ao longo dos
24 meses de duracao do programa, isso seria ate ~$10.000 de um total de
$13.344 (internacional) -- MAS NAO HA CONFIRMACAO EXPLICITA de que as
duas bolsas sejam cumulativas para o mesmo aluno (as fontes as listam
separadamente, nao como um pacote combinado). Registrado como
possibilidade, nao certeza.

=========================================================================
3) RANKING/EMPREGABILIDADE/SALARIO -- NAO ENCONTRADO PARA O MSBA
=========================================================================
Buscas extensas nao encontraram taxa de empregabilidade, salario medio
ou ranking ESPECIFICOS do MS in Business Analytics da UCO. Dois cuidados
de desambiguacao durante a pesquisa:
  - Uma busca retornou dados de RANKING da "UC" (University of
    Cincinnati) em vez de "UCO" -- descartado imediatamente, escola
    errada.
  - Outra busca retornou "100% job placement" associado a "national
    selling competitions" -- linguagem tipica de um programa de VENDAS
    (Sales), nao do MSBA. Descartado por risco de contaminacao entre
    programas (mesmo criterio ja aplicado varias vezes nesta sessao).
CONFIRMADO (e adicionado aos destaques, nao a um campo de ranking, ja
que nao e um ranking numerico): o MSBA e credenciado pela AACSB
International, selo mantido por "less than 5% of business schools
worldwide" -- e um sinal real de qualidade, mas nao substitui um
ranking US News/QS proprio do programa, que nao foi encontrado (UCO e
uma universidade regional de porte medio, sem foco em pesquisa -- e
plausivel que o MSBA especifico simplesmente nao seja ranqueado por
essas publicacoes).
'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade %', 'Salário
Médio Pós USD' e 'Principais Empregadores' permanecem "Verificar".
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Central Oklahoma",
    "Estado": "OK",
    "Cidade": "Edmond",
}

UPDATES = {
    "Tipos de Bolsa": (
        "UCO International Achievement Award (mérito, valor exato não "
        "divulgado); Distinguished Graduate Scholar (~$3.000/ano, "
        "CONFIRMADO); Graduate Merit Global Scholar (~$2.000/ano, até 2 "
        "anos, CONFIRMADO -- exige GPA de graduação ≥ 3,25, renovável "
        "com GPA ≥ 3,00, aberta a candidatos de qualquer país). Todas "
        "avaliadas automaticamente pelo Office of Global Affairs após a "
        "admissão -- SEM candidatura separada exigida"
    ),
    "Quão Realista (Intl)": (
        "Moderada a Alta -- bolsas confirmadas são automáticas (sem "
        "candidatura extra, revisão pós-admissão) e ~$700 mil/ano são "
        "de fato distribuídos a internacionais na UCO. Se qualificar "
        "para as duas bolsas com valor confirmado (GPA de graduação >= "
        "3,25), o potencial combinado é de até ~$5.000/ano -- mas as "
        "fontes não confirmam se as duas são cumulativas para o mesmo "
        "aluno, então trate como estimativa, não garantia"
    ),
    "Destaques Currículo": (
        "32 créditos; MSBA credenciado pela AACSB International, selo "
        "mantido por menos de 5% das escolas de negócios do mundo"
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
        raise ValueError("Nenhuma linha encontrada para U of Central Oklahoma (Edmond, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Central Oklahoma, Edmond, OK)")

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
