"""
Data enrichment (round 2): atualiza APENAS a linha 'VT Global BA' (MS in
Business Administration -- concentração em Global Business Analytics /
MSBA-GBA, Pamplin College of Business, Blacksburg, VA) no
msba_programs_database.xlsx.

Pedido da usuaria: confirmar tuition para o perfil dela (internacional),
bolsas/descontos e se sao realistas, e buscar rankings/empregabilidade/
salario medio inicial.

=========================================================================
1) TUITION -- JA CONFIRMADA EM AUDITORIA ANTERIOR DESTA SESSAO
=========================================================================
$1.000/crédito, mesma tarifa para VA e não-VA (oficial
masters.pamplin.vt.edu/globalbusinessanalytics), 30 créditos = $30.000
total -- ja reconfirmado no lote das "20 mais baratas" (sem alteracao
necessaria aqui).

=========================================================================
2) ACHADO CRITICO -- PROGRAMA FECHADO PARA CANDIDATURA GERAL
=========================================================================
A pagina oficial de candidatura (apply-now.html, indexada por busca --
fetch direto retornou 404 mas o snippet foi corroborado por DUAS buscas
independentes com o mesmo texto especifico) diz:
  "The MSBA-GBA program is not accepting applications for academic year
   2025-2026, unless you are a student in the 3+1+1 program between
   Virginia Tech and Narsee Monjee Institute of Management Studies."

Ou seja: para o ano letivo 2025-2026, este programa NAO esta aceitando
candidaturas gerais -- so aceita alunos do pathway 3+1+1 com o NMIMS
(Mumbai, India), que cursam 2 das 10 disciplinas ainda como
graduandos da VT e completam as 8 restantes no MSBA-GBA. Isso e
coerente com o historico conturbado do programa: lancado Fall 2024 no
campus de Northern Virginia, mudou para o Innovation Campus (Alexandria)
na Spring 2025, e migrou de novo para Blacksburg no Fall 2025 --
tres localizacoes em menos de 2 anos.
Status para 2026-2027 (quando a usuaria provavelmente se candidataria)
NAO FOI CONFIRMADO -- nenhuma fonte encontrada especifica esse ano. Isso
foi registrado em 'Deadline Aplicação' e 'Rounds Aplicação' como
ressalva central, ja que muda a pergunta de "e caro/tem bolsa" para
"da pra se candidatar, e quando".

=========================================================================
3) BOLSAS/DESCONTOS -- NAO ENCONTRADO PARA ESTE PROGRAMA ESPECIFICO
=========================================================================
Um processo de Graduate Assistantship (checkbox na aplicacao, sem
garantia, compensacao variavel aprovada pelo Board of Visitors) foi
encontrado, MAS pertence ao MSBA-BA -- um programa DIFERENTE e mais
antigo (fundado 2015, administrado pelo Center for Business Analytics/
CBA, msba.vt.edu/ba/), nao ao MSBA-GBA (Global). Confirmado por busca
dedicada: CBA "operates the Masters of Science in Business
Administration -- Business Analytics degree program" -- e a trilha
"Business Analytics" (sem "Global"), historicamente estabelecida, com
"Class of 2025" ja formada. NAO USADO nesta linha para evitar
contaminacao entre programas (mesmo criterio ja aplicado varias vezes
nesta sessao -- LSU, USC MBA vs MSBA, FSU predecessor, UGA "two-year").
Nenhuma pagina especifica do MSBA-GBA menciona bolsas/scholarships --
'Bolsas Disponíveis', 'Tipos de Bolsa' e 'Graduate Assistantships'
permanecem "Verificar".

=========================================================================
4) RANKING/EMPREGABILIDADE/SALARIO -- NAO ENCONTRADO PARA ESTE PROGRAMA
=========================================================================
Os numeros encontrados (97,5% de colocacao, salario mediano $82k, bonus
$5k, ranqueado #20 pela TechGuide, empregadores NASA/Leidos/Advance
Auto/Norfolk Southern/Lockheed Martin) TAMBEM pertencem ao MSBA-BA (o
programa irmao mais antigo, msba.vt.edu/ba/about.html), NAO ao MSBA-GBA.
NAO USADOS por risco de contaminacao -- alem disso, faz sentido
estrutural que o MSBA-GBA nao tenha esses dados ainda publicados: e um
programa lancado em 2024, que mudou de campus duas vezes e pausou
admissao geral para 2025-26 -- mesmo que ja tenha uma primeira turma
formanda, e improvavel que estatisticas de emprego consolidadas ja
existam ou tenham sido publicadas. 'Ranking US News', 'Ranking QS',
'Taxa Empregabilidade %', 'Salário Médio Pós USD' e 'Principais
Empregadores' permanecem "Verificar".

=========================================================================
5) STEM DESIGNATED -- NAO CONFIRMADO PARA ESTE PROGRAMA
=========================================================================
Um resultado de busca dizia "Business Information Technology is a STEM
major" -- mas BIT e um curso de GRADUACAO (bit.vt.edu), nao o MSBA-GBA.
Descartado. 'STEM Designated' permanece "Verificar" -- dado critico p/
elegibilidade de OPT que nao pode ser presumido sem fonte direta.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "VT Global BA",
    "Estado": "VA",
    "Cidade": "Blacksburg",
}

UPDATES = {
    "Deadline Aplicação": (
        "ATENÇÃO: para o ano letivo 2025-2026, o programa NÃO está "
        "aceitando candidaturas gerais -- restrito a alunos do pathway "
        "3+1+1 com o NMIMS (Mumbai, Índia). Status para 2026-2027 não "
        "confirmado; verificar diretamente com o programa antes de "
        "aplicar"
    ),
    "Rounds Aplicação": (
        "Admissão contínua (rolling) quando aberta -- mas ver ressalva "
        "crítica em 'Deadline Aplicação': candidatura geral pode estar "
        "fechada"
    ),
    "Quão Realista (Intl)": (
        "Incerto -- não há dado de bolsa/GA específico deste programa "
        "(o único achado pertence ao MSBA-BA, programa irmão diferente), "
        "e o acesso em si está em dúvida: sem candidatura geral aberta "
        "para 2025-26 (restrito ao pathway NMIMS), status 2026-27 não "
        "confirmado. Recomenda-se contato direto com o programa antes "
        "de considerar esta opção"
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
        raise ValueError("Nenhuma linha encontrada para VT Global BA (Blacksburg, VA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (VT Global BA, Blacksburg, VA)")

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
