# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): terceira tentativa de Deep Research, por
angulos ainda nao tentados nas rodadas 1 e 2, focada em UMA unica linha
-- University of Alabama (Manderson/Culverhouse), Tuscaloosa, AL -- MS
in Business Analytics, presencial, STEM.

Pedido da usuaria: repetiu o mesmo pedido de Data Enrichment ja feito
antes (ver etl_enrich_alabama_round2.py) -- tratado como pedido para
tentar de novo os campos que continuavam "Verificar" apos a rodada 2.

Angulos tentados NESTA rodada, que ainda nao tinham sido tentados antes:
pagina de recrutamento corporativo do Manderson, brochures/fact sheets
de terceiros (Poets&Quants, mim-essay), FAQ do MBA re-verificada
especificamente por mencao a Specialized Masters/MSBA, e -- o achado que
deu resultado -- a pagina oficial de custo de vida para fins de visto
(I-20) de alunos internacionais.

ACHADOS:

1) SEGURO SAUDE PREENCHIDO (estava "Verificar"): USD 2.428/ano,
   confirmado na pagina oficial international.ua.edu/isss (orcamento de
   custo de vida exigido para emissao do I-20, ano letivo 2026-27) --
   USD 1.018 (Fall) + USD 1.410 (Spring), pago em parcelas mas cobrindo
   os 12 meses.

2) TRANSPORTE/MES PREENCHIDO (estava "Verificar"): USD 100/mes,
   confirmado explicitamente na mesma fonte oficial (a pagina nao lista
   transporte como linha separada, mas o texto explicativo confirma que
   esse valor ja esta embutido em "Personal Expenses" como "$100 per
   month").

NAO ENCONTRADO NESTA RODADA (permanece "Verificar", mesmo apos multiplos
angulos de busca diferentes das rodadas 1 e 2 -- pagina de recrutamento
corporativo, agregadores de ranking especializados (Poets&Quants,
mim-essay), FAQ re-verificada, brochures/fact sheets): Tamanho de Turma,
GPA Medio Admitidos, GMAT Medio, Taxa de Aceitacao %, Taxa de
Empregabilidade %, Entrevista, Utilities/mes, Custo Vida Total
(programa), Ranking US News, Ranking QS. Conclusao reforcada da rodada
anterior: o Manderson NAO publica esses numeros especificos do MSBA
(Specialized Masters) em nenhuma pagina oficial encontrada -- so
existem dados agregados de TODOS os "Specialized Masters" juntos (nao
so Business Analytics) ou dados exclusivos do MBA, nenhum dos dois
atribuivel com seguranca a linha do MSBA sem risco de contaminacao
entre programas (mesmo erro ja identificado e evitado nesta sessao para
outras faculdades, ex. University of Washington).

Nao foi possivel calcular "Custo Vida Total (programa)" com confianca
mesmo apos os 2 novos dados: falta ainda o valor de Utilities/mes, entao
qualquer soma feita agora ficaria silenciosamente incompleta (um total
que parece fechado mas na verdade omite uma categoria real de custo) --
preferiu-se manter "Verificar" a publicar um total sabidamente
subestimado.

Fonte adicional desta rodada:
- international.ua.edu/isss/future-new-students/expenses/ (orcamento
  oficial de custo de vida para I-20, ano letivo 2026-27, tabela
  "Undergraduate & Graduate")

Uso: python etl_enrich_alabama_round3.py (executar a partir da raiz do
repositorio).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Alabama",
    "Estado": "AL",
    "Cidade": "Tuscaloosa",
}

UPDATES = {
    "Seguro Saúde (programa) USD": 2428,
    "Transporte/mês USD": 100,
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
        raise ValueError("Nenhuma linha encontrada para U of Alabama (Tuscaloosa, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Alabama, Tuscaloosa, AL)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (U of Alabama):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
