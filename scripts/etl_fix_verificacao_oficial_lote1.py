"""
Verificacao contra PAGINAS OFICIAIS -- LOTE 1: as 6 linhas que nunca
haviam passado por nenhum script de enriquecimento (portanto os valores
vinham direto da planilha de pesquisa manual original, sem fonte
rastreavel). Priorizadas por risco.

RESULTADO DO LOTE: 3 correcoes confirmadas em fonte oficial, 1 achado
critico de status de programa, 2 nao resolvidas (fonte bloqueada).

============================================================
1) GEORGE MASON UNIVERSITY (Fairfax, VA) -- ERRO GRAVE, CORRIGIDO
============================================================
Fonte oficial: https://business.gmu.edu/programs/graduate-degree/ms-business-analytics
  "$1,025 per credit" (Virginia/DC/Maryland residents)
  "$1,400 per credit" (out-of-state/international)
  30 creditos -> $30.750 (residente) / $42.000 (internacional)
  Ano letivo 2026-2027.

A planilha tinha $1.835,50/credito para AMBOS (residente e intl) e total
de $55.065 para ambos -- valor sem fonte rastreavel, que alem de estar
errado no numero, apagava a distincao residente/nao-residente que a GMU
de fato pratica (a GMU e PUBLICA, entao cobrar o mesmo de residente e
internacional ja era um sinal de alerta).
Erro para internacional: $55.065 -> $42.000 (planilha SUPERESTIMAVA em
$13.065). Erro para residente: $55.065 -> $30.750 (superestimava em
$24.315).

============================================================
2) RENSSELAER (RPI, Troy, NY) -- VALOR DESATUALIZADO, CORRIGIDO
============================================================
Fonte oficial: https://finance.rpi.edu/bursar/tuition-and-fee-amounts-and-tuition-adjustments
  Tabela 2026-2027, graduate full-time (ate 16 creditos):
  "$33,150.00 por semestre"
  (acima de 16 creditos: $2.760/credito; part-time: $2.210/credito)

O MS in Business Analytics da Lally e concluido em 1 ano de estudo
full-time (confirmado na pagina do programa), ou seja 2 semestres:
  2 x $33.150 = $66.300
A planilha tinha $60.360 -- que corresponde a uma tabela ANTIGA (~2024,
citada como "$60,360.00 per academic year" em fonte secundaria).
Diferenca: +$5.940.
RPI e PRIVADA -> mesma tarifa para residente e internacional.
'Custo/Crédito' permanece "Verificar" DE PROPOSITO: a cobranca
full-time e FLAT por semestre, nao por credito (mesmo padrao ja
documentado para UGA, SMU e SF State) -- preencher um "custo por
credito" aqui induziria a uma multiplicacao errada no futuro.
'Total Créditos' tambem permanece "Verificar": nao foi encontrada fonte
oficial confirmando o numero de creditos (as buscas ecoaram o numero que
eu mesmo havia colocado na query, o que nao e evidencia).

============================================================
3) LSU (Baton Rouge, LA) -- ACHADO CRITICO DE STATUS
============================================================
Fonte oficial: https://www.lsu.edu/business/sdeis/ms-analytics.php
  "On-campus Master of Science in Analytics program applications are
   closed."  Apenas o programa ONLINE segue aceitando candidaturas.

A planilha classificava esta linha como 'Presencial/Hybrid'. Para uma
candidata internacional que precisa de programa PRESENCIAL (requisito
de visto F-1), uma linha cujo programa presencial esta com inscricoes
FECHADAS nao pode continuar aparecendo como opcao presencial viavel sem
aviso. 'Formato' atualizado para deixar isso explicito.
NAO foi alterado o Tuition Total ($29.770): a pagina oficial nao publica
custo, e como o programa presencial esta fechado, refinar esse numero
tem pouco valor pratico. Fica registrado que o numero segue sem fonte
oficial confirmada.

============================================================
4) SOUTHERN CONNECTICUT STATE (SCSU, New Haven, CT) -- PARCIAL
============================================================
Fonte oficial: https://www.southernct.edu/academics/programs/business-analytics-ms
  "30 credit hours of graduate coursework, with up to an additional 3
   credit hours of prerequisite courses"
'Total Créditos' preenchido com 30 (o nucleo do programa; os ate 3
creditos de pre-requisito sao condicionais e nao se aplicam a todos).
TUITION NAO PREENCHIDA: a pagina do programa nao publica custo, e a
pagina oficial de tuition da SCSU (inside.southernct.edu) expoe os
valores apenas dentro de PDFs nao acessiveis pela ferramenta. Uma busca
retornou "$496/credito (residente CT) / $508/credito (out-of-state)",
mas isso veio de agregadores terceiros, nao de fonte primaria -- e a
diferenca de apenas $12 entre residente e nao-residente e
suficientemente atipica para eu NAO gravar sem confirmacao oficial.
Mantido "Verificar".

============================================================
5) OKLAHOMA STATE (Stillwater, OK) -- NAO RESOLVIDA, SUSPEITA FORTE
============================================================
NAO CORRIGIDA por falta de valor oficial exato (bursar.okstate.edu e
business.okstate.edu retornaram HTTP 403).
SUSPEITA REGISTRADA: a planilha tem 'Tuition Total (Intl)' = $20.786,
mas DUAS fontes secundarias independentes descrevem $20.786 como valor
POR ANO ("tuition fee for the program is USD 20,786 per year"), e o
programa tem 38 creditos / 21 MESES de duracao (~1,75 anos). Se
confirmado, o total real seria da ordem de $36.000-$46.000, nao
$20.786 -- ou seja, a planilha pode estar SUBESTIMANDO em ~2x. Uma
terceira fonte cita "$26.513 em tuition anual", o que conflita com as
outras duas. Como as tres divergem entre si e nenhuma e primaria, o
valor NAO foi alterado -- alterar exigiria escolher arbitrariamente
entre fontes contraditorias.

============================================================
6) GWU (Washington, DC) -- NAO RESOLVIDA
============================================================
NAO CORRIGIDA. studentaccounts.gwu.edu e business.gwu.edu retornaram
HTTP 403 em todas as tentativas. A planilha tem $2.055/credito x 30 =
$61.650. Uma busca sugeriu ~$2.083/credito (~$62.500 total), diferenca
de ~1,4% -- dentro da faixa de um reajuste anual normal, mas sem fonte
primaria acessivel para confirmar qual dos dois e o vigente. Mantido
como esta.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

FIXES = [
    (
        ("George Mason University", "Fairfax"),
        {
            "Custo/Crédito (Residente) USD": 1025,  # CORRECAO: era 1835.5 (mesmo valor que intl)
            "Custo/Crédito (Intl) USD": 1400,  # CORRECAO: era 1835.5
            "Tuition Total (Residente) USD": 30750,  # CORRECAO: era 55065 (superestimava $24.315)
            "Tuition Total (Intl) USD": 42000,  # CORRECAO: era 55065 (superestimava $13.065)
        },
    ),
    (
        ("Rensselaer Polytechnic Institute (RPI)", "Troy"),
        {
            "Tuition Total (Residente) USD": 66300,  # CORRECAO: era 60360 (tabela antiga); 2 x $33.150/semestre
            "Tuition Total (Intl) USD": 66300,  # privada: mesma tarifa
        },
    ),
    (
        ("LSU", "Baton Rouge"),
        {
            "Formato": (
                "ATENÇÃO: inscrições do programa PRESENCIAL estão "
                "FECHADAS (fonte oficial LSU, set/2026) -- apenas a "
                "versão ONLINE segue aceitando candidaturas, o que não "
                "atende ao requisito de presencial para visto F-1"
            ),
        },
    ),
    (
        ("Southern Connecticut State University (SCSU)", "New Haven"),
        {
            "Total Créditos": 30,  # oficial: 30 créditos de núcleo (+ até 3 de pré-requisito, condicionais)
        },
    ),
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [c.value for c in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}
    col_uni = header_to_col["Nome da Universidade"]
    col_cid = header_to_col["Cidade"]
    no_fill = PatternFill(fill_type=None)

    total = 0
    for (uni, cidade), updates in FIXES:
        unknown = [h for h in updates if h not in header_to_col]
        if unknown:
            raise ValueError(f"Colunas inexistentes no schema: {unknown}")

        target = [
            row[0].row
            for row in ws.iter_rows(min_row=2)
            if row[col_uni - 1].value == uni and row[col_cid - 1].value == cidade
        ]
        if len(target) != 1:
            raise ValueError(f"Match ambíguo/inexistente para {uni} ({cidade}): linhas {target}. Abortando.")

        row_idx = target[0]
        print(f"\nLinha {row_idx} — {uni} ({cidade})")
        for header, new_value in updates.items():
            cell = ws.cell(row=row_idx, column=header_to_col[header])
            antes = cell.value
            cell.value = new_value
            cell.fill = no_fill
            print(f"  - {header}:")
            print(f"      Antes:  {antes!r}")
            print(f"      Depois: {new_value!r}")
            total += 1

    wb.save(DB_PATH)
    print(f"\n{total} células corrigidas. Arquivo salvo.")


if __name__ == "__main__":
    main()
