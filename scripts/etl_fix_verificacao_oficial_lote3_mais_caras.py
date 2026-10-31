"""
Verificacao contra PAGINAS OFICIAIS -- LOTE 3: as 20 linhas MAIS CARAS do
dataset (maior 'Tuition Total (Intl) USD'), a pedido da usuaria.

Racional da amostra: complementa o lote 2 (20 mais baratas). Nas mais
caras, um erro percentualmente pequeno representa muitos milhares de
dolares em valor absoluto.

=========================================================================
PLACAR DO LOTE
=========================================================================
  CONFIRMADAS CORRETAS (bateram exato) ...... 7
  ERRO CORRIGIDO ............................ 1  (Georgetown)
  JA CORRIGIDA EM AUDITORIA ANTERIOR ........ 1  (USC, nesta sessao)
  INCERTA (metodologia) ..................... 1  (Northwestern)
  NAO VERIFICADAS NESTE LOTE ................ 10

=========================================================================
ERRO: GEORGETOWN (McDonough, Washington DC) -- TARIFA DESATUALIZADA
=========================================================================
Fonte: msb.georgetown.edu/msba/admissions-financing/
  Tarifa 2026-2027: $2.758,00 por credito-hora
  30 creditos exigidos -> $82.740

A planilha tinha $2.379/credito e total $71.370 -- tarifa de um ciclo
anterior. Diferenca: +$11.370 (a planilha SUBESTIMAVA em ~14%).
Georgetown e PRIVADA -> mesma tarifa para residente e internacional.

=========================================================================
CONFIRMADAS CORRETAS -- bateram EXATO com a fonte oficial
=========================================================================
MIT (MBAn) .............. $96.884 para 2026-2027 (mitsloan.mit.edu).
                          Bate exato. NOTA util: o Capstone gera um
                          abatimento de tuition de mais de $22.000, o
                          que NAO esta refletido em nenhuma coluna do
                          schema -- vale considerar na comparacao real.
WashU (Olin) ............ OFICIAL: programa presencial de 3 semestres,
                          $33.933/semestre x 3 = $101.799. Bate exato.
                          Fees: $900 (programa) + $100 (governo
                          estudantil) + $350 (saude) = $1.350/semestre
                          x 3 = $4.050. Tambem bate EXATO com a
                          planilha. Bolsa de merito: todo admitido e
                          considerado automaticamente, sem inscricao
                          separada.
Columbia ................ $93.024 pelo programa completo, 36 creditos
                          (~$2.584/credito). Bate exato. Valor e do
                          ciclo 2025-26; 2026-27 ainda nao publicado.
BU (Questrom) ........... OFICIAL 2026-2027: Fall $36.512 + Spring
                          $36.512 + Summer $13.692 = $86.716. Bate
                          EXATO. Fees $1.823 -> tambem bate EXATO.
U. of Miami (Herbert) ... OFICIAL: $2.561/credito x 33 = $84.523. Bate
                          exato. Fees $714/semestre x 2 = $1.438 ->
                          tambem bate EXATO. (Pagina ja refere AY
                          2027-2028.)
Baylor .................. $98.087 mantido. CONFIRMADO na fonte oficial
                          o dado mais importante desta linha, que a
                          planilha JA registrava corretamente: "All
                          admitted students for the on-campus MSBA
                          program automatically receive the same
                          tuition scholarship of 65%" -- e vale
                          inclusive para internacionais. Ou seja, o
                          custo efetivo e ~$34.330, nao $98.087. A
                          planilha ja tem isso em 'Tipos de Bolsa'
                          ("65% tuition scholarship (todos admitidos)")
                          e 'Quao Realista (Intl)' ("Muito Alta
                          (automatico)"). Nenhuma alteracao necessaria.
TCU ..................... $71.663 ja reconciliado na auditoria
                          aritmetica anterior ($2.289 x 30 + $3.000 de
                          fees = $71.670). Sem alteracao.
USC ..................... ja CORRIGIDA nesta sessao de $72.500 para
                          $83.853 (33 unidades x $2.541, total oficial
                          publicado pela Marshall).

=========================================================================
INCERTA -- NAO ALTERADA: NORTHWESTERN (MLDS, Evanston, IL)
=========================================================================
A tarifa por trimestre foi CONFIRMADA na fonte oficial
(mccormick.northwestern.edu): "Tuition for the 2026-2027 academic year
is $23,662 per quarter". Isso valida a base de calculo da planilha.

A duvida e o NUMERO DE TRIMESTRES FATURADOS. A planilha tem $118.310,
que e exatamente 5 x $23.662 (coerente com os 15 meses de duracao ja
cadastrados). Porem a mesma pagina oficial diz que a tuition e cobrada
"em tres trimestres no primeiro ano letivo (outono, inverno e
primavera)" e que no trimestre final deve-se esperar "an approximate
tuition increase of 5%" -- uma leitura de 4 trimestres faturados daria
3 x $23.662 + $24.845 = $95.831.

Faixa plausivel: $95.831 (4 trimestres) a $118.310 (5 trimestres).
NAO alterada: trocar um numero incerto por outro igualmente incerto nao
melhora o dado. Esta e uma boa candidata a email direto ao programa,
junto com Baruch e Oklahoma State.

=========================================================================
NAO VERIFICADAS NESTE LOTE (10)
=========================================================================
Johns Hopkins (403 em carey.jhu.edu; fontes conflitam: planilha tem
  $2.541,67/credito, busca sugeriu $1.890/credito -- diferenca grande,
  merece atencao), Cornell (Johnson ainda nao publicou 2026-2027),
  CMU, UC Berkeley, UCLA, UC Irvine, UC San Diego, Syracuse,
  U. of Denver, American University.
NOTA: UCLA e UC Berkeley tem 'Total Créditos' = "Verificar", entao
essas duas linhas nao passam nem pela checagem aritmetica interna --
sao as unicas do top 20 sem nenhuma forma de validacao cruzada.
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
        ("Georgetown", "Washington"),
        {
            "Custo/Crédito (Residente) USD": 2758,  # tarifa oficial 2026-2027
            "Custo/Crédito (Intl) USD": 2758,  # CORRECAO: era 2379 (ciclo anterior)
            "Tuition Total (Residente) USD": 82740,  # 30 x 2758
            "Tuition Total (Intl) USD": 82740,  # CORRECAO: era 71370 -- subestimava $11.370
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
