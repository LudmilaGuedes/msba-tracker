"""
Verificacao contra PAGINAS OFICIAIS -- LOTE 2: as 20 linhas MAIS BARATAS
do dataset (menor 'Tuition Total (Intl) USD'), a pedido da usuaria.

Racional da amostra: sao as candidatas economicamente mais atraentes, e
portanto aquelas em que um numero errado teria maior impacto na decisao --
tanto um valor SUBESTIMADO (que faz a escola parecer viavel quando nao e)
quanto um SUPERESTIMADO (que a descarta injustamente).

=========================================================================
PLACAR DO LOTE (20 linhas)
=========================================================================
  CONFIRMADAS CORRETAS ....... 10
  ERROS CORRIGIDOS ........... 2   (WTAMU, Cal Poly Pomona)
  JA VERIFICADAS NESTA SESSAO . 2  (USF, UCF Professional)
  FLAGRADA ANTES .............. 1  (LSU -- presencial fechado, lote 1)
  NAO RESOLVIDAS .............. 7  (fonte primaria inacessivel/conflitante)

=========================================================================
ERRO 1: WEST TEXAS A&M (WTAMU, Canyon, TX) -- SUBESTIMAVA ~2x
=========================================================================
Fonte oficial:
https://www.wtamu.edu/academics/college-business/academics/graduate-programs/program-costs.html
  "Approximately $490/credit hour"   -> Texas resident
  "Approximately $920/credit hour"   -> Out-of-State US Residents
  "Approximately $980/credit hour"   -> International, Non-US Residents

A planilha tinha $500/credito registrado como custo INTERNACIONAL. Esse
valor e praticamente a tarifa de RESIDENTE DO TEXAS ($490) -- ou seja, a
linha registrava a tarifa domestica na coluna de internacional, fazendo a
WTAMU aparecer como a 2a escola mais barata do dataset inteiro.

Correcao (33 creditos, valor de creditos ja existente na planilha):
  Residente:     33 x $490 =  $16.170  (antes: $16.500, praticamente igual
                                        -- confirma que o numero antigo era
                                        a tarifa domestica)
  Internacional: 33 x $980 =  $32.340  (antes: $16.500 -> subestimava
                                        $15.840, quase o dobro)

RESSALVA: a pagina oficial usa "approximately" nas tres tarifas, e NAO
confirma o total de creditos do programa (os 33 vem do valor que ja
estava na planilha, nao reconfirmado nesta rodada). O ganho aqui e
corrigir uma tarifa que estava categoricamente na coluna errada, nao
atingir precisao ao dolar.

=========================================================================
ERRO 2: CAL POLY POMONA (Pomona, CA) -- SUPERESTIMAVA
=========================================================================
Fonte oficial: https://www.cpp.edu/cba/sgsb/cost/msba-costs.shtml
  "$825 per credit/unit for in-state and out-of-state students
   (includes international students)"
  Custo total do programa: ~$28.050  (= 34 unidades x $825)

A planilha tinha, para INTERNACIONAL, $884,33/unidade e total $30.067 --
superestimava $2.017.
E tinha, para RESIDENTE, $413,33/unidade e total $14.053 -- ou seja
SUBESTIMAVA pela metade, porque assumia uma estrutura "tarifa base baixa
+ sobretaxa de nao-residente" que nao se aplica a este programa. Como a
fonte oficial diz que os $825/unidade valem para in-state E out-of-state
E internacionais, as duas colunas passam a $825/unidade e $28.050 no
total. Efeito liquido: a CPP fica MAIS CARA para residente (14.053 ->
28.050) e um pouco MAIS BARATA para internacional (30.067 -> 28.050).

NOTA IMPORTANTE (contradicao resolvida a favor da fonte primaria): uma
busca secundaria afirmava que nao-residentes pagariam +$471/unidade de
"non-resident supplemental tuition" (o que levaria a $1.296/unidade e
~$44.064 no total). A pagina OFICIAL do proprio programa contradiz isso
explicitamente, dizendo que os $825 ja INCLUEM internacionais. Isso e
coerente com o fato de o MSBA da CPP ser um programa self-support da
Singelyn Graduate School of Business -- no sistema CSU, programas
self-support sao tipicamente isentos da sobretaxa de nao-residente. O
valor de $471 e real, mas e a sobretaxa CSU aplicavel a programas
state-support (ex.: e exatamente o mesmo $471 ja documentado para a SF
State em etl_enrich_sf_state.py, que E state-support). Prevalece a fonte
primaria do programa.

=========================================================================
CONFIRMADAS CORRETAS (nenhuma alteracao necessaria)
=========================================================================
U of Central Oklahoma ... $417/cr nao-residente (oficial) x 32 = $13.344 OK
                          (agregadores citam $15.400-$19.088 de "1o ano",
                          mas esses ja embutem taxas; a coluna e tuition)
Dakota State ............ $698,10/cr nao-residente/intl, tabela oficial
                          2026-27 (dsu.edu) x 30 = $20.943 OK
East Texas A&M .......... nao-residente $7.703 / 9 creditos = ~$856/cr,
                          bate com os $855,89 cadastrados x 30 = $25.677 OK
Miami U Ohio ............ $995/cr oficial, 30 creditos, $29.850 OK
VT Global BA ............ $1.000/cr, mesma tarifa VA e nao-VA (oficial
                          Pamplin), 30 creditos, $30.000 OK
Stony Brook SUNY ........ $1.097/cr out-of-state (specialty master's),
                          x 30 = $32.910 OK
UT Arlington ............ $32.055 total do programa (33h, 16 meses) OK
U of South Carolina ..... OFICIAL sc.edu: $1.100/cr, 30 creditos,
                          $33.000 total, e "The Moore School does not
                          offer an out-of-state rate for the MSBA
                          program" -> tarifa unica, o que valida tambem
                          o fato de intl == residente nesta linha.
                          (Um agregador citava $60.384 out-of-state --
                          DESCARTADO, contradito pela fonte primaria.)
USF / UCF (Professional)  ja verificadas nos enriquecimentos desta sessao

=========================================================================
NAO RESOLVIDAS -- fonte primaria inacessivel ou fontes conflitantes
=========================================================================
Oklahoma State .... bursar.okstate.edu e business.okstate.edu = HTTP 403.
                    Suspeita ja registrada no lote 1: $20.786 parece ser
                    valor ANUAL de um programa de 21 meses.
UMass Amherst ..... isenberg.umass.edu = HTTP 403. Fontes conflitam
                    ($28.675 vs $37.791/ano). Planilha tem $28.675.
Baruch CUNY ....... zicklin.baruch.cuny.edu = erro de certificado TLS.
                    Fontes conflitam: $855/cr vs $1.110/cr out-of-state.
                    Se for $1.110, o total saltaria de $30.780 p/ $39.960.
Binghamton SUNY ... busca indica $1.021/cr out-of-state; planilha tem
                    $991 (diferenca de 3%, possivelmente tabela anterior).
                    Sem fonte primaria acessivel para decidir.
Wichita State ..... a pagina oficial de custos para internacionais
                    (wichita.edu) exibe estimativas com data "Effective
                    Fall 2016" -- desatualizada demais para uso.
Texas State ....... 30 creditos CONFIRMADOS no catalogo oficial (ja
                    corretos na planilha). Tuition nao confirmada: as
                    tarifas gerais de graduate ($357 residente / $767
                    nao-residente) nao batem com os $881,61 cadastrados,
                    mas nao ha pagina de custo especifica do programa.
Oakland Univ. ..... tarifa de graduate nao-residente nao localizada; a
                    unica cifra encontrada ($707,50/cr) e de cursos de
                    nivel 3000+ para RESIDENTES.

Nenhuma dessas foi alterada: sem fonte primaria, mudar o numero apenas
trocaria um valor incerto por outro igualmente incerto.
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
        ("West Texas A&M University (WTAMU)", "Canyon"),
        {
            "Custo/Crédito (Residente) USD": 490,  # oficial "approximately $490/credit hour" (Texas resident)
            "Custo/Crédito (Intl) USD": 980,  # CORRECAO: era 500 (que era a tarifa domestica!)
            "Tuition Total (Residente) USD": 16170,  # 33 x 490
            "Tuition Total (Intl) USD": 32340,  # CORRECAO: era 16500 -- subestimava $15.840
        },
    ),
    (
        ("Cal Poly Pomona", "Pomona"),
        {
            "Custo/Crédito (Residente) USD": 825,  # oficial: tarifa unica, inclui internacionais
            "Custo/Crédito (Intl) USD": 825,  # CORRECAO: era 884.33
            "Tuition Total (Residente) USD": 28050,  # 34 x 825
            "Tuition Total (Intl) USD": 28050,  # CORRECAO: era 30067 -- superestimava $2.017
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
