"""
Auditoria de consistencia de TUITION em todas as 111 linhas do
msba_programs_database.xlsx + correcao das divergencias reais encontradas.

METODOLOGIA DA AUDITORIA (rodada sobre TODAS as linhas):
  1. Checagem aritmetica: Custo/Crédito x Total Créditos == Tuition Total?
     (tolerancia: maior entre $60 e 2%)
  2. Mesma checagem somando 'Fees Adicionais USD'
  3. Sanidade: Tuition Intl >= Tuition Residente?
  4. Outliers: Tuition Intl fora da faixa $8.000-$130.000?
  5. Publicas cobrando intl == residente (suspeito)

RESULTADO: checagens 3, 4 e 5 passaram limpas em 111/111 linhas. A
checagem 1 levantou 6 linhas; investigadas uma a uma:

  - TCU e U of Arkansas: RECONCILIAM ao somar 'Fees Adicionais'
    (TCU: 2289x30=68.670 + 3.000 = 71.670 ~ 71.663 registrado;
     Arkansas: 1974,78x30=59.243 + 3.340 = 62.583 exato). OK, sem acao.

  - UGA: divergencia INTENCIONAL e ja documentada em
    etl_enrich_uga.py -- programa full-time usa a tarifa FLAT de tempo
    integral ($8.940/semestre residente, $19.068/semestre intl) x 2
    semestres, nao a tarifa por credito x 33. OK, sem acao.

  - SF State: divergencia INTENCIONAL e ja documentada em
    etl_enrich_sf_state.py -- 'Tuition Total' usa a estimativa OFICIAL
    de custo total do programa divulgada pela universidade ($31.000
    residente / $41.000 intl), que ja embute fees acumuladas ao longo
    de 3-4 semestres. OK, sem acao.

  - SMU: divergencia REAL mas o valor cadastrado esta CORRETO -- faltava
    documentacao. Confirmado na tabela oficial do Bursar da SMU
    (2026-2027): a tarifa de $1.570/credito vale apenas para 1-11
    horas (PART-TIME); para 12-18 horas (FULL-TIME, que e o caso do
    MSBA) a cobranca e FLAT de $25.000/semestre. O MSBA de 33 creditos
    em 2 semestres full-time = 2 x $25.000 = $50.000, exatamente o
    valor ja cadastrado. As fees tambem reconciliam: Program Fee
    full-time $1.386/semestre x 2 = $2.772, valor ja cadastrado.
    ACAO: NENHUMA alteracao na planilha -- todos os valores ja estao
    corretos. A explicacao NAO foi escrita dentro da celula
    'Custo/Crédito' porque essa coluna e do tipo "currency" e
    parseNumeric() (msba-dashboard/src/utils/format.ts) remove todo
    caractere nao-numerico de strings: um texto ali seria concatenado
    em um numero-lixo e exibido como moeda no dashboard, quebrando
    ordenacao e filtro. Esta docstring e o registro do achado -- se uma
    auditoria futura reclamar que 1.570 x 33 = 51.810 != 50.000, a
    resposta esta aqui: a tarifa por credito e part-time, o total e
    flat full-time.

  - USC: ERRO REAL. Corrigido nesta rodada (detalhes abaixo).

CORRECAO USC (Marshall School of Business, Los Angeles, CA):
Fonte oficial:
https://www.marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics/tuition-and-fees
  - "Current 2025-2026 tuition rate is $2,541 per unit"
  - 33 unidades
  - Total de tuition divulgado pela propria pagina: $83.853,00
    (identico nas opcoes de 3 ou 4 semestres)
  - Fees: $4.046 (opcao 3 semestres) a $4.243 (opcao 4 semestres)
  - Total com fees: $87.899 (3 sem.) a $88.096 (4 sem.)

A planilha tinha 'Tuition Total (Intl)' e '(Residente)' = $72.500, valor
sem origem rastreavel em nenhum script do repositorio e que NAO bate com
a tarifa por unidade ja cadastrada (2.541 x 33 = 83.853, diferenca de
-$11.353). Como a fonte oficial publica explicitamente o total de
$83.853 E ele reconcilia exatamente com a tarifa por unidade ja
cadastrada, o $72.500 foi considerado incorreto/desatualizado e
substituido.
USC e PRIVADA -> mesma tarifa para residente e internacional, entao
ambos os totais recebem $83.853 e 'Custo/Crédito (Residente)' (que
estava "Verificar") recebe $2.541, igual ao intl.
'Fees Adicionais' refinado de $4.000 (arredondado, sem origem) para
$4.046 (valor oficial da opcao de 3 semestres, que e a estrutura
padrao ja registrada em 'Duração (meses)' = 12).

LIMITACAO DESTA AUDITORIA (importante): a checagem e de CONSISTENCIA
INTERNA. Ela detecta linhas onde os campos se contradizem, mas NAO
detecta uma linha onde 'Custo/Crédito' e 'Tuition Total' estejam ambos
errados de forma coerente (ex.: ambos de um ano letivo antigo). Para
esses casos so a reconferencia na fonte oficial, uma escola por vez,
resolve.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

# Cada entrada: (chave de match, dict de updates)
TARGETS = [
    (
        {"Nome da Universidade": "USC", "Estado": "CA", "Cidade": "Los Angeles"},
        {
            "Custo/Crédito (Residente) USD": 2541,  # privada: mesma tarifa que intl (estava "Verificar")
            "Tuition Total (Intl) USD": 83853,  # CORRECAO: 72500 -> 83853 (33 x 2541, total oficial)
            "Tuition Total (Residente) USD": 83853,  # CORRECAO: idem, privada = mesma tarifa
            "Fees Adicionais USD": 4046,  # refinado de 4000 p/ valor oficial (opcao 3 semestres)
        },
    ),
    # SMU: NENHUMA alteracao. Todos os valores ja estao corretos (ver
    # docstring). Deliberadamente NAO foi escrita uma nota explicativa
    # dentro de 'Custo/Crédito' porque essa coluna e do tipo "currency" no
    # schema e parseNumeric() (utils/format.ts) remove todo caractere
    # nao-numerico -- um texto ali viraria um numero-lixo no dashboard e
    # quebraria ordenacao/filtro. A documentacao vive aqui neste script.
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}
    no_fill = PatternFill(fill_type=None)

    for match, updates in TARGETS:
        unknown = [h for h in updates if h not in header_to_col]
        if unknown:
            raise ValueError(f"Colunas inexistentes no schema: {unknown}")

        target_rows = []
        for row in ws.iter_rows(min_row=2):
            if all(row[header_to_col[k] - 1].value == v for k, v in match.items()):
                target_rows.append(row[0].row)

        label = f"{match['Nome da Universidade']} ({match['Cidade']}, {match['Estado']})"
        if len(target_rows) == 0:
            raise ValueError(f"Nenhuma linha encontrada para {label}. Abortando.")
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha para {label} (linhas {target_rows}). Abortando.")

        row_idx = target_rows[0]
        print(f"\nLinha alvo confirmada: {row_idx} ({label})")

        for header, new_value in updates.items():
            cell = ws.cell(row=row_idx, column=header_to_col[header])
            antes = cell.value
            cell.value = new_value
            cell.fill = no_fill
            print(f"  - {header}:")
            print(f"      Antes: {antes!r}")
            print(f"      Depois: {new_value!r}")

    wb.save(DB_PATH)
    print("\nArquivo salvo.")


if __name__ == "__main__":
    main()
