"""
Auditoria linha-por-linha (double-check) + correcao de uma CLASSE DE
DEFEITO sistemica encontrada nas 111 linhas do msba_programs_database.xlsx:
TEXTO GRAVADO EM COLUNAS DE TIPO NUMERICO.

COMO O DEFEITO SE MANIFESTA:
27 colunas do schema (src/types/columns.ts) sao do tipo currency/percent/
number. Ao renderizar essas colunas, formatCellValue() chama
parseNumeric() (msba-dashboard/src/utils/format.ts), que faz:

    value.replace(/[^0-9.\\-]/g, "")

ou seja, REMOVE todo caractere nao-numerico e concatena os digitos que
sobram. Uma celula com texto explicativo vira um NUMERO-LIXO exibido
como moeda/porcentagem, e esse numero-lixo tambem entra na ordenacao e
nos filtros do dashboard -- silenciosamente, sem erro visivel.

Exemplos reais encontrados (valor exibido ANTES desta correcao):
  Wake Forest, Taxa Empregabilidade: "92% (Classe 2025); 93% media 3
    anos..." era exibido como 922.025.933%
  U. of Washington, GMAT Medio: "655 (GMAT) / 324 (GRE) - Class of
    2024..." era exibido como 6.553.242.024
  Texas A&M, Fees Adicionais: "Incluido no total (tuition, taxas,
    licencas de software, livros e orientacao)" era exibido como $70
  Purdue, Ranking QS: "#24 MS in Business Analytics (QS Global, 2026)"
    era exibido como 242.026

VARREDURA COMPLETA: 56 celulas com texto em coluna numerica nas 111
linhas. Destas, a maioria produz parseNumeric() == None -- nesses casos
formatCellValue() cai no fallback String(value) e exibe o texto
corretamente, sem dano (ex.: "3.0 (recomendado)", "18-24", "Pequena",
"Cohorts pequenas", "111-147"). Essas foram DEIXADAS COMO ESTAO: o texto
carrega nuance real e nao corrompe nada.

Este script corrige APENAS as celulas onde parseNumeric() devolve um
numero ERRADO (dano silencioso) ou onde o dado esta na COLUNA ERRADA. A
nuance textual removida de cada celula fica registrada abaixo, campo a
campo, ja que o schema nao tem coluna de notas livres e a coluna
numerica nao pode receber texto (foi exatamente isso que causou o bug).

--- CORRECOES DE TUITION / FEES (foco da auditoria pedida) ---

Texas A&M (Houston, TX) | Fees Adicionais USD
  Antes: "Incluido no total (tuition, taxas, licencas de software,
         livros e orientacao)"  -> exibido como $70
  Depois: 0
  Justificativa: o proprio texto diz que NAO ha custo adicional -- o
  total de $70.000 ja e all-inclusive (confirmado em
  etl_enrich_texas_am.py, linhas 70-73). Zero e o valor numerico que
  representa fielmente "nada a pagar alem do total", e elimina o
  falso "$70".

--- CORRECOES DE CREDITOS (afetam calculo credito x creditos) ---

Loyola Marymount (LMU) | Total Creditos
  Antes: "~30 (waiver ate 6 c/ business background)" -> lido como 306
  Depois: 30
  Confere: 1.802 x 30 = 54.060 == Tuition Total ja cadastrado.
  Nuance perdida: com background em business ha waiver de ate 6
  creditos (o programa pode cair para ~24).

Oakland University | Total Creditos
  Antes: "30 (+ ate 6 foundation)" -> lido como 306
  Depois: 30
  Confere: 1.078,25 x 30 = 32.347,50 ~ 32.348 == Tuition Total.
  Nuance perdida: pode haver ate 6 creditos de foundation ADICIONAIS
  para quem nao tem base em business (elevando a ~36).

Bentley University | Total Creditos
  Antes: "10-16 cursos (c/ waivers)" -> parseNumeric None (nao corrompia,
         mas era incoerente com as demais colunas numericas)
  Depois: 30
  Confere: etl_enrich_bentley.py documenta Tuition Total = 10 cursos x
  $6.335 = $63.350 e Custo/Credito = 6.335/3 = 2.111,67 -> a base de
  calculo e 10 cursos x 3 creditos = 30 creditos.
  Nuance perdida: numero de cursos varia com waivers.

Santa Clara University | Total Creditos
  Antes: "42 unidades" -> lido como 42 (coincidentemente correto)
  Depois: 42  (limpeza de tipo, sem mudanca de valor)

DePaul University | Total Creditos
  Antes: "48 (menos com waivers)" -> lido como 48 (correto)
  Depois: 48  (limpeza de tipo; nuance perdida: waivers reduzem)

--- CORRECOES DE COLUNA ERRADA ---

University of Washington | Tamanho Turma
  Antes: "~61-69% internacional" -> lido como 6.169 alunos
  Depois: "Verificar"
  Justificativa: esse dado e PERCENTUAL DE INTERNACIONAIS, nao tamanho
  de turma -- estava na coluna errada e nao existe coluna de "% intl"
  no schema. Melhor "Verificar" (desconhecido) do que um tamanho de
  turma inventado de 6.169.

--- CORRECOES DE RANKING / EMPREGABILIDADE / GMAT ---

Purdue | Ranking QS: "#24 MS in Business Analytics (QS Global, 2026)"
  -> 242026.  Depois: 24
UW-Madison | Ranking QS: "#7 public master's em business analytics
  (QS 2025)" -> 72025.  Depois: 7
  Nuance perdida: e ranking entre PUBLICAS, nao geral.
Wake Forest | Taxa Empregabilidade: "92% (Classe 2025); 93% media 3
  anos..." -> 922025933.  Depois: 92  (usada a Classe 2025, a mais
  recente e especifica)
Bentley | Taxa Empregabilidade: "94% empregados em 3 meses (citado --
  Verificar)" -> 943.  Depois: 94
  ATENCAO: valor "citado", nao confirmado em fonte primaria.
Dakota State | Taxa Empregabilidade: "92,9% (emprego OU continuacao de
  estudos combinados, nao e taxa de emprego pura)" -> 929.
  Depois: 92.9
  ATENCAO: metodologia DIFERENTE das demais linhas -- combina emprego
  com continuacao de estudos, entao nao e comparavel 1:1.
Wake Forest | GMAT Medio: "685 (GMAT) / 320 (GRE) -- media dos ultimos
  3 anos" -> 6853203.  Depois: 685
U. of Washington | GMAT Medio: "655 (GMAT) / 324 (GRE) -- Class of
  2024, ambos opcionais" -> 6553242024.  Depois: 655
Dakota State | GMAT Medio: "N/A (GRE, nao GMAT): ... 320 ..." -> 320
  Depois: "Verificar"
  Justificativa: o proprio texto diz que NAO ha GMAT (a nota 320 e de
  GRE). Deixar 320 na coluna "GMAT Medio" seria registrar uma nota de
  GRE como se fosse GMAT.

--- CORRECOES DE GPA ENGANOSO ---

USC | GPA Minimo: "Sem minimo formal; candidatos competitivos
  geralmente tem GPA..." -> lido como 3.0
  Depois: "Verificar"
  Justificativa: o texto afirma explicitamente que NAO existe minimo
  formal; exibir "3.0" como se fosse um minimo oficial e enganoso.
U. of Arkansas | GPA Minimo: "3.2 e o limiar de dispensa de GMAT/GRE
  para domesticos; sem minimo geral divulgado" -> lido como 3.2
  Depois: "Verificar"
  Justificativa: 3.2 e o limiar de DISPENSA DE TESTE, nao um minimo de
  admissao -- e o proprio texto diz que o minimo geral nao e divulgado.

--- NAO ALTERADO (parseNumeric devolve None -> texto exibido correto) ---
MIT "111-147" creditos; Wichita State / U. of Houston / UT Dallas /
Texas A&M / UTSA / Dakota State (Duracao em meses, faixas); Bentley
"Pequena" e Cornell "Cohorts pequenas" (Tamanho Turma); U. of Arizona
"Faixa oficial: $65.000-$105.000" (Salario Medio); U. of Washington e
CU Boulder e SMU (Ranking, texto descritivo); UW-Madison GMAT Medio
(so GRE divulgado); e ~20 celulas de GPA Minimo com qualificadores
("3.0 recomendado", "3.0 cumulativo", etc.) que ou parseiam para o
valor correto ou para None.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
VERIFICAR = "Verificar"

# (nome da universidade, cidade) -> {coluna: novo valor}
# Cidade incluida para desambiguar nomes repetidos (ex.: 'Madison' aparece
# tanto em UW-Madison/WI quanto em Dakota State/SD).
FIXES = [
    (("Texas A&M", "Houston"), {"Fees Adicionais USD": 0}),
    (("Loyola Marymount University (LMU)", "Los Angeles"), {"Total Créditos": 30}),
    (("Oakland University", "Rochester (MI)"), {"Total Créditos": 30}),
    (("Bentley University", "Waltham"), {"Total Créditos": 30, "Taxa Empregabilidade %": 94}),
    (("Santa Clara University", None), {"Total Créditos": 42}),
    (("DePaul University", None), {"Total Créditos": 48}),
    (("University of Washington", "Seattle"), {"Tamanho Turma": VERIFICAR, "GMAT Médio": 655}),
    (("Purdue", "West Lafayette"), {"Ranking QS": 24}),
    (("UW-Madison", "Madison"), {"Ranking QS": 7}),
    (("Wake Forest University", "Winston-Salem"), {"Taxa Empregabilidade %": 92, "GMAT Médio": 685}),
    (("Dakota State", "Madison"), {"Taxa Empregabilidade %": 92.9, "GMAT Médio": VERIFICAR}),
    (("USC", "Los Angeles"), {"GPA Mínimo": VERIFICAR}),
    (("U of Arkansas", "Fayetteville"), {"GPA Mínimo": VERIFICAR}),
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}
    col_uni = header_to_col["Nome da Universidade"]
    col_cid = header_to_col["Cidade"]

    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

    total = 0
    for (uni, cidade), updates in FIXES:
        unknown = [h for h in updates if h not in header_to_col]
        if unknown:
            raise ValueError(f"Colunas inexistentes no schema: {unknown}")

        target_rows = [
            row[0].row
            for row in ws.iter_rows(min_row=2)
            if row[col_uni - 1].value == uni
            and (cidade is None or row[col_cid - 1].value == cidade)
        ]
        label = f"{uni}" + (f" ({cidade})" if cidade else "")
        if len(target_rows) == 0:
            raise ValueError(f"Nenhuma linha encontrada para {label}. Abortando.")
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha para {label} (linhas {target_rows}). Abortando.")

        row_idx = target_rows[0]
        print(f"\nLinha {row_idx} — {label}")
        for header, new_value in updates.items():
            cell = ws.cell(row=row_idx, column=header_to_col[header])
            antes = cell.value
            if isinstance(antes, str) and len(antes) > 70:
                antes_disp = antes[:70] + "..."
            else:
                antes_disp = antes
            cell.value = new_value
            cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
            print(f"  - {header}:")
            print(f"      Antes:  {antes_disp!r}")
            print(f"      Depois: {new_value!r}")
            total += 1

    wb.save(DB_PATH)
    print(f"\n{total} células corrigidas. Arquivo salvo.")


if __name__ == "__main__":
    main()
