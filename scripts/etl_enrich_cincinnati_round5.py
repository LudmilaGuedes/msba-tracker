"""
Data enrichment (round 5): atualiza APENAS a linha da 'U of Cincinnati'
(Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

FONTE DESTA RODADA: resposta por EMAIL DIRETO da admissão/financial aid da
UC para a usuaria, perguntando especificamente sobre tuition e bolsas do
MSBA. E uma fonte de PRIMEIRA MAO (a propria universidade, respondendo a
pergunta especifica da candidata) -- pesa mais que qualquer calculadora de
terceiros ou pagina generica de bursar usada em rodadas anteriores.

CORRECAO CRITICA DE TUITION -- reverte a decisao do round 1/round 2:
o email diz explicitamente: "As an international student, your tuition
rate is an additional $250 per credit hour" -- ou seja, o aluno
internacional paga a tarifa de residente ($1.145/credito, "Business (22)
Graduate Premium Programs", ja confirmada na tabela do Bursar em rodadas
anteriores) MAIS um sobretaxa de $250/credito. Isso e EXATAMENTE o que a
tabela geral do Bursar ja indicava (e que tinha sido descartada no round 1
por nao bater com o total de $33.195, corroborado a epoca por 2
calculadoras de terceiros). Com uma fonte de primeira mao agora
confirmando a mesma logica da tabela geral, a conclusao muda: o total de
$33.195 (usado nas rodadas 1-4) estava ERRADO -- os calculadores de
terceiros provavelmente nao capturaram a sobretaxa "Business Premium"
especifica da Lindner corretamente.

NOVO CALCULO (40 creditos, confirmado em rodadas anteriores):
  - Residente (Ohio): $1.145/credito x 40 = $45.800
  - Internacional: ($1.145 + $250)/credito = $1.395/credito x 40 = $55.800
'Custo/Crédito' passa de valor DERIVADO (829,88, back-calculado do total
antigo) para o valor REAL cobrado por credito, agora confirmado por fonte
direta.
'Caminho Tuition Residente' permanece "Não" (dado original, nao alterado
por nenhuma rodada) -- o email nao menciona nenhum caminho para
internacional pagar a tarifa de residente, entao permanece consistente.

CORRECAO DE BOLSAS -- Graduate Assistantships NAO cobrem tuition:
o email esclarece: "Graduate Assistantships: Apply through Handshake
after admission... These roles do NOT cover tuition but pay an hourly
rate." Isso CORRIGE o texto do round 4, que (baseado na pagina GERAL de
politica de assistantship da UC, grad.uc.edu) descrevia GA como incluindo
"tuition scholarship parcial ou integral" -- a fonte direta da
Lindner/admissoes para o MSBA especificamente contradiz isso: os cargos
via Handshake pagam so por hora, sem bolsa de tuition embutida. Mantido o
processo (Handshake, apos admissao+deposito) ja documentado, mas corrigida
a cobertura.

NOVO -- University Graduate Scholarships (nao documentado antes):
bolsa distinta do GA, automaticamente considerada na admissao (sem
formulario separado), valor $2.500-$8.500, variando por competitividade
da coorte e ano. Isso e DIFERENTE dos nomes "Graduate Incentive Award
(GIA)" e "Graduate Assistant Scholarship (GAS)" que ja estavam
cadastrados desde antes desta sessao de trabalho (dado original, nao
verificado por mim) -- o email nao usa esses nomes. Mantidos os nomes
antigos (GIA/GAS) no campo 'Graduate Assistantships' por nao terem sido
eles proprios contraditos (apenas a cobertura de tuition foi corrigida),
mas o campo 'Tipos de Bolsa' foi atualizado com o nome e valores que a
propria universidade usou na resposta a usuaria (fonte mais recente e
mais direta).

NOVO -- recursos adicionais e emprestimos: email menciona recursos de
busca de bolsa de terceiros (IEFA, InternationalScholarships.com,
InternationalStudent.com) e que emprestimos americanos sao dificeis para
internacionais (poucos credores, normalmente exigem co-signatario cidadao
americano) -- mas a UC aceita emprestimos do pais de origem do aluno.
Incorporado ao campo 'Tipos de Bolsa' de forma resumida.

'Quão Realista (Intl)' mantido como "Moderada", mas com nota de que os
valores de bolsa por merito sao modestos frente ao novo total de tuition
($2.500-$8.500 cobre so ~4-15% dos $55.800), e que GA nao ajuda com
tuition (e so um trabalho de meio periodo remunerado por hora).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Cincinnati",
    "Estado": "OH",
    "Cidade": "Cincinnati",
}

TOTAL_CREDITOS = 40
CUSTO_CREDITO_RESIDENTE = 1145.0
CUSTO_CREDITO_INTL = 1145.0 + 250.0  # 1395.0

UPDATES = {
    "Custo/Crédito (Residente) USD": CUSTO_CREDITO_RESIDENTE,
    "Custo/Crédito (Intl) USD": CUSTO_CREDITO_INTL,
    "Tuition Total (Residente) USD": round(CUSTO_CREDITO_RESIDENTE * TOTAL_CREDITOS),  # 45800
    "Tuition Total (Intl) USD": round(CUSTO_CREDITO_INTL * TOTAL_CREDITOS),  # 55800
    "Tipos de Bolsa": (
        "University Graduate Scholarships: automaticamente consideradas "
        "na admissão (sem formulário separado), $2.500-$8.500, variando "
        "por competitividade da coorte e ano (fonte: resposta direta da "
        "admissão da UC por email, set/2026). Recursos externos citados "
        "pela própria UC: Online Scholarship Listing, IEFA "
        "(iefa.org), InternationalScholarships.com, "
        "InternationalStudent.com. Empréstimos americanos são difíceis "
        "para internacionais (poucos credores, geralmente exigem "
        "co-signatário cidadão dos EUA) -- UC aceita empréstimos do país "
        "de origem do aluno"
    ),
    "Graduate Assistantships": (
        "Graduate Incentive Award (GIA) e Graduate Assistant Scholarship "
        "(GAS) -- bolsas financiadas pela universidade (nomes de fonte "
        "anterior à esta sessão de trabalho). CORREÇÃO (email direto da "
        "admissão da UC, set/2026): as vagas de Graduate Assistantship "
        "divulgadas via Handshake NÃO cobrem tuition -- pagam apenas por "
        "hora. A política geral de assistantship da UC (stipend mínimo "
        "$785,79 quinzenal + tuition scholarship parcial/integral, "
        "elegibilidade ao GSHI) pode não se aplicar da mesma forma às "
        "vagas específicas oferecidas ao MSBA via Handshake"
    ),
    "Processo TA/RA": (
        "Candidato precisa estar admitido, com programa confirmado e "
        "taxa de matrícula paga (quando aplicável) antes de poder se "
        "candidatar a vagas via Handshake. Essas vagas NÃO cobrem "
        "tuition -- pagam apenas por hora (confirmado por email direto "
        "da admissão da UC, set/2026)"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- University Graduate Scholarships são automáticas "
        "mas modestas ($2.500-$8.500, cobrem só ~4-15% do novo total de "
        "$55.800 intl); Graduate Assistantships via Handshake NÃO ajudam "
        "com tuition (só pagam por hora); empréstimos americanos difíceis "
        "para internacionais sem co-signatário; UC aceita empréstimo do "
        "país de origem"
    ),
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
        raise ValueError("Nenhuma linha encontrada para U of Cincinnati (Cincinnati, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Cincinnati, Cincinnati, OH)")

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
