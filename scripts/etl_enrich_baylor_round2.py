"""
Data enrichment (round 2): atualiza a linha da 'Baylor' (Hankamer
School of Business, MSBA on-campus) no msba_programs_database.xlsx, a
pedido do usuário: confirmar tuition, bolsa automática (como funciona,
se é o valor máximo ou dá pra conseguir mais), requisitos de aplicação,
e empregabilidade/salário.

FONTES (oficiais, hankamer.baylor.edu, acessadas via WebFetch/
WebSearch em 2026-09-08):
- https://hankamer.baylor.edu/msba/admissions/tuition-financial-awards
- https://hankamer.baylor.edu/msba/admissions
- https://hankamer.baylor.edu/msba

CONFIRMADO SEM MUDANÇA:
- Custo/crédito $2.651 e Tuition Total $98.087 (37 créditos) --
  confirmado via US News (fonte secundária específica, consistente
  com o valor já na planilha; a página oficial de tuition não lista o
  número diretamente, só linka pro site de "One Stop Student Financial
  Services").
- Taxa de aplicação ($0): CONFIRMADO oficial -- "There is no
  application fee for this program."
- GMAT/GRE: já estava 'Verificar' -- CONFIRMADO como "optional"
  (opcional, não obrigatório). Preenchido.
- Pré-requisitos (Python + 3h Stats/QBA): CONFIRMADO oficial, sem
  mudança.

CORRIGIDO:
- Cartas de recomendação: a planilha tinha "3 cartas de recomendação",
  mas a página oficial diz claramente "you can submit up to three
  NAMES as references but only ONE letter of recommendation is
  required" -- ou seja, você indica até 3 pessoas como referência,
  mas só UMA efetivamente escreve e envia a carta. CORRIGIDO de 3
  para 1 carta.
- Documentos Necessários: adicionado o ensaio (não estava listado
  antes) -- exige explicar por que quer cursar business analytics e
  por que na Baylor especificamente.

BOLSA AUTOMÁTICA DE 65% -- COMO FUNCIONA (resposta à pergunta do
usuário sobre se é o valor máximo ou se dá pra conseguir mais):
- É um valor FIXO E UNIFORME para todo mundo admitido no programa
  presencial (não varia por mérito, nota, nem cidadania/status
  internacional) -- não existe uma versão "maior" dessa bolsa
  institucional específica para conseguir através da Baylor.
- É automática: não tem aplicação separada, não tem formulário de
  bolsa. Aparece como crédito na conta do aluno depois que ele se
  matricula nas disciplinas.
- PARA CONSEGUIR MAIS dinheiro além desses 65%, as vias são
  DIFERENTES da bolsa institucional (não é "aumentar" o percentual):
  (1) Graduate assistantships -- uma vez admitida, você fica elegível
  para vagas de emprego estudantil dentro da Baylor (trabalho
  remunerado, não é desconto de tuition); (2) bolsas externas, via a
  base "Scholarship Universe" da própria Baylor; (3) outras formas de
  ajuda financeira (grants, empréstimos, work-study) através do
  Student Financial Services -- mas isso é financiamento, não
  desconto/bolsa.
- CONCLUSÃO: os 65% são o teto do desconto direto de tuition dado
  pela própria Baylor -- qualquer valor adicional viria de fora desse
  mecanismo (trabalho remunerado ou bolsa externa), não de a Baylor
  aumentar a porcentagem.

EMPREGABILIDADE E SALÁRIO -- NÃO CONFIRMADO, campos mantidos
'Verificar': a página do programa MSBA (hankamer.baylor.edu/msba) não
tem nenhum dado de outcomes (emprego, salário, empregadores). Uma
busca geral encontrou uma estatística de "98% empregados em 90 dias,
salário médio $81.160" ligada à página "isba/programs/analytics-major"
-- mas essa página é sobre o "Information Systems & Business
Analytics" como MAJOR DE GRADUAÇÃO (undergrad, "second major only,
Fall 2026"), não o programa de MESTRADO (MSBA). Para não repetir o
erro já cometido antes nesta base (misturar dado de programa
diferente), esse número NÃO foi usado para a linha do MSBA de
pós-graduação. Não encontrei nenhum relatório de emprego
específico do MSBA de pós-graduação publicado pela Baylor.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Baylor",
    "Estado": "TX",
    "Cidade": "Waco",
}

UPDATES = {
    "Documentos Necessários": (
        "Transcripts (não oficiais na inscrição, oficiais após "
        "admissão), resume, 1 carta de recomendação (CORRIGIDO: era "
        "listado como 3 -- você pode indicar até 3 referências, mas "
        "só 1 carta é efetivamente exigida), ensaio (por que "
        "business analytics e por que na Baylor), GMAT/GRE opcional; "
        "internacionais: TOEFL/IELTS/PTE"
    ),
    "GMAT/GRE Obrigatório": "Opcional (não exigido; quem enviar precisa depois pedir o score oficial se for admitido)",
    "Tipos de Bolsa": (
        "Tuition Scholarship de 65% -- valor FIXO E UNIFORME para "
        "TODOS os admitidos no programa presencial (não varia por "
        "mérito/GPA/status internacional; não é 'até 65%', é sempre "
        "65%). Automática, sem aplicação separada -- aparece como "
        "crédito na conta após matrícula nas disciplinas. Não há "
        "versão maior dessa bolsa institucional; dinheiro adicional "
        "vem de fontes separadas: graduate assistantships (emprego "
        "estudantil remunerado, elegível após admissão), bolsas "
        "externas via 'Scholarship Universe' da Baylor, ou outras "
        "formas de ajuda financeira (grants/empréstimos/work-study)."
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
        raise ValueError("Nenhuma linha encontrada para Baylor. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baylor)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
