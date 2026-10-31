"""
Data enrichment (round 2): atualiza a linha da 'Notre Dame' (Mendoza
College of Business) no msba_programs_database.xlsx.

CORREÇÃO DE ERRO INTRODUZIDO NO LOTE ANTERIOR
(etl_enrich_application_fees_batch1.py): aquele lote, baseado num
resumo de WebSearch sobre a página geral do programa, mudou a 'Taxa
Aplicação USD' de $50 para $100 -- ISSO ESTAVA ERRADO. Ao acessar
diretamente a página oficial de aplicação (fonte primária, não um
resumo de busca), o texto exato confirma:

  "$50 non-refundable application fee is the final requirement in the
  application process."
  "[Early Action] your application fee will be waived, and upon
  invitation, you'll receive priority interview scheduling."

Ou seja, o valor CORRETO sempre foi $50 (o valor original antes do
round anterior já estava certo; o "100" era o erro). O campo
'Documentos Necessários' já mencionava corretamente "taxa de $50" --
a inconsistência entre os dois campos (um dizia 50, outro tinha sido
mudado para 100) é o que expôs o erro ao revisar a fundo agora.
REVERTIDO para 50.

LIÇÃO: nos lotes de pesquisa em massa (9 lotes paralelos), resumos de
WebSearch por vezes conflitam com a página oficial quando checada
diretamente via WebFetch -- fetch direto da página de aplicação é mais
confiável que um resumo de busca para valores específicos como taxa.

OUTRAS CORREÇÕES/ENRIQUECIMENTOS (fonte: WebFetch direto de
https://mendoza.nd.edu/graduate-programs/business-analytics-msba/apply/,
em 2026-09-08, citações exatas obtidas da página):

- Cartas de recomendação: a página confirma "One recommendation is
  required" -- CORRIGIDO de "2 cartas" (erro da pesquisa original)
  para 1.
- Documentos adicionais não documentados antes: currículo (1 página),
  "Applicant Snapshot" (apresentação em PDF de 4 slides destacando
  qualidades pessoais além do histórico acadêmico/profissional), e uma
  "avaliação em vídeo" (perguntas feitas pela equipe de admissão para
  avaliação holística) -- ADICIONADOS.
- GMAT/GRE: regra correta é dispensa para quem tem ou terá diploma de
  graduação de QUALQUER universidade dos EUA (não só Notre Dame, como
  a planilha dizia antes) -- CORRIGIDO. Internacionais sem diploma dos
  EUA devem enviar GMAT ou solicitar isenção. Scores aceitos só até 5
  anos de idade.
- Prova de inglês: dispensada para quem completou pelo menos 6
  semestres ou 9 trimestres integrais em universidade de instrução em
  inglês -- detalhe adicionado.
- Pré-requisitos: página diz "One Statistics course" + "Two business
  courses (Accounting, Finance, Management, or Marketing)" --
  CORRIGIDO (a planilha tinha "1 estatística + 1 contabilidade + 1
  negócio adicional", que restringia incorretamente um dos dois cursos
  de negócio à contabilidade especificamente).
- Entrevista: confirmado que faz parte do processo para todos os
  candidatos (não só Early Action) -- Early Action só dá agendamento
  PRIORITÁRIO, não é a única via de entrevista -- CORRIGIDO (antes a
  planilha dizia que não ficava claro se era obrigatória nos demais
  rounds).
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
    "Nome da Universidade": "Notre Dame",
    "Estado": "IN",
    "Cidade": "Notre Dame",
}

UPDATES = {
    "Taxa Aplicação USD": 50,  # CORREÇÃO: revertendo erro do lote anterior (era 100, correto é 50)
    "Documentos Necessários": (
        "Transcripts oficiais de todas as instituições (com tradução "
        "para inglês se aplicável), 1 carta de recomendação (CORRIGIDO: "
        "era listado como 2), currículo (1 página), ensaio de até 500 "
        "palavras (planos de carreira pós-mestrado), 'Applicant "
        "Snapshot' (apresentação em PDF de 4 slides sobre qualidades "
        "pessoais), avaliação em vídeo (perguntas da equipe de "
        "admissão), GMAT/GRE (dispensado para quem tem/terá diploma de "
        "graduação de universidade dos EUA; demais internacionais "
        "enviam GMAT ou pedem isenção; scores válidos só até 5 anos), "
        "TOEFL/IELTS/PTE/Duolingo (internacionais, dispensável com 6+ "
        "semestres ou 9+ trimestres integrais em universidade de "
        "instrução em inglês), taxa de $50 (isenta p/ Early Action)"
    ),
    "GMAT/GRE Obrigatório": (
        "Dispensado para quem tem ou terá diploma de graduação de "
        "qualquer universidade dos EUA (CORRIGIDO: não é exclusivo a "
        "egressos da própria Notre Dame). Demais internacionais devem "
        "enviar GMAT ou solicitar isenção. Scores aceitos até 5 anos "
        "de idade. Sem nota mínima publicada."
    ),
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/PTE Academic/Duolingo (internacionais) -- "
        "dispensado para quem completou pelo menos 6 semestres ou 9 "
        "trimestres integrais numa universidade de instrução em "
        "inglês. Nota mínima específica do MSBA não publicada "
        "separadamente (mín. TOEFL 80/IELTS 7.0 é a prática usual da "
        "escola, confirmada para o MBA)."
    ),
    "Matérias Pré-Requisito": (
        "1 curso de estatística + 2 cursos de negócios dentre: "
        "contabilidade, finanças, gestão ou marketing (CORRIGIDO: não "
        "é obrigatoriamente 1 de contabilidade + 1 outro, e sim "
        "quaisquer 2 dessas 4 áreas)"
    ),
    "Entrevista": (
        "Faz parte do processo de admissão para todos os candidatos "
        "(CORRIGIDO: não é exclusiva do Early Action). Candidatos de "
        "Early Action recebem agendamento PRIORITÁRIO de entrevista, "
        "mediante convite."
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
        raise ValueError("Nenhuma linha encontrada para Notre Dame. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Notre Dame)")

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
