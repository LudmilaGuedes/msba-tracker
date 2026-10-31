"""
Data enrichment: atualiza APENAS a linha da 'Binghamton SUNY' (School of
Management, Binghamton, NY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Data Analytics (MSDA), presencial, STEM, programa
interdisciplinar (School of Management + Harpur College of Arts and
Sciences + Thomas J. Watson College of Engineering), 30 créditos, 10
meses. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Binghamton, NY.

ACHADO CRITICO DA AUDITORIA: 'Tipos de Bolsa' estava 'SUNY
assistantships', mas a página oficial de FAQ do programa afirma
EXPLICITAMENTE: "We do not offer any assistantships." CORRIGIDO --
removida a referência a assistantships (que não existem neste programa
especifico) e substituída pela informação real de financiamento:
bolsas de mérito automáticas (GPA >= 3.0) e a Clifford D. Clark
Diversity Fellowship (competitiva, exclusiva para cidadãos/residentes
permanentes dos EUA -- NÃO disponível a internacionais). Por
consequência, 'TA/RA Disponível' e 'Processo TA/RA' (estavam
'Verificar') foram preenchidos refletindo essa ausência, e 'Quão
Realista (Intl)' foi REBAIXADO de 'Moderada' para 'Baixa', já que a
única bolsa nomeada encontrada exclui explicitamente internacionais e
não há via de assistantship.

TUITION: a planilha tinha $28.900 (Intl) sem origem rastreável. A
tabela oficial de tuition da Binghamton University (Fall, mais recente
encontrada: ciclo 2023, tarifas SUNY tendem a ser estáveis/de reajuste
lento) mostra $991/crédito (não-residente/internacional) e $471/crédito
(residente NY). CORRIGIDO: 30 créditos x $991 = $29.730 (internacional)
e 30 créditos x $471 = $14.130 (residente, preenchido). 'Fees
Adicionais' preenchido com a soma das taxas obrigatórias por semestre
(Academic Excellence, Activity, Health Services, Recreation,
Technology, Transportation, College Fee, etc. -- todas com teto a
partir de 9+ créditos) x 2 semestres (Fall+Spring, os únicos com a
estrutura completa de fees detalhada oficialmente) = $2.470. Não inclui
os termos de Winter/Summer (3+3 créditos cada, que fecham os 30
créditos/ano) por não haver confirmação de que a mesma fee cheia se
aplica a esses termos parciais.

'Seguro Saúde (programa)' permanece 'Verificar': o único valor
encontrado ($2.260/ano) é uma linha combinada de "seguro saúde
internacional E service fees", sem separação confiável entre os dois
conceitos -- não usado para não misturar categorias do schema.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.binghamton.edu/academics/programs/data-analytics/
  admissions.html (GPA 3.0, GRE/GMAT opcional, TOEFL 80/IELTS 6.5,
  experiência profissional recomendada mas não obrigatória)
- https://www.binghamton.edu/academics/programs/data-analytics/faq.html
  (confirmação direta de que NÃO há assistantships; estrutura de
  créditos por termo; pré-requisitos quantitativos; Clark Fellowship)
- https://www.binghamton.edu/student-accounts/tuition-fees/
  tuition-semester/fall/graduate.html (tabela oficial de tuition e fees
  por crédito, ciclo Fall 2023 -- mais recente disponível)
- Busca web: aluguel de quarto privativo perto do campus, Binghamton NY
  (múltiplos anúncios convergentes em torno de $600-700)
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
    "Nome da Universidade": "Binghamton SUNY",
    "Estado": "NY",
    "Cidade": "Binghamton",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 991,  # oficial, mais recente encontrado (Fall 2023)
    "Custo/Crédito (Residente) USD": 471,  # oficial, mais recente encontrado (Fall 2023)
    "Tuition Total (Intl) USD": 29730,  # CORREÇÃO: planilha tinha $28.900 sem origem rastreável; 30 créditos x $991
    "Tuition Total (Residente) USD": 14130,  # 30 créditos x $471
    "Fees Adicionais USD": 2470,  # taxas obrigatórias por semestre (Fall+Spring, capped 9+ créditos) somadas
    "Aluguel Privativo/mês USD": 650,  # quarto privativo perto do campus (múltiplos anúncios convergentes)
    "Tipos de Bolsa": (
        "CORREÇÃO: o programa NÃO oferece assistantships (confirmado "
        "oficialmente). Bolsas de mérito automáticas para GPA ≥ 3.0; "
        "Clifford D. Clark Diversity Fellowship (competitiva, exclusiva "
        "para cidadãos/residentes permanentes dos EUA, deadline 15 "
        "jan — NÃO disponível a internacionais); bolsas externas"
    ),
    "Quão Realista (Intl)": (
        "Baixa — o programa não oferece assistantships, e a única "
        "bolsa nomeada encontrada (Clark Fellowship) exclui "
        "explicitamente candidatos internacionais"
    ),
    "TA/RA Disponível": "Não — o programa não oferece assistantships (confirmado oficialmente pela FAQ)",
    "Processo TA/RA": "N/A — sem assistantships neste programa",
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": "Opcional (aceito, mas não exigido)",
    "Experiência Profissional": (
        "Não obrigatória, mas recomendada (2+ anos em ambiente de "
        "negócios/indústria); internacionais devem ter ao menos 1 ano, "
        "preferência para 2+"
    ),
    "Matérias Pré-Requisito": (
        "Cálculo (diferencial, para otimização), probabilidade e "
        "estatística, programação (estruturas de dados e algoritmos "
        "básicos); não exige diploma STEM prévio, mas familiaridade "
        "com essas áreas"
    ),
    "Testes Inglês Aceitos": "TOEFL mín. 80 / IELTS mín. 6.5 / Duolingo mín. 105 / PTE mín. 53",
    "Documentos Necessários": "Histórico acadêmico, 2 cartas de recomendação, resume, personal statement",
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
        raise ValueError("Nenhuma linha encontrada para Binghamton SUNY (Binghamton, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Binghamton SUNY, Binghamton, NY)")

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
