"""
Data enrichment: atualiza APENAS a linha da 'Oakland University' (School
of Business Administration, Rochester Hills, MI) no
msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato ON-CAMPUS (presencial
-- distinto da versão 100% online, que usa tarifa in-state para todos
independente de residência), STEM, 30 créditos (+ até 6 de foundation,
que não contam para o diploma), 12 meses. Deep research em: custo total
real de tuition para internacionais, bolsas/assistantships, requisitos
de GPA/GMAT/GRE/inglês, e custo de vida em Rochester Hills, MI.

TUITION: PREENCHIDO (estava 'Verificar' em ambos os campos). A página
oficial de Tuition Rates lista as tarifas gerais de pós-graduação
("special graduate, certification, master's..." status): $951,50/
crédito (residente Michigan) / $1.078,25/crédito (não-residente). A
página não confirma uma tarifa "specialty master's" alternativa
especificamente para o MSBA on-campus (diferente da versão online, que
usa tarifa in-state para todos) -- por isso as tarifas gerais foram
usadas como a melhor fonte oficial disponível. PREENCHIDO 'Custo/
Crédito' e 'Tuition Total' (30 créditos): $28.545 (residente) / $32.348
(não-residente/Intl).

'GMAT/GRE Obrigatório' CORRIGIDO: a planilha tinha 'Exigido', mas a
página oficial do programa confirma que o teste é DISPENSÁVEL sob
vários critérios (pós-graduação prévia, GPA forte de graduação --
3.0+ AACSB, 3.2+ não-AACSB --, certificações profissionais como
CPA/CFA/CFP/FRM, ou experiência profissional relevante).

'Taxa Empregabilidade %' e 'Salário Médio Pós USD' preenchidos com
dados OFICIAIS e ESPECÍFICOS do MSBA (coorte 2024-25): 59% empregados
em até 4 meses, salário médio $76.200 -- baseado em uma amostra pequena
(26 egressos rastreados, 58% taxa de resposta à pesquisa), o que é
documentado aqui como ressalva de tamanho amostral.

Custo de vida: 'Transporte/mês' preenchido com $0 -- o SMART bus (ônibus
regional, conecta a Detroit e arredores) é GRATUITO para alunos da
Oakland University mediante apresentação do OneCard. 'Seguro Saúde
(programa) USD' preenchido com $2.382/ano -- prêmio oficial 2026-27 do
plano UnitedHealthcare obrigatório para internacionais F-1/J-1.
'Aluguel Privativo/mês' preenchido com $800 (faixa $600-1.200
encontrada para moradia fora do campus em Rochester Hills, perto do
campus).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.oakland.edu/financialaid/costs/tuition-rates/ (tarifas
  oficiais gerais de pós-graduação, residente e não-residente)
- https://www.oakland.edu/business/graduate-executive-programs/
  master-of-science-in-business-analytics/ (formato on-campus/online,
  30+6 créditos, GPA/GMAT/GRE com critérios de dispensa, TOEFL,
  outcomes oficiais da coorte 2024-25, SAS Analytics Tier 2)
- Busca web: taxa de aplicação $45; deadlines recomendados 1º mar
  (Fall)/1º nov (Winter) para bolsas, admissão rolling
- Busca web: SMART bus gratuito para alunos OU; seguro saúde $2.382/
  ano (oficial 2026-27); aluguel fora do campus em Rochester Hills, MI
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
    "Nome da Universidade": "Oakland University",
    "Estado": "MI",
    "Cidade": "Rochester (MI)",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1078.25,  # oficial, tarifa geral não-residente
    "Custo/Crédito (Residente) USD": 951.50,  # oficial, tarifa geral residente
    "Tuition Total (Intl) USD": 32348,  # 30 créditos x $1.078,25
    "Tuition Total (Residente) USD": 28545,  # 30 créditos x $951,50
    "GMAT/GRE Obrigatório": (
        "CORREÇÃO: dispensável com pós-graduação prévia, GPA de "
        "graduação ≥3.0 (AACSB) ou ≥3.2 (não-AACSB), certificações "
        "profissionais (CPA/CFA/CFP/FRM), ou experiência profissional "
        "relevante"
    ),
    "GPA Mínimo": (
        "3.0 (com mestrado prévio) / 3.0+ (graduação em escola AACSB) "
        "/ 3.2+ (graduação em escola não-AACSB) / 2.75+ com 3+ anos de "
        "experiência profissional"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL exigido para não-nativos, salvo diploma de bacharelado "
        "dos EUA ou 24+ créditos cursados em instituição dos EUA"
    ),
    "Taxa Aplicação USD": 45,
    "Deadline Aplicação": "Recomendado: 1º mar (Fall) / 1º nov (Winter) para consideração de bolsas; admissão contínua (rolling) fora desses prazos",
    "Taxa Empregabilidade %": 59,  # coorte 2024-25 do MSBA, dentro de 4 meses (amostra pequena: 26 egressos rastreados, 58% resposta)
    "Salário Médio Pós USD": 76200,  # coorte 2024-25 do MSBA (mesma ressalva de amostra pequena)
    "Transporte/mês USD": 0,  # SMART bus gratuito para alunos OU (OneCard)
    "Seguro Saúde (programa) USD": 2382,  # oficial 2026-27, obrigatório para internacionais F-1/J-1
    "Aluguel Privativo/mês USD": 800,  # faixa $600-1.200 encontrada, Rochester Hills MI (perto do campus)
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
        raise ValueError("Nenhuma linha encontrada para Oakland University (Rochester (MI), MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Oakland University, Rochester (MI), MI)")

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
