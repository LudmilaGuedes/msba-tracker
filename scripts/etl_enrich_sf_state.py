"""
Data enrichment: atualiza APENAS a linha da 'SF State' (Lam Family
College of Business, San Francisco, CA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM, 30
unidades (10 disciplinas), completável em 3-4 semestres. Deep research
em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida em
San Francisco, CA.

TUITION: CORRIGIDO. A planilha tinha $21.000 (Intl) sem origem
rastreável. A página oficial de tuition do College of Business (ciclo
Fall 2026) fornece dois níveis de dado: (1) tarifas por unidade -- $321
(residente) / $792 (não-residente/intl, = $321 base + $471 de fee
suplementar para não-residentes) -- e (2) uma estimativa de CUSTO TOTAL
DO PROGRAMA já divulgada pela própria universidade: $31.000 (residente)
/ $41.000 (não-residente). Como o total oficial inclui fees adicionais
(Off Campus Program Fee de $700-1.200/semestre, taxas de campus, etc.)
que se acumulam ao longo dos 3-4 semestres do programa -- e não apenas
30 unidades x tarifa/unidade --, PREENCHIDO 'Custo/Crédito' com as
tarifas por unidade e CORRIGIDO 'Tuition Total' para os valores totais
oficiais ($31.000 residente / $41.000 internacional, este último
substituindo o $21.000 sem origem rastreável). 'Fees Adicionais'
permanece 'Verificar' -- o total oficial já embute essas taxas, e
tentar extrair um valor "adicional" separado dobraria a contagem.

Custo de vida: 'Transporte/mês' preenchido com $0 -- o "Gator Pass"
(taxa de trânsito obrigatória, cobrada como parte das taxas de campus)
dá a todo aluno da SF State acesso ILIMITADO e gratuito ao MUNI
(ônibus/trens, exceto bondes turísticos) no ponto de uso; como essa fee
provavelmente já está embutida no 'Tuition Total' oficial acima, contar
um custo mensal de transporte separado duplicaria a cobrança.
'Seguro Saúde (programa)' preenchido com $2.385/ano -- tarifa oficial
corrente do plano SHIP para internacionais (2025-26, via Student
Medicover/UnitedHealthcare). 'Aluguel Privativo/mês' preenchido com
$1.100 -- San Francisco é uma das cidades mais caras dos EUA; esse
valor reflete o extremo "econômico" da faixa de mercado encontrada
($900-1.300 em bairros mais acessíveis, dividindo moradia com colegas),
não a média geral da cidade (que é muito mais alta para uma unidade
inteira).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://cob.sfsu.edu/resources/graduate-programs/graduate-program-fees
  (tuition oficial Fall 2026: tarifas por unidade e custo total
  estimado do programa)
- https://cob.sfsu.edu/graduate-programs/business-analytics (formato,
  duração 3-4 semestres, currículo, STEM)
- https://cob.sfsu.edu/admissions (deadlines: Fall 1º mai, Spring 1º
  out)
- Busca web: GPA mínimo 3.0 (admissão excepcional entre 2.75-2.99),
  GMAT/GRE opcional com múltiplos critérios de dispensa, TOEFL 88 (570
  PBT)/IELTS 6.5/PTE 65, documentos exigidos (histórico não-oficial,
  resume, personal statement, cartas de recomendação se aplicável)
- Busca web: Gator Pass (MUNI ilimitado gratuito) para alunos SF State
- Busca web: SHIP internacional $2.385/ano (2025-26, oficial)
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
    "Nome da Universidade": "SF State",
    "Estado": "CA",
    "Cidade": "San Francisco",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 792,  # oficial Fall 2026: $321 base + $471 fee suplementar não-residente
    "Custo/Crédito (Residente) USD": 321,  # oficial Fall 2026
    "Tuition Total (Intl) USD": 41000,  # CORREÇÃO: planilha tinha $21.000 sem origem rastreável; estimativa oficial de custo total do programa
    "Tuition Total (Residente) USD": 31000,  # estimativa oficial de custo total do programa
    "Aluguel Privativo/mês USD": 1100,  # quarto privativo, faixa "econômica" de mercado em SF (bairros mais acessíveis, dividindo moradia)
    "Transporte/mês USD": 0,  # Gator Pass — MUNI ilimitado gratuito para alunos SF State (via fee obrigatória já embutida no tuition total)
    "Seguro Saúde (programa) USD": 2385,  # oficial 2025-26, plano internacional SHIP
    "GPA Mínimo": 3.0,  # admissão excepcional possível entre 2.75-2.99
    "GMAT/GRE Obrigatório": (
        "Opcional — dispensável com mestrado/doutorado em área "
        "quantitativa, ou GPA ≥ 3.0 + 5 anos de experiência "
        "profissional, ou graduação em Business (AACSB) com GPA ≥ 3.0; "
        "se enviado, deadline 1º mai (Fall) / 1º out (Spring)"
    ),
    "Testes Inglês Aceitos": "TOEFL mín. 88 (570 PBT) / IELTS mín. 6.5 / PTE mín. 65",
    "Documentos Necessários": (
        "Histórico acadêmico não-oficial, resume, personal statement, "
        "cartas de recomendação (se aplicável); GRE (código 4684) e "
        "TOEFL (código 4683) se aplicável"
    ),
    "Deadline Aplicação": "Fall: 1º mai; Spring: 1º out",
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
        raise ValueError("Nenhuma linha encontrada para SF State (San Francisco, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SF State, San Francisco, CA)")

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
