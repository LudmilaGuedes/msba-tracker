"""
Data enrichment: atualiza APENAS a linha da 'UConn' (School of
Business, campus Hartford/Stamford) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: o domínio business.uconn.edu retornou erro de
certificado SSL expirado em todas as tentativas via WebFetch, tornando
a página específica do programa inacessível por esse caminho. Contornado
navegando por catalog.uconn.edu (catálogo acadêmico oficial) e
grad.uconn.edu (Graduate School), e finalmente obtendo a planilha
oficial de tuition 2026-27 diretamente do domínio do Bursar
(bursar.media.uconn.edu), que PÔDE ser lida integralmente via
openpyxl após o download.

Foco: Master of Science in Business Analytics and Project Management
(MSBAPM), formato presencial (o programa também tem versões online e
híbrida, não tocadas aqui), 37 créditos (já confirmado oficialmente
pelo catálogo), full-time ou part-time.

TUITION: CONFIRMADO E COMPLETADO via planilha oficial "2026-2027
Graduate Program Rates" do Bursar da UConn. A linha exata do programa
("Business Analytics and Project Management MS") mostra: Cost per
Credit In-State = $1.250 / Out-of-State = $1.250 -- tarifa FLAT, sem
diferenciação de residência (confirma e valida o valor Intl já
presente na planilha, $1.250/crédito, exatamente idêntico).
PREENCHIDO 'Custo/Crédito (Residente)' com o mesmo valor ($1.250) e
'Tuition Total (Residente)' com 37 × $1.250 = $46.250 (idêntico ao
total Intl já presente, confirmando consistência).

'Fees Adicionais USD' preenchido com $0 -- a mesma planilha oficial
lista "Part Time Fees" e "Full Time Fees" como $0 para este programa
específico (diferente de outros programas na mesma planilha, que têm
fees > 0), indicando que é um programa "fee-based"/self-support com
fees já embutidas no valor por crédito, sem cobrança adicional
separada.

'Caminho Tuition Residente' preenchido: N/A -- tuition é flat/
self-support, sem premium de não-residência, confirmado oficialmente.

'TA/RA Disponível' e 'Graduate Assistantships' preenchidos com
CONFIRMAÇÃO NEGATIVA OFICIAL: a mesma planilha rotula explicitamente
este programa como "Not Eligible for GA Waiver" na coluna de Grad
Assistant Fees -- ou seja, diferente de outros programas de mestrado
da UConn, o MSBAPM NÃO oferece a isenção de tuition padrão para
graduate assistants. Não é uma lacuna de pesquisa, é uma negativa
confirmada pela própria fonte oficial de tuition.

'GPA Mínimo' preenchido com o critério geral (tiered) da Graduate
School da UConn, aplicável a todos os programas de mestrado salvo
indicação contrária do programa específico: GPA cumulativo ≥3.0 em
qualquer diploma de bacharelado ou superior; OU ≥3.0 nos 2 anos mais
recentes de graduação full-time; OU ≥3.5 no ano mais recente de
graduação full-time; OU ≥3.0 em pelo menos 1 semestre (9+ créditos) de
coursework pós-graduação.

'Testes Inglês Aceitos' refinado com os mínimos oficiais gerais da
Graduate School: TOEFL iBT (nova escala) mín. 4.5; IELTS mín. 6.5; PTE
Academic mín. 60; Duolingo mín. 110; validade de 2 anos.

'Taxa Aplicação USD' preenchido: $75 (dispensável em certas
circunstâncias). 'Documentos Necessários' preenchido com a lista
oficial geral: diploma de bacharelado de instituição credenciada;
transcripts oficiais de todas as instituições pós-secundárias;
avaliação de credencial por terceiros para diplomas não-americanos
(salvo o programa dispense); para internacionais: transcripts e
diplomas em inglês E no idioma original.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Duração (meses)' (o programa aceita entrada em Fall/Spring/Summer e
tem ritmo full-time OU part-time flexível, sem uma duração única
confirmada), 'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio'
(GMAT/GRE já estava marcado como 'Opcional', não Verificar, não
tocado), 'Taxa Aceitação %', 'Ranking US News', 'Ranking QS', 'Taxa
Empregabilidade %', 'Salário Médio Pós USD', 'Principais
Empregadores', 'Matérias Pré-Requisito' (catálogo confirma
explicitamente que não há pré-requisitos formais), 'Deadline
Aplicação'/'Rounds Aplicação' (a Graduate School geral só menciona que
"alguns programas têm deadlines mais cedo", sem especificar o do
MSBAPM), 'Entrevista', 'Seguro Saúde (programa) USD', e todos os
campos mensais de custo de vida em Hartford/Stamford, CT.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://catalog.uconn.edu/graduate/degree-programs/business-analytics-project-management-ms/
  (formato, 37 créditos confirmados, currículo, ausência de
  pré-requisitos formais)
- https://grad.uconn.edu/admissions/requirements/ (GPA tiered geral,
  mínimos de inglês, taxa de aplicação $75, documentos)
- https://bursar.media.uconn.edu/wp-content/uploads/sites/3730/2026/07/2026-2027-Graduate-Program-Rates.xlsx
  (planilha oficial do Bursar, 2026-27 — linha específica do MSBAPM:
  tuition flat $1.250/crédito, fees $0, não elegível para isenção de
  GA)
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
    "Nome da Universidade": "UConn",
    "Estado": "CT",
    "Cidade": "Hartford/Stamford",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 1250,  # oficial 2026-27, tarifa flat (planilha do Bursar)
    "Tuition Total (Residente) USD": 46250,  # 37 x $1.250, idêntico ao total Intl já presente
    "Fees Adicionais USD": 0,  # oficial: "Part Time Fees" e "Full Time Fees" = $0 para este programa (fee-based/self-support)
    "Caminho Tuition Residente": "N/A — tuition flat/self-support, mesma tarifa residente e não-residente (confirmado oficialmente)",
    "TA/RA Disponível": "Não disponível — confirmado oficialmente: programa marcado como \"Not Eligible for GA Waiver\" na planilha de tuition do Bursar",
    "Graduate Assistantships": "Não elegível para isenção de tuition de GA (confirmado oficialmente, diferente de outros mestrados da UConn)",
    "GPA Mínimo": (
        "≥3.0 cumulativo em qualquer diploma de bacharelado ou "
        "superior; OU ≥3.0 nos 2 anos mais recentes de graduação "
        "full-time; OU ≥3.5 no ano mais recente; OU ≥3.0 em 1+ "
        "semestre de coursework pós-graduação (critério geral da "
        "Graduate School)"
    ),
    "Testes Inglês Aceitos": "TOEFL iBT (nova escala) mín. 4.5; IELTS mín. 6.5; PTE Academic mín. 60; Duolingo mín. 110; validade de 2 anos",
    "Taxa Aplicação USD": 75,  # dispensável em certas circunstâncias
    "Documentos Necessários": (
        "Diploma de bacharelado de instituição credenciada; "
        "transcripts oficiais de todas as instituições; avaliação de "
        "credencial por terceiros para diplomas não-americanos (salvo "
        "dispensa do programa); internacionais: transcripts e "
        "diplomas em inglês E no idioma original"
    ),
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal (confirmado oficialmente pelo catálogo)",
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
        raise ValueError("Nenhuma linha encontrada para UConn (Hartford/Stamford, CT). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UConn, Hartford/Stamford, CT)")

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
