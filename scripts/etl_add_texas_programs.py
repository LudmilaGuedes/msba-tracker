"""
ETL: adiciona novos programas de MSBA do Texas (TX) ao msba_programs_database.xlsx.

Extracao: pesquisa web (feita manualmente pelo assistente, fora deste script).
Validacao de duplicatas: as 12 universidades do Texas ja presentes na aba
"MSBA Programs" (UT Austin, UT Dallas, Texas A&M, Baylor, SMU, TCU, UTSA,
UT Arlington, Texas Tech, U of Houston, Texas State, UNT) foram conferidas
e excluidas manualmente antes de popular NEW_PROGRAMS abaixo. Rice University
foi pesquisada e descartada: nao possui MS em Business Analytics dedicado
(apenas MBA/MAcc/PhD).
Transformacao: cada registro segue estritamente as 54 colunas do schema, na
mesma ordem dos cabecalhos ja existentes na planilha. Campos nao confirmados
pela pesquisa == "Verificar".
Carga: apenda as novas linhas na aba "MSBA Programs", replicando o estilo
visual (fonte, borda, fill amarelo em celulas "Verificar") das linhas
existentes, atualiza o auto-filtro e os totais na aba "Resumo".
"""

import shutil
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"

VERIFICAR = "Verificar"

NEW_PROGRAMS = [
    {
        "Nome da Universidade": "East Texas A&M University (ETAMU)",
        "Estado": "TX",
        "Cidade": "Commerce",
        "Business School": "College of Business",
        "Tipo (Pub/Priv)": "Pública",
        "Nome Oficial do Programa": "Master of Science in Business Analytics",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": 18,
        "Início": "Fall, Spring, Summer",
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Concentrações em General Business Analytics ou Applied Analytics "
            "(12h cada); exame abrangente ao final do curso; Python, SQL, "
            "modelagem estatística"
        ),
        "Total Créditos": 30,
        "Custo/Crédito (Intl) USD": 855.89,
        "Custo/Crédito (Residente) USD": 427.89,
        "Tuition Total (Intl) USD": 25677,
        "Tuition Total (Residente) USD": 12837,
        "Fees Adicionais USD": VERIFICAR,
        "Aluguel Privativo/mês USD": VERIFICAR,
        "Mercado/mês USD": VERIFICAR,
        "Seguro Saúde (programa) USD": VERIFICAR,
        "Seguro Residencial/mês USD": VERIFICAR,
        "Utilities/mês USD": VERIFICAR,
        "Celular/mês USD": VERIFICAR,
        "WiFi/mês USD": VERIFICAR,
        "Transporte/mês USD": VERIFICAR,
        "Outros Gastos/mês USD": VERIFICAR,
        "Custo Vida Total (programa) USD": VERIFICAR,
        "Bolsas Disponíveis": VERIFICAR,
        "Tipos de Bolsa": VERIFICAR,
        "Quão Realista (Intl)": VERIFICAR,
        "TA/RA Disponível": VERIFICAR,
        "Processo TA/RA": VERIFICAR,
        "Caminho Tuition Residente": VERIFICAR,
        "Graduate Assistantships": VERIFICAR,
        "Ranking US News": VERIFICAR,
        "Ranking QS": VERIFICAR,
        "Taxa Empregabilidade %": VERIFICAR,
        "Salário Médio Pós USD": VERIFICAR,
        "Principais Empregadores": VERIFICAR,
        "GPA Mínimo": VERIFICAR,
        "GPA Médio Admitidos": VERIFICAR,
        "GMAT/GRE Obrigatório": "Dispensa condicional (GRE obrigatório apenas se GPA ≤ 2.5)",
        "GMAT Médio": VERIFICAR,
        "Taxa Aceitação %": VERIFICAR,
        "Experiência Profissional": VERIFICAR,
        "Matérias Pré-Requisito": "Nenhum pré-requisito obrigatório",
        "Testes Inglês Aceitos": VERIFICAR,
        "Deadline Aplicação": "Intl - Fall: 15 mai; Spring: 15 out; Summer: 15 mar",
        "Taxa Aplicação USD": VERIFICAR,
        "Rounds Aplicação": VERIFICAR,
        "Documentos Necessários": "Personal statement (até 500 palavras), resume/CV, histórico acadêmico",
        "Entrevista": VERIFICAR,
        "Dificuldade Entrada": VERIFICAR,
    },
    {
        "Nome da Universidade": "Texas A&M University-Texarkana (A&M-Texarkana)",
        "Estado": "TX",
        "Cidade": "Texarkana",
        "Business School": "College of Business, Engineering and Technology",
        "Tipo (Pub/Priv)": "Pública",
        "Nome Oficial do Programa": "Master of Science in Business and Data Analytics",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": 15,
        "Início": "Fall",
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Três pilares de analytics (descritivo, preditivo, prescritivo); "
            "Python, SQL, Tableau; programa lançado no Fall 2026"
        ),
        "Total Créditos": 30,
        "Custo/Crédito (Intl) USD": VERIFICAR,
        "Custo/Crédito (Residente) USD": VERIFICAR,
        "Tuition Total (Intl) USD": VERIFICAR,
        "Tuition Total (Residente) USD": VERIFICAR,
        "Fees Adicionais USD": VERIFICAR,
        "Aluguel Privativo/mês USD": VERIFICAR,
        "Mercado/mês USD": VERIFICAR,
        "Seguro Saúde (programa) USD": VERIFICAR,
        "Seguro Residencial/mês USD": VERIFICAR,
        "Utilities/mês USD": VERIFICAR,
        "Celular/mês USD": VERIFICAR,
        "WiFi/mês USD": VERIFICAR,
        "Transporte/mês USD": VERIFICAR,
        "Outros Gastos/mês USD": VERIFICAR,
        "Custo Vida Total (programa) USD": VERIFICAR,
        "Bolsas Disponíveis": VERIFICAR,
        "Tipos de Bolsa": VERIFICAR,
        "Quão Realista (Intl)": VERIFICAR,
        "TA/RA Disponível": VERIFICAR,
        "Processo TA/RA": VERIFICAR,
        "Caminho Tuition Residente": VERIFICAR,
        "Graduate Assistantships": VERIFICAR,
        "Ranking US News": VERIFICAR,
        "Ranking QS": VERIFICAR,
        "Taxa Empregabilidade %": VERIFICAR,
        "Salário Médio Pós USD": VERIFICAR,
        "Principais Empregadores": VERIFICAR,
        "GPA Mínimo": VERIFICAR,
        "GPA Médio Admitidos": VERIFICAR,
        "GMAT/GRE Obrigatório": VERIFICAR,
        "GMAT Médio": VERIFICAR,
        "Taxa Aceitação %": VERIFICAR,
        "Experiência Profissional": VERIFICAR,
        "Matérias Pré-Requisito": VERIFICAR,
        "Testes Inglês Aceitos": VERIFICAR,
        "Deadline Aplicação": VERIFICAR,
        "Taxa Aplicação USD": VERIFICAR,
        "Rounds Aplicação": VERIFICAR,
        "Documentos Necessários": VERIFICAR,
        "Entrevista": VERIFICAR,
        "Dificuldade Entrada": VERIFICAR,
    },
    {
        "Nome da Universidade": "University of Dallas",
        "Estado": "TX",
        "Cidade": "Irving",
        "Business School": "Satish & Yasmin Gupta College of Business",
        "Tipo (Pub/Priv)": "Privada",
        "Nome Oficial do Programa": "Master of Science in Business Analytics",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": VERIFICAR,
        "Início": VERIFICAR,
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Liderança e ética integradas ao currículo técnico; visualização "
            "de dados, SQL, cloud computing, aplicações de IA"
        ),
        "Total Créditos": 30,
        "Custo/Crédito (Intl) USD": 1300,
        "Custo/Crédito (Residente) USD": 1300,
        "Tuition Total (Intl) USD": 39000,
        "Tuition Total (Residente) USD": 39000,
        "Fees Adicionais USD": 385,
        "Aluguel Privativo/mês USD": VERIFICAR,
        "Mercado/mês USD": VERIFICAR,
        "Seguro Saúde (programa) USD": VERIFICAR,
        "Seguro Residencial/mês USD": VERIFICAR,
        "Utilities/mês USD": VERIFICAR,
        "Celular/mês USD": VERIFICAR,
        "WiFi/mês USD": VERIFICAR,
        "Transporte/mês USD": VERIFICAR,
        "Outros Gastos/mês USD": VERIFICAR,
        "Custo Vida Total (programa) USD": VERIFICAR,
        "Bolsas Disponíveis": VERIFICAR,
        "Tipos de Bolsa": VERIFICAR,
        "Quão Realista (Intl)": VERIFICAR,
        "TA/RA Disponível": VERIFICAR,
        "Processo TA/RA": VERIFICAR,
        "Caminho Tuition Residente": VERIFICAR,
        "Graduate Assistantships": VERIFICAR,
        "Ranking US News": VERIFICAR,
        "Ranking QS": VERIFICAR,
        "Taxa Empregabilidade %": VERIFICAR,
        "Salário Médio Pós USD": VERIFICAR,
        "Principais Empregadores": VERIFICAR,
        "GPA Mínimo": VERIFICAR,
        "GPA Médio Admitidos": VERIFICAR,
        "GMAT/GRE Obrigatório": "Não obrigatório",
        "GMAT Médio": VERIFICAR,
        "Taxa Aceitação %": VERIFICAR,
        "Experiência Profissional": VERIFICAR,
        "Matérias Pré-Requisito": VERIFICAR,
        "Testes Inglês Aceitos": VERIFICAR,
        "Deadline Aplicação": VERIFICAR,
        "Taxa Aplicação USD": VERIFICAR,
        "Rounds Aplicação": VERIFICAR,
        "Documentos Necessários": VERIFICAR,
        "Entrevista": VERIFICAR,
        "Dificuldade Entrada": VERIFICAR,
    },
    {
        "Nome da Universidade": "West Texas A&M University (WTAMU)",
        "Estado": "TX",
        "Cidade": "Canyon",
        "Business School": "Paul and Virginia Engler College of Business",
        "Tipo (Pub/Priv)": "Pública",
        "Nome Oficial do Programa": "Master of Science in Computer Information Systems and Business Analytics (MSCISBA)",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": VERIFICAR,
        "Início": "Fall, Spring, Summer",
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Seis áreas de competência (análise quantitativa, data analytics, "
            "data mining/management, desenvolvimento de sistemas, "
            "redes/cibersegurança, integração de negócios); capstone final (CIDM 6395)"
        ),
        "Total Créditos": 33,
        "Custo/Crédito (Intl) USD": 500,
        "Custo/Crédito (Residente) USD": 500,
        "Tuition Total (Intl) USD": 16500,
        "Tuition Total (Residente) USD": 16500,
        "Fees Adicionais USD": VERIFICAR,
        "Aluguel Privativo/mês USD": VERIFICAR,
        "Mercado/mês USD": VERIFICAR,
        "Seguro Saúde (programa) USD": VERIFICAR,
        "Seguro Residencial/mês USD": VERIFICAR,
        "Utilities/mês USD": VERIFICAR,
        "Celular/mês USD": VERIFICAR,
        "WiFi/mês USD": VERIFICAR,
        "Transporte/mês USD": VERIFICAR,
        "Outros Gastos/mês USD": VERIFICAR,
        "Custo Vida Total (programa) USD": VERIFICAR,
        "Bolsas Disponíveis": VERIFICAR,
        "Tipos de Bolsa": VERIFICAR,
        "Quão Realista (Intl)": VERIFICAR,
        "TA/RA Disponível": VERIFICAR,
        "Processo TA/RA": VERIFICAR,
        "Caminho Tuition Residente": VERIFICAR,
        "Graduate Assistantships": VERIFICAR,
        "Ranking US News": VERIFICAR,
        "Ranking QS": VERIFICAR,
        "Taxa Empregabilidade %": VERIFICAR,
        "Salário Médio Pós USD": VERIFICAR,
        "Principais Empregadores": VERIFICAR,
        "GPA Mínimo": VERIFICAR,
        "GPA Médio Admitidos": VERIFICAR,
        "GMAT/GRE Obrigatório": "Dispensa disponível",
        "GMAT Médio": VERIFICAR,
        "Taxa Aceitação %": VERIFICAR,
        "Experiência Profissional": VERIFICAR,
        "Matérias Pré-Requisito": "Proficiência em estatística (curso prévio, exame ou CIDM 6300)",
        "Testes Inglês Aceitos": VERIFICAR,
        "Deadline Aplicação": VERIFICAR,
        "Taxa Aplicação USD": VERIFICAR,
        "Rounds Aplicação": VERIFICAR,
        "Documentos Necessários": VERIFICAR,
        "Entrevista": VERIFICAR,
        "Dificuldade Entrada": VERIFICAR,
    },
]


def main():
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"data/msba_programs_database.backup-{timestamp}.xlsx"
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    missing = [h for h in headers if h not in NEW_PROGRAMS[0]]
    if missing:
        raise ValueError(f"Colunas do schema nao preenchidas nos novos registros: {missing}")

    ref_cell = ws.cell(row=2, column=1)
    base_font = Font(name=ref_cell.font.name, sz=ref_cell.font.sz)
    base_border = Border(
        left=Side(style="thin", color="00D0D0D0"),
        right=Side(style="thin", color="00D0D0D0"),
        top=Side(style="thin", color="00D0D0D0"),
        bottom=Side(style="thin", color="00D0D0D0"),
    )
    base_alignment = Alignment(vertical="top", wrap_text=True)
    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")

    added_names = []
    next_row = ws.max_row + 1
    for program in NEW_PROGRAMS:
        for header, col_idx in header_to_col.items():
            value = program[header]
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.font = base_font
            cell.border = base_border
            cell.alignment = base_alignment
            if value == VERIFICAR:
                cell.fill = verificar_fill
        added_names.append(program["Nome da Universidade"])
        next_row += 1

    last_col_letter = ws.cell(row=1, column=len(headers)).column_letter
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    summary = wb[SHEET_SUMMARY]
    for row in summary.iter_rows(min_row=2):
        label = row[0].value
        if label in ("Total de Programas", "Programas Presenciais STEM"):
            row[1].value = (row[1].value or 0) + len(NEW_PROGRAMS)
        if label == "Data de Compilação":
            row[1].value = f"{row[1].value} + {len(NEW_PROGRAMS)} programas TX em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"Adicionadas {len(added_names)} linhas em '{SHEET_DATA}':")
    for name in added_names:
        print(f"  - {name}")


if __name__ == "__main__":
    main()
