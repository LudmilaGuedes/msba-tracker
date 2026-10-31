"""
ETL: adiciona novos programas de MSBA da Florida (FL) ao msba_programs_database.xlsx.

Extracao: pesquisa web (feita manualmente pelo assistente, fora deste script).
Validacao de duplicatas: programas cujas universidades ja constam na aba
"MSBA Programs" foram excluidos manualmente antes de popular NEW_PROGRAMS abaixo.
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

# Novos programas (headers identicos aos da planilha original, na ordem do schema de 54 colunas).
NEW_PROGRAMS = [
    {
        "Nome da Universidade": "University of North Florida (UNF)",
        "Estado": "FL",
        "Cidade": "Jacksonville",
        "Business School": "Coggin College of Business",
        "Tipo (Pub/Priv)": "Pública",
        "Nome Oficial do Programa": "Master of Science in Business Analytics (MSBA)",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": 12,
        "Início": "Fall, Spring, Summer",
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Concentrações em Sports Analytics, FinTech, Marketing Analytics, "
            "Econometrics e Logistics/Supply Chain Management; projetos aplicados "
            "com empresas parceiras (FIS, Regency Centers, Crowley); AACSB accredited"
        ),
        "Total Créditos": 30,
        "Custo/Crédito (Intl) USD": 1126.88,
        "Custo/Crédito (Residente) USD": 493.53,
        "Tuition Total (Intl) USD": 33806,
        "Tuition Total (Residente) USD": 14806,
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
        "GPA Mínimo": 3.0,
        "GPA Médio Admitidos": VERIFICAR,
        "GMAT/GRE Obrigatório": "Dispensa condicional (STEM/AACSB, 2 anos exp. ou pós-graduação)",
        "GMAT Médio": VERIFICAR,
        "Taxa Aceitação %": VERIFICAR,
        "Experiência Profissional": VERIFICAR,
        "Matérias Pré-Requisito": VERIFICAR,
        "Testes Inglês Aceitos": VERIFICAR,
        "Deadline Aplicação": "Fall: 1 ago (Intl: 1 jun); Spring: 1 dez (Intl: 1 out); Summer: 15 abr",
        "Taxa Aplicação USD": VERIFICAR,
        "Rounds Aplicação": VERIFICAR,
        "Documentos Necessários": VERIFICAR,
        "Entrevista": VERIFICAR,
        "Dificuldade Entrada": VERIFICAR,
    },
    {
        "Nome da Universidade": "Florida Institute of Technology (Florida Tech)",
        "Estado": "FL",
        "Cidade": "Melbourne",
        "Business School": "Nathan M. Bisk College of Business",
        "Tipo (Pub/Priv)": "Privada",
        "Nome Oficial do Programa": "Master of Science in Business Analytics",
        "STEM Designated": "Sim",
        "Formato": "Presencial",
        "Duração (meses)": VERIFICAR,
        "Início": VERIFICAR,
        "Tamanho Turma": VERIFICAR,
        "Destaques Currículo": (
            "Machine learning e deep learning aplicados a segmentação de clientes, "
            "forecasting de demanda e detecção de fraude/anomalias; capstone em "
            "grupo com organização real; oferecido no campus de Melbourne e online"
        ),
        "Total Créditos": 30,
        "Custo/Crédito (Intl) USD": 1310,
        "Custo/Crédito (Residente) USD": 1310,
        "Tuition Total (Intl) USD": 39300,
        "Tuition Total (Residente) USD": 39300,
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
        "Matérias Pré-Requisito": "Cursos de nivelamento podem ser exigidos conforme histórico prévio",
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
            row[1].value = f"{row[1].value} + {len(NEW_PROGRAMS)} programas FL em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"Adicionadas {len(added_names)} linhas em '{SHEET_DATA}':")
    for name in added_names:
        print(f"  - {name}")


if __name__ == "__main__":
    main()
