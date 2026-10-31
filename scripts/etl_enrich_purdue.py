"""
Data enrichment: atualiza APENAS a linha da 'Purdue' (Mitch Daniels School of
Business, West Lafayette, IN) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics and Information Management (MSBAIM),
presencial, STEM. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e
custo de vida local.

ACHADO CRITICO DA AUDITORIA: a planilha tinha 'Tuition Total (Residente) USD'
= 27.000, que é a tarifa por crédito de residente de Indiana de 2025-26
($750 x 36). A página oficial de fee rates 2026-27 da Bursar confirma reajuste
para $790/crédito -> $28.440 total. CORRIGIDO. A tarifa internacional
($1.575/crédito x 36 = $56.700) NAO mudou entre os dois anos -- confirmada,
sem alteração.

Custo de vida: 'Seguro Saúde (programa)' e 'Custo Vida Total' mantidos como
"Verificar" -- não foi possível confirmar o prêmio anual do plano de saúde
para estudantes de pós-graduação do campus West Lafayette (PGSHIP/Boiler
Insurance, via Academic HealthPlans) com uma fonte oficial acessível; sem
esse componente, a soma do custo de vida total ficaria incompleta e não deve
ser inventada.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://business.purdue.edu/masters/programs/ms-business-analytics-and-information-management/
  (duração, créditos, STEM designation)
- https://www.purdue.edu/treasurer/finance/bursar-office/tuition/fee-rates-2026-2027/professional-masters-program-fees-2026-2027/
  (tarifas por crédito 2026-27: residente $790, não-residente $1.020,
  internacional $1.575; taxa fixa internacional de $500/semestre)
- https://business.purdue.edu/masters/programs/ms-business-analytics-and-information-management/application.php
  (deadlines: Early 15/dez, Internacional 1/abr, Doméstico 1/jul; taxa de
  inscrição $75 intl; documentos exigidos; GMAT/GRE dispensável c/ GPA>=3.0;
  IELTS 7.0 e Duolingo 130 confirmados -- o TOEFL "4.5" exibido na página
  é claramente um artefato de template, não um score real, então NÃO foi
  usado)
- https://krannert.purdue.edu/.../description-of-application-materials.php
  (TOEFL 93 total, mínimos por seção W20/S22/L18/R22 -- página legada
  Krannert, mas é o único valor de TOEFL em escala plausível encontrado)
- https://business.purdue.edu/masters/programs/ms-business-analytics-and-information-management/profile-msbaim.pdf
  (class profile oficial, nov/2025: GPA médio 3.41, experiência 1-2 anos,
  42% mulheres / 58% homens)
- Página oficial de scholarships: bolsa de mérito automática p/ todos os
  candidatos, sem inscrição separada; assistantships não são tipicamente
  oferecidas pelo MSBAIM em si.
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
    "Nome da Universidade": "Purdue",
    "Estado": "IN",
    "Cidade": "West Lafayette",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1575,  # confirma Tuition Total Intl: 36 x 1575 = 56.700 (já cadastrado, correto)
    "Custo/Crédito (Residente) USD": 790,  # CORRECAO: tarifa 2026-27 (era 750, tarifa de 2025-26)
    "Tuition Total (Residente) USD": 28440,  # CORRECAO: 36 x 790 (2026-27); planilha tinha 27.000 (tarifa antiga)
    "Fees Adicionais USD": 1500,  # $500/semestre (internacional) x 3 semestres (Fall+Spring+Summer, programa de 12 meses)
    "GPA Mínimo": 3.0,  # 3.5 é o valor preferencial, não o mínimo formal
    "GPA Médio Admitidos": 3.41,  # class profile oficial (compilado de coortes recentes, nov/2025)
    "GMAT/GRE Obrigatório": (
        "Dispensável para GPA de graduação ≥ 3.0 (3.5 é o valor "
        "preferencial); recomendado enviar para reforçar candidatura e "
        "elegibilidade a bolsas; GMAT preferencial ≥ 600; GRE aceito no "
        "lugar do GMAT"
    ),
    "Experiência Profissional": (
        "Não obrigatória, mas 2 estágios relevantes ou 12+ meses de "
        "experiência profissional são incentivados; média da turma "
        "recente: 1-2 anos"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 93 total (mín. por seção: Writing 20/Speaking 22/"
        "Listening 18/Reading 22) ou IELTS 7.0 (todas as seções) ou "
        "Duolingo 130; dispensável se bacharelado/mestrado concluído nos "
        "últimos 36 meses em instituição de ensino em inglês"
    ),
    "Deadline Aplicação": (
        "Early (prioridade p/ ajuda financeira): 15/dez; Internacional "
        "(não-EUA, deadline final): 1/abr; Doméstico/EUA: 1/jul — início "
        "em agosto"
    ),
    "Taxa Aplicação USD": 75,  # internacional; isenção p/ residentes permanentes, ex-alunos Purdue, militares/veteranos
    "Rounds Aplicação": (
        "Sem rounds formais; 3 deadlines fixos: Early (15/dez), "
        "Internacional (1/abr), Doméstico (1/jul)"
    ),
    "Documentos Necessários": (
        "2 cartas de recomendação (somente online), transcripts oficiais, "
        "resume, Academic Statement of Purpose (até 500 palavras), "
        "Personal History Statement (até 500 palavras), TOEFL/IELTS/"
        "Duolingo (internacionais), GMAT/GRE (opcional, dispensável com "
        "GPA ≥ 3.0)"
    ),
    "Entrevista": (
        "Apenas sob convite (não é padrão para todos os candidatos); "
        "candidatos muito bem qualificados podem ter essa exigência "
        "dispensada"
    ),
    "TA/RA Disponível": (
        "Não oferecidas diretamente pelo MSBAIM; vagas de assistente de "
        "pós-graduação podem existir em outros departamentos da "
        "universidade, mas exigem busca ativa"
    ),
    "Processo TA/RA": (
        "Sem processo formal do programa; opção de emprego part-time no "
        "Krenicki Center for Business Analytics & Machine Learning, ou "
        "busca de vagas de graduate assistant fora do MSBAIM"
    ),
    "Graduate Assistantships": (
        "Não oferecidas diretamente pelo MSBAIM (programa full-time de 12 "
        "meses); todos os candidatos são automaticamente considerados "
        "para bolsas de mérito no processo de admissão, sem inscrição "
        "separada"
    ),
    "Tipos de Bolsa": (
        "Merit (automática para os melhores candidatos, sem inscrição "
        "separada); resultados liberados a admitidos a cada ~2 semanas"
    ),  # refinamento: planilha tinha só "Merit (top applicants)"
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
        raise ValueError("Nenhuma linha encontrada para Purdue (West Lafayette, IN). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Purdue, West Lafayette, IN)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
