"""
Data enrichment: atualiza APENAS a linha da 'Columbia' (IEOR/Columbia
Business School, New York, NY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), programa conjunto do
departamento de IEOR (Fu Foundation School of Engineering/SEAS) com a
Columbia Business School, presencial, STEM (CIP 52.1302), 36 créditos,
12 meses. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Nova York.

TUITION (já estava correto, apenas complementado): $93.024 para 36
créditos = $2.584/crédito. Como é universidade PRIVADA, PREENCHIDOS
'Custo/Crédito (Residente)' e 'Tuition Total (Residente)' (estavam
'Verificar') com os MESMOS valores do Intl.

ACHADO SOBRE GMAT/GRE: a planilha tinha 'Aceito' (vago). A fonte mais
direta (FAQ oficial do IEOR, via busca) indica que o GRE/GMAT NÃO É
EXIGIDO para o ciclo de admissão 2026 (Spring/Fall) -- mudança recente
em relação à política histórica, que era obrigatória. Outra fonte
(agregador) cita uma mediana de GMAT de 730 entre quem envia pontuação
voluntariamente. CORRIGIDO/PREENCHIDO refletindo essa política atual
com a ressalva histórica.

'Seguro Saúde (programa)' permanece 'Verificar': o único valor
encontrado ($3.676/ano) é do ciclo 2020-2021 -- a própria fonte que o
cita recomenda buscar o valor atual por estar desatualizado. Mesmo
critério de cautela já aplicado para Fordham (outra universidade de
NYC nesta planilha): preferir 'Verificar' a um número provavelmente
muito abaixo do custo real hoje, dado o intervalo de tempo e a
volatilidade de custos de saúde em NYC.

'GPA Médio Admitidos' (3.5) vem de um agregador terceiro (nível de
confiança moderado, não confirmado em página oficial direta -- as
tentativas de fetch direto em ieor.columbia.edu retornaram HTTP 403,
bloqueio de bot).

'TA/RA Disponível' e 'Graduate Assistantships' preenchidos com achado
relevante: ao contrário da maioria das universidades públicas desta
planilha, assistantships na Columbia NÃO cobrem tuition -- um
diferencial importante para o perfil de custo real do programa.

Fontes principais (pesquisadas via busca web, ago/2026):
- Busca web cruzada (ieor.columbia.edu FAQ, mim-essay.com, yocket,
  collegedunia): tuition $93.024, GRE/GMAT não exigido no ciclo 2026,
  vídeo interview obrigatório (3 perguntas, 90s de reflexão + 3min de
  resposta cada), taxa de aplicação $85, deadline final 15 fev 2026,
  bolsas parciais de mérito com consideração automática, fellowships
  nomeadas via IEOR (variam ano a ano), assistantships sem cobertura
  de tuition
- Busca web: seguro saúde $3.676/ano (2020-21, desatualizado, não
  usado)

Páginas oficiais ieor.columbia.edu retornaram HTTP 403 (bloqueio de
bot) em fetch direto — dados complementados via busca web e fontes
terceiras, sem contradição relevante entre si (exceto onde notado).
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
    "Nome da Universidade": "Columbia",
    "Estado": "NY",
    "Cidade": "New York",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2584,  # $93.024 / 36 créditos
    "Custo/Crédito (Residente) USD": 2584,  # privada, mesma taxa do Intl
    "Tuition Total (Residente) USD": 93024,  # privada, mesma taxa do Intl
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Bolsas parciais de mérito concedidas a candidatos fortes "
        "(consideração automática na aplicação, sem formulário "
        "extra); Fellowships nomeadas ocasionais via IEOR, dependentes "
        "de financiamento de alumni/corporativo, variam ano a ano"
    ),
    "TA/RA Disponível": "Limitado — assistantships existem mas NÃO cobrem tuition (diferente da maioria das universidades públicas)",
    "Graduate Assistantships": "Existem, mas não cobrem tuition (diferencial importante em relação a programas públicos)",
    "GPA Médio Admitidos": 3.5,  # fonte agregadora, confiança moderada (não confirmado em página oficial direta)
    "GMAT/GRE Obrigatório": (
        "Não exigido para o ciclo de admissão 2026 (mudança recente; "
        "pode ser enviado voluntariamente); historicamente era "
        "obrigatório"
    ),
    "GMAT Médio": 730,  # mediana entre candidatos que enviam pontuação (fonte agregadora)
    "Taxa Aplicação USD": 85,
    "Documentos Necessários": (
        "Histórico acadêmico, resume, GMAT/GRE (opcional no ciclo "
        "2026, self-reported), vídeo interview obrigatório (3 "
        "perguntas aleatórias, 90s para pensar + 3min para responder "
        "cada)"
    ),
    "Entrevista": "Sim — vídeo interview obrigatório, parte da aplicação (não é com pessoa ao vivo)",
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
        raise ValueError("Nenhuma linha encontrada para Columbia (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Columbia, New York, NY)")

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
