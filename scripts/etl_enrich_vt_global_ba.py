"""
Data enrichment: atualiza APENAS a linha da 'VT Global BA' (Pamplin
College of Business, cadastrada originalmente em "Northern Virginia", VA)
no msba_programs_database.xlsx.

Foco exclusivo: MSBA com concentração em Global Business Analytics
(MSBA-GBA), presencial, 30 créditos (10 disciplinas). Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida.

ACHADO CRITICO DA AUDITORIA: o programa MUDOU DE CAMPUS. Lançado em Fall
2024 no Virginia Tech Northern Virginia Center (Falls Church), passou
para o Innovation Campus (Alexandria/Potomac Yard) na Spring 2025, e a
partir de Fall 2025 (ano letivo corrente) o programa passou a ser
sediado em BLACKSBURG -- o mesmo campus do outro programa de Virginia
Tech já cadastrado nesta planilha (MSBA-BA, linha "Virginia Tech"/
Blacksburg). Isso muda completamente a proposta de valor do programa
(deixou de estar na região metropolitana de DC, que era seu diferencial
de rede/estágios). CORRIGIDO: 'Cidade' de "Northern Virginia" para
"Blacksburg". Múltiplas fontes independentes confirmam a mudança
(página oficial "About Our Program", busca web cruzada).

TUITION: CORRIGIDO. A planilha tinha $1.100/crédito ($33.000 total,
flat). A página oficial do programa (masters.pamplin.vt.edu) e buscas
cruzadas confirmam a taxa oficial atual de $1.000/crédito, igual para
residentes e não-residentes (flat tuition) -- o valor de $1.100 parece
ter vindo de uma fonte terceira (IDP) divergente da oficial. CORRIGIDO
para $1.000/crédito = $30.000 total (30 créditos), mantendo Intl =
Residente (programa flat).

'STEM Designated' estava 'Sim', mas NENHUMA fonte oficial encontrada
confirma ou nega explicitamente a designação STEM para esta concentração
especifica (GBA). É um sinal de alerta relevante que o programa-irmão
mais proximo (MSBA-BA, mesmo "guarda-chuva" MSBA / mesmo tipo de
concentração dentro do mesmo diploma) NÃO é STEM-designado (CIP
52.0201, confirmado em enrichment anterior desta mesma planilha) --
mas como MSBA-GBA pode ter CIP proprio (é o programa do parceria 3+1+1
com o NMIMS, Mumbai, com estrutura curricular distinta), não assumimos
o mesmo status sem confirmação direta. Diante da incerteza e do risco
de um dado "Sim" ja se mostrar errado (mesmo padrão do MSBA-BA antes de
corrigido), CORRIGIDO de 'Sim' para 'Verificar' até uma fonte oficial
explicita ser encontrada -- especialmente relevante pois STEM é
requisito obrigatorio do perfil de busca da usuaria.

Custo de vida: como o programa agora funciona fisicamente em Blacksburg,
reaproveitados 2 fatos ja verificados de fonte oficial no enrichment do
MSBA-BA (mesma cidade/campus, mesmas politicas universitárias, não
específicas de mercado de aluguel): 'Transporte/mês' = $0 (Blacksburg
Transit gratuito para alunos de Virginia Tech) e 'Seguro Saúde
(programa)' = $3.993 (taxa padrão não-subsidiada do VT Student Medical
Insurance Plan 2026-27). Aluguel/Mercado/Utilities NÃO foram copiados do
MSBA-BA porque são estimativas de mercado especificas daquele
levantamento, não fatos institucionais -- permanecem 'Verificar' até
pesquisa própria.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://masters.pamplin.vt.edu/globalbusinessanalytics/about-our-program.html
  (relocação de campus, formato, duração, créditos, curriculo, pathway
  3+1+1 com NMIMS)
- Busca web cruzada confirmando tuition oficial $1.000/crédito flat
  (vs. $1.100 de fonte terceira/IDP)
- Busca web: GPA mínimo 3.0, GMAT/GRE não exigido, TOEFL exigido para
  não-nativos exceto graduados em instituição de instrução em inglês,
  personal statement de 1-2 páginas + vídeo + experiência profissional
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
    "Nome da Universidade": "VT Global BA",
    "Estado": "VA",
    "Cidade": "Northern Virginia",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Cidade": "Blacksburg",  # CORRECAO CRITICA: programa relocado de Northern Virginia para Blacksburg a partir de Fall 2025
    "Nome Oficial do Programa": "Masters of Science in Business Administration with a concentration in Global Business Analytics (MSBA-GBA)",
    "STEM Designated": VERIFICAR,  # CORRECAO: "Sim" não confirmado por nenhuma fonte oficial direta; programa-irmão (MSBA-BA) não é STEM, risco de erro semelhante
    "Duração (meses)": 12,  # 2 semestres, formato tradicional full-time
    "Destaques Currículo": (
        "5 cursos core de MBA (contabilidade, finanças, comportamento "
        "organizacional, estatística, marketing) + 5 cursos de "
        "analytics/BI/visualização + capstone; 1 eletiva de "
        "especialização (gestão internacional, marketing, finanças ou "
        "liderança em tecnologia); pathway 3+1+1 com o NMIMS (Mumbai, "
        "Índia)"
    ),
    "Custo/Crédito (Intl) USD": 1000,  # CORREÇÃO: oficial $1.000/crédito flat (planilha tinha $1.100, de fonte terceira)
    "Custo/Crédito (Residente) USD": 1000,  # CORREÇÃO: oficial $1.000/crédito flat
    "Tuition Total (Intl) USD": 30000,  # CORREÇÃO: 30 créditos x $1.000
    "Tuition Total (Residente) USD": 30000,  # CORREÇÃO: 30 créditos x $1.000
    "Transporte/mês USD": 0,  # Blacksburg Transit gratuito para alunos de Virginia Tech (mesmo campus do MSBA-BA)
    "Seguro Saúde (programa) USD": 3993,  # VT Student Medical Insurance 2026-27, taxa padrão não-subsidiada (mesma política institucional do MSBA-BA)
    "GPA Mínimo": 3.0,
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal exigido para admissão",
    "Documentos Necessários": (
        "Personal statement (1–2 páginas, espaçamento duplo, objetivos "
        "de carreira), componente de vídeo, informações de experiência "
        "profissional, TOEFL/IELTS (internacionais, salvo graduação em "
        "instituição de instrução em inglês)"
    ),
    "Rounds Aplicação": "Admissão contínua (rolling); aplicação antecipada é recomendada",
    "Entrevista": VERIFICAR,
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
        raise ValueError("Nenhuma linha encontrada para VT Global BA (Northern Virginia, VA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (VT Global BA, originalmente Northern Virginia, VA)")

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
