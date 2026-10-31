"""
Data enrichment: atualiza APENAS a linha da 'U of Rochester' (Simon Business
School, Rochester, NY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (full-time, presencial, STEM). Deep
research em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida local.

ACHADO CRITICO DA AUDITORIA: 'Tuition Total (Intl) USD' estava em 72.000.
A página oficial de tuition da Simon (tabela por programa, datada de
14/nov/2025, ano letivo 2026-27) lista MS in Business Analytics em
$68.000/ano -- CORRIGIDO. Também: 'Taxa Aplicação USD' estava 0; a página
oficial de admissions lista $90 -- CORRIGIDO.

Custo de vida: os 8 componentes mensais (aluguel, mercado, seguro
residencial, utilities, celular, wifi, transporte, outros) já estavam
preenchidos na planilha e foram MANTIDOS como estão -- a página oficial da
Simon também publica uma tabela de "estimated living expenses" (Housing
$10.206/ano, Food $6.273/ano, Travel $2.250/ano, Personal $2.466/ano, Books
$1.310/ano), mas essas categorias não mapeiam 1:1 para as colunas do schema
(ex.: "Travel" ali é viagem internacional de ida/volta, não transporte local
mensal; "Housing" é uma estimativa genérica de moradia compartilhada, não
necessariamente "quarto individual + banheiro privativo" que o schema pede).
Substituir os valores já pesquisados por essa soma agregada arriscaria trocar
um dado específico por um de escopo diferente -- não foi feito.

O único componente que faltava e mapeia 1:1 com o schema é 'Seguro Saúde
(programa)': a tabela oficial tem uma linha "Health Insurance" dedicada
($3.906/ano) -- usada para preencher esse campo, e a partir dela (mais os 8
componentes mensais já existentes x 12 meses) foi possível calcular
'Custo Vida Total (programa) USD', que também estava "Verificar".

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://simon.rochester.edu/programs/full-time-ms/tuition-financial-aid
  (tuition por programa 2026-27, fees obrigatórias, seguro saúde, tabela de
  living expenses -- datada de 14/nov/2025)
- https://simon.rochester.edu/programs/full-time-ms/admissions
  (deadlines por round, taxa de inscrição $90, documentos exigidos, GMAT/GRE
  opcional, scores de inglês recomendados)
- https://simon.rochester.edu/programs/full-time-ms-in-business-analytics/academics/class-profile
  (Classe de 2025: 91 alunos, GPA médio 3.3, GMAT 630-740, GRE Q159-170/
  V149-164, IELTS médio 7.3, TOEFL médio 103, Duolingo médio 130,
  experiência média 1.7 anos, 59% direto da graduação, >90% recebem bolsa)
- Employment report (dados 2025, atualizado 1/abr/2026): 73% empregados em
  até 6 meses; salário médio (base + bônus) $75.801 -- BATE exatamente com o
  valor já cadastrado na planilha, confirmando que esse dado já estava
  correto (sem alteração).
- Ranking QS/US News encontrados (QS #101 global / #38 EUA; US News MBA #34)
  são rankings GERAIS da escola/MBA, não específicos do ranking QS Business
  Analytics Masters usado nas outras linhas desta planilha -- por isso
  'Ranking QS' e 'Ranking US News' foram mantidos como "Verificar" para não
  misturar categorias de ranking diferentes.
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
    "Nome da Universidade": "U of Rochester",
    "Estado": "NY",
    "Cidade": "Rochester",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Intl) USD": 68000,  # CORRECAO: planilha tinha 72.000; oficial 2026-27 = 68.000/ano
    "Fees Adicionais USD": 2867,  # Activity Fee 800 + Mandatory Health Fee 1.027 + Intl Services Office Fee 140 + Intl Student Programming Fee 900 (único)
    "Seguro Saúde (programa) USD": 3906,  # prêmio anual oficial 2026-27 (programa de 12 meses = 1 ano de cobertura)
    "Custo Vida Total (programa) USD": 28626,  # (1.245+350+15+144+50+60+46+150) x 12 meses + 3.906 (seguro saúde)
    "Tamanho Turma": 91,  # Classe de 2025 (dado mais recente encontrado)
    "GPA Médio Admitidos": 3.3,  # Classe de 2025
    "GMAT Médio": 685,  # ponto médio da faixa oficial reportada (630-740); teste opcional, poucos submetem
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 (média da turma: 103) / IELTS mín. 7.0, "
        "nenhuma seção abaixo de 6.5 (média da turma: 7.3) / Duolingo "
        "mín. 125 (média da turma: 130)"
    ),  # CORRECAO/detalhamento: planilha tinha só "TOEFL/IELTS" vago
    "Taxa Aplicação USD": 90,  # CORRECAO: planilha tinha 0; isenção disponível p/ militares/veteranos e outros elegíveis
    "Deadline Aplicação": (
        "Fall: Round 1 15/out (decisão 15/dez); Round 2 5/jan (decisão "
        "15/mar); Round 3 1/mar (decisão 15/abr); Round 4 1/mai (decisão "
        "1/jun)"
    ),  # CORRECAO: planilha tinha só "~Jul 1 (intl rolling)"
    "Rounds Aplicação": (
        "4 rounds para entrada em Fall (ver Deadline de Aplicação)"
    ),
    "Documentos Necessários": (
        "Application online, resume em formato de negócios (máx. 2 "
        "páginas), 1 ensaio (250-500 palavras), transcripts oficiais de "
        "todas as instituições, GMAT/GRE/Executive Assessment (opcional), "
        "TOEFL/IELTS/Duolingo (internacionais), 2 referências, "
        "vídeo-ensaio"
    ),
    "Experiência Profissional": (
        "Não obrigatória — 59% da turma entra direto da graduação "
        "(<0,6 anos de experiência); média geral da turma: 1,7 anos"
    ),
    "Taxa Empregabilidade %": 73,  # empregados em até 6 meses, dados 2025 (atualizado 1/abr/2026)
    "Tipos de Bolsa": (
        "Merit (automática, sem inscrição separada; mais de 90% dos "
        "admitidos recebem); bolsa garantida de $15.000 para ex-alunos de "
        "graduação/mestrado/doutorado da U of Rochester; mínimo de "
        "$10.000/ano para egressos de escolas parceiras da Simon"
    ),  # detalhamento: planilha tinha só "Merit (>90% recebem)"
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
        raise ValueError("Nenhuma linha encontrada para U of Rochester (Rochester, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Rochester, Rochester, NY)")

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
