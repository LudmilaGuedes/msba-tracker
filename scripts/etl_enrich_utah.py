"""
Data enrichment: atualiza APENAS a linha da 'U of Utah' (David Eccles School
of Business, Salt Lake City, UT) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM. Deep
research em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida local.

Esta linha tinha muito mais "Verificar" que as demais (inclusive Duração e
Total Créditos, campos normalmente básicos) -- a maioria das páginas oficiais
da Eccles (finances/, admissions/, página principal do programa) retornaram
conteúdo truncado durante a coleta automatizada, então parte da pesquisa
usou agregados terceiros triangulados entre si (2-3 fontes independentes
convergindo no mesmo valor) em vez de uma única página oficial.

ACHADO: 'Custo/Crédito (Intl/Residente) USD' já estava correto ($1.300,
tarifa fixa sem diferença intl/residente -- consistente com 'Caminho Tuition
Residente' já cadastrado como "Flat tuition (sem premium intl)"). Com o
Total de Créditos agora confirmado (30, fonte oficial: página de academics
da Eccles), foi possível CALCULAR 'Tuition Total' pela primeira vez:
30 x $1.300 = $39.000.

Custo de vida: as páginas oficiais de cost-of-attendance da Eccles/
Financial Aid da U of Utah não puderam ser lidas de forma confiável
(bloqueio/truncamento), e os únicos dados de aluguel encontrados via
terceiros eram médias genéricas de mercado em Salt Lake City (não
específicas de "quarto individual + banheiro privativo" para pós-
graduandos) -- mesmo critério já aplicado em etl_enrich_uc_davis.py e
etl_enrich_umn.py: não usados, para não inventar precisão que a fonte não
tem. Todos os 8 componentes mensais de custo de vida permanecem "Verificar".

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- Página de academics da Eccles (via busca): confirma 30 créditos (fonte
  mais recente, corrigindo uma cifra antiga de terceiros de 33 créditos)
- Duração: triangulada entre TopMBA (16 meses), Shiksha (15-18 meses) e
  GyanDhan (15 meses) -- usado 16 meses como estimativa central, sem fonte
  oficial única disponível
- Página de admissions da Eccles (via busca): GPA mínimo 3.0; GMAT/GRE
  opcional (não aceita GRE at-home); TOEFL competitivo 90; IELTS
  obrigatório 7.0; 2 cartas de recomendação (idealmente de supervisores
  profissionais); sem exigência mínima de experiência profissional
  (já cadastrado corretamente, sem alteração)
- Busca por deadlines: prioritário 1/set (decisão antecipada + prioridade
  de escolha de cursos); planilha já tinha "Jun 1 (Fall, first-come)" --
  interpretado como o prazo rolling final do mesmo ciclo, mantido e
  complementado (não contraditório: prioritário ~1 ano antes, rolling até
  ~1/jun do ano de início)
- Taxa de inscrição: $55 doméstico / $65 internacional (usado o valor
  internacional, foco do dataset)
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
    "Nome da Universidade": "U of Utah",
    "Estado": "UT",
    "Cidade": "Salt Lake City",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 16,  # estimativa central triangulada entre 3 fontes terceiras (15-18 meses); sem fonte oficial única
    "Total Créditos": 30,  # fonte oficial (página de academics da Eccles), corrige cifra antiga de 33 créditos vista em terceiros
    "Tuition Total (Intl) USD": 39000,  # 30 créditos x $1.300/crédito (já confirmado, tarifa fixa)
    "Tuition Total (Residente) USD": 39000,  # tarifa fixa, sem diferença intl/residente
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Não obrigatório (opcional); recomendado enviar se já tiver "
        "pontuação; o programa não aceita a versão at-home do GRE"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT competitivo: 90 / IELTS Academic obrigatório: 7.0 "
        "(de 9)"
    ),  # CORRECAO/detalhamento: planilha tinha só "TOEFL/IELTS" vago
    "Taxa Aplicação USD": 65,  # internacional; doméstico é $55
    "Deadline Aplicação": (
        "Prioritário: 1/set (decisão antecipada + prioridade de escolha "
        "de cursos); admissão contínua (rolling) até ~1/jun do ano de "
        "início, sujeita a vagas disponíveis"
    ),  # detalhamento: planilha tinha só "Jun 1 (Fall, first-come)"
    "Rounds Aplicação": "Prioritário + admissão contínua (rolling) até preencher vagas",
    "Documentos Necessários": (
        "Transcripts oficiais de todas as instituições, 2 cartas de "
        "recomendação (idealmente de supervisores profissionais), "
        "GMAT/GRE (opcional), TOEFL/IELTS (candidatos internacionais)"
    ),
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
        raise ValueError("Nenhuma linha encontrada para U of Utah (Salt Lake City, UT). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Utah, Salt Lake City, UT)")

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
