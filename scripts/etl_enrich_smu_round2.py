"""
Data enrichment (round 2): atualiza APENAS a linha 'SMU' (Cox School of
Business, Dallas, TX) no msba_programs_database.xlsx.

Pedido da usuaria: "não tem nada sobre o MSBA da SMU" -- na verdade a
linha já tinha bastante coisa preenchida de um enrichment anterior
(etl_enrich_smu.py), mas com lacunas reais (custo de vida detalhado,
rankings QS, GPA/GMAT médio, empregadores, pré-requisitos) e uma
inconsistência a resolver (ver item 1). Esta rodada faz nova pesquisa
completa e preenche o máximo possível.

=========================================================================
1) ACHADO CRÍTICO -- RESOLVE INCONSISTÊNCIA DO ROUND 1
=========================================================================
O script etl_enrich_smu.py (round 1) tinha em seu UPDATES:
  "Taxa Empregabilidade %": 96  # atualização (planilha tinha 82;
  fonte mais recente confirma 96%+ em 6 meses)
Só que a CÉLULA ATUAL continua mostrando 82 -- ou seja, esse valor de 96
nunca chegou a ficar de fato registrado (execução anterior não
persistiu esse campo específico, por motivo não documentado).

Nesta rodada, PESQUISA NOVA E MAIS CUIDADOSA reverte essa decisão: o
número "96%+" só aparece em agregadores de baixa confiabilidade
(mim-essay.com, ameerkhatri.com), sem citação especifica do relatório
oficial. Já o número 82% aparece de forma CONSISTENTE e com citação
especifica em múltiplas buscas independentes: "82% of the Class of 2025
achieved a career or educational outcome within six months of
graduation from the SMU Cox Master of Science in Business Analytics
program" -- frase especifica o suficiente para parecer extraída
diretamente do relatório oficial (EMPLOYMENT REPORT MS Business
Analytics Class of 2025, PDF que não pôde ser lido diretamente por
bloqueio de parsing, mas cujo conteúdo está sendo citado por terceiros
de forma consistente). CONCLUSÃO: 82% (já cadastrado) é o valor correto
-- o "96%" do round 1 NÃO foi aplicado propositalmente nesta rodada, por
não resistir a nova verificação. Mantido 82, sem alteração.
Um outro número ("79% career/education outcome, 67% full-time job
accepted") apareceu em UMA ÚNICA sintese de busca, sem confirmação
cruzada -- descartado por falta de repetição/confiabilidade.

Já o salário ($79.800) e a taxa de aceitação (46%) foram RECONFIRMADOS
por múltiplas fontes nesta rodada -- sem alteração.

=========================================================================
2) SEGURO SAÚDE (novo, antes "Verificar")
=========================================================================
Fonte oficial (smu.edu Bob Smith Health Center + confirmação cruzada):
SMU Student Health Insurance Plan (SHIP), obrigatório para todo aluno
com 1+ crédito (internacionais sem exceção): $2.115,00/semestre. MSBA
dura 10 meses = 2 semestres (Fall+Spring) -> $2.115 x 2 = $4.230.

=========================================================================
3) CUSTO DE VIDA -- DEMAIS COMPONENTES MENSAIS (novos, antes "Verificar")
=========================================================================
- Seguro Residencial/mês: ~$18 (média de renters insurance em Dallas,
  TX, 2026 -- múltiplas fontes convergem entre $12-25/mês; usado o
  ponto médio mais citado)
- Celular/mês: $50 (estimativa padrão já usada em outras linhas do
  dataset para este tipo de despesa; não foi encontrado dado
  especifico de Dallas que justificasse valor diferente)
- WiFi/mês: $60 (idem)
- Transporte/mês: $126 (DART Local Monthly Pass, tarifa cheia
  publicada; NÃO foi confirmado se a SMU participa do "DART Higher
  Education Program" com tarifa reduzida para seus alunos -- usado o
  valor cheio por precaução, já que reduzido/gratuito só foi confirmado
  para Dallas College, uma instituição DIFERENTE)
- Outros Gastos/mês: $150 (estimativa padrão já usada em outras linhas)

=========================================================================
4) CUSTO VIDA TOTAL (programa) -- calculado, antes "Verificar"
=========================================================================
Fórmula já usada em outras linhas deste dataset (confirmada e replicada
também no recálculo automático implementado no dashboard):
  (soma dos 8 componentes mensais) x duração (10 meses) + seguro saúde
  = (1100+300+18+225+50+60+126+150) x 10 + 4.230
  = 2.029 x 10 + 4.230 = 20.290 + 4.230 = 24.520

=========================================================================
5) MATÉRIAS PRÉ-REQUISITO (novo, antes "Verificar")
=========================================================================
Fonte oficial (currículo MSBA, catalog.smu.edu indexado + página do
programa): NÃO há pré-requisito formal obrigatório de curso -- o
programa oferece, em vez disso, um "self-paced bridge course" de
estatística (refresher opcional) e um período de self-study no verão
antes do início, para preparar o aluno.

=========================================================================
6) RANKING QS -- preenchido com ressalva (mesmo padrão já usado no
   campo 'Ranking US News' desta mesma linha, que já tinha uma nota
   similar)
=========================================================================
Cox School of Business (não o MSBA especificamente) ranqueia #43/108
nos EUA no QS Business ranking 2026 -- ranking GERAL da escola, não do
programa. Um dado mais antigo e menos confiável (QS Masters in Business
Analytics 2023) colocava a SMU na faixa 51-60 -- desatualizado (2023,
não 2026) e não reconfirmado nesta rodada. Ambos registrados com
ressalva explícita.

=========================================================================
7) NÃO ENCONTRADO / MANTIDO "Verificar" (mesmo após buscas adicionais)
=========================================================================
- GPA Médio Admitidos, GMAT Médio: programa é test-optional, e não foi
  encontrada nenhuma média divulgada especificamente para o MSBA.
- Principais Empregadores: a única lista de empregadores encontrada
  (JPMorgan, Goldman Sachs, Houlihan Lokey, BCG, PMG, Deloitte, Bank of
  America, Jefferies, EY, Citi, KPMG) é EXPLICITAMENTE do programa de
  MBA (Full-Time MBA, "Class of 2025", dados "2026 embargo") -- NÃO
  usada aqui, mesmo critério de não-contaminação já aplicado a
  praticamente toda escola pesquisada nesta sessão.
- TA/RA Disponível, Processo TA/RA, Graduate Assistantships: encontrada
  apenas uma descrição GENÉRICA de assistantships na SMU como um todo
  (20-25h/semana, ajuda a se sustentar financeiramente), sem confirmar
  se está disponível para o MSBA especificamente nem se internacionais
  são elegíveis -- não usada, para não presumir algo não confirmado.
- Ranking US News: já tinha nota própria de que é ranking geral da SMU,
  não do MSBA -- mantido sem alteração.

=========================================================================
8) QUÃO REALISTA (Intl) -- preenchido, antes "Verificar"
=========================================================================
Classificado como "Moderada": bolsas de mérito são automáticas (sem
candidatura separada), mas em "número limitado"; nenhuma Graduate
Assistantship confirmada para o programa; e internacionais têm
restrição prática no primeiro semestre (não podem fazer estágio, só
podem usar CPT a partir de janeiro) -- limita a geração de renda extra
logo no início do programa.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "SMU",
    "Estado": "TX",
    "Cidade": "Dallas",
}

UPDATES = {
    "Seguro Saúde (programa) USD": 4230,  # SHIP oficial $2.115/semestre x 2 semestres
    "Seguro Residencial/mês USD": 18,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 126,  # DART Local Monthly Pass, tarifa cheia
    "Outros Gastos/mês USD": 150,
    "Custo Vida Total (programa) USD": 24520,  # (1100+300+18+225+50+60+126+150)x10 + 4230
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal obrigatório -- self-paced bridge "
        "course de estatística (refresher opcional) e self-study de "
        "verão disponíveis para preparação antes do início"
    ),
    "Ranking QS": (
        "Cox School of Business (não o MSBA especificamente): #43/108 "
        "nos EUA no QS Business Ranking 2026. Dado mais antigo e não "
        "reconfirmado: QS Masters in Business Analytics 2023 colocava a "
        "SMU na faixa 51-60 (desatualizado)"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- bolsas de mérito automáticas mas em número "
        "limitado; nenhuma Graduate Assistantship confirmada para o "
        "MSBA; internacionais não podem estagiar no 1º semestre (só "
        "CPT a partir de janeiro), o que limita renda extra no início "
        "do programa"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para SMU (Dallas, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SMU, Dallas, TX)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
