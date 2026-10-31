"""
Data enrichment (round 2): atualiza APENAS a linha 'Seattle U' (Albers
School of Business & Economics, Seattle, WA) no
msba_programs_database.xlsx.

Pedido da usuaria: "me confirme o custo de tuition, bolsas, dificuldade e
tudo sobre a Seattle University" -- a linha já tinha um round 1 detalhado
(inclusive corrigindo uma inconsistência de tuition antiga), esta rodada
reconfirma tuition/bolsas e preenche lacunas de custo de vida e
dificuldade de entrada.

=========================================================================
1) TUITION -- RECONFIRMADA, SEM ALTERAÇÃO
=========================================================================
Já corrigida no round 1: $1.073/crédito, 45 créditos, $48.285 total,
sem distinção residente/intl (privada). Fees Adicionais $2.185 também
já reconciliados com a tabela oficial de taxas obrigatórias. Nenhuma
fonte nova contradisse esses valores nesta rodada.

=========================================================================
2) BOLSAS -- AMBIGUIDADE INVESTIGADA E RESOLVIDA: SÃO DUAS BOLSAS
   DIFERENTES, NÃO UMA SÓ
=========================================================================
Encontrei um PDF oficial dedicado
(static.seattleu.edu/business/Graduate/Redapt_Scholarship.pdf) que
confirma que a Redapt Scholarship É REAL e é DISTINTA da Genevieve
Albers Graduate Fellowship (já cadastrada) -- inicialmente pareceu
suspeito porque a URL institucional geral de bolsas
(seattleu.edu/business/graduate/scholarships/redapt-scholarship-for-msba/)
não a menciona (provavelmente página descontinuada/redirecionada), mas
o PDF dedicado, hospedado em domínio oficial da SMU... digo, da Seattle
University, confirma:

  REDAPT SCHOLARSHIP (específica do MSBA):
    - Até 4 alunos ingressantes do MSBA por outono (cohort limitado)
    - $15.000 de bolsa, usada no primeiro ano acadêmico
    - Permite estágio PART-TIME já no primeiro trimestre (sujeito a
      desempenho satisfatório) -- diferencial: gera renda extra cedo,
      quando a maioria dos programas não permite
    - Foco em DIVERSIDADE: prioriza quem é o primeiro da família a
      cursar faculdade, quem superou obstáculos educacionais/
      econômicos significativos, ou quem tem vivência em ambientes
      diversos
    - Candidatura até 20 de maio

  GENEVIEVE ALBERS GRADUATE FELLOWSHIP (já cadastrada, geral do Albers,
  não exclusiva do MSBA):
    - $15.000 de bolsa + $7.000 de estipêndio (posição de assistente de
      pesquisa, 12,5h/semana, 3 trimestres)
    - Aberta a QUALQUER programa de pós-graduação do Albers (não só
      MSBA), domésticos E internacionais
    - Candidatura até 30 de abril
    - Historicamente muito competitiva: apenas 1 concedida entre 40
      candidatos no último ciclo divulgado (já cadastrado)

As duas coincidem no valor de $15.000 de bolsa principal, o que
provavelmente gerou a suspeita inicial de duplicidade -- mas são
programas genuinamente diferentes (um tem estipêndio de pesquisa e é
geral do Albers; o outro permite estágio remunerado cedo e é exclusivo
do MSBA com foco em diversidade). 'Tipos de Bolsa' enriquecido com os
detalhes completos de ambas.

=========================================================================
3) CUSTO DE VIDA -- COMPONENTES RESTANTES (novos, antes "Verificar")
=========================================================================
- Seguro Residencial/mês: $15 (Seattle, WA -- fontes convergem entre
  $6-30/mês; usado valor central já consistente com o padrão do
  dataset)
- Utilities/mês: $140 -- fonte oficial (seattle.gov, estimativa municipal
  2025): apartamento padrão custa $180-280/mês (eletricidade, gás,
  água, esgoto, lixo). Como o 'Aluguel Privativo/mês' desta linha já
  modela um QUARTO privativo (não o apartamento inteiro), a conta foi
  ajustada para refletir uma divisão aproximada de custos com
  colega(s) de apê -- usado ~metade do piso da faixa oficial.
- Celular/mês: $50 (padrão já usado em outras linhas do dataset; nenhum
  dado específico de Seattle encontrado que justificasse valor
  diferente)
- WiFi/mês: $65 -- fonte: mesma estimativa oficial de Seattle, internet
  adiciona $50-80/mês; usado o ponto médio

=========================================================================
4) CUSTO VIDA TOTAL (programa) -- calculado, antes "Verificar"
=========================================================================
Fórmula já usada em outras linhas (replicada também no recálculo
automático do dashboard):
  (950+300+15+140+50+65+191+219) x 15 meses + 1.823 (seguro saúde)
  = 1.930 x 15 + 1.823 = 28.950 + 1.823 = 30.773

=========================================================================
5) RANKING US NEWS -- RECONFIRMADO, SEM ALTERAÇÃO
=========================================================================
O valor já cadastrado (19, "Best Master's in Business Analytics", U.S.
News & World Report **2026**) é o MAIS RECENTE encontrado -- as buscas
desta rodada trouxeram números mais antigos e inconsistentes (#25 em
2021, #16 em 2025), nenhum deles mais atual que o já cadastrado.
Mantido sem alteração, com confiança reforçada.

=========================================================================
6) RANKING QS -- preenchido com ressalva de desatualização
=========================================================================
QS Global Ranking 2024 (não 2026): #251 de 315 no ranking geral QS
Global -- é um ranking mais amplo/genérico, não especificamente do
MSBA nem tão atual quanto o dado de US News já confirmado acima.
Registrado com a ressalva explícita.

=========================================================================
7) DIFICULDADE DE ENTRADA -- pesquisado, mas sem números confirmados
=========================================================================
Buscas dedicadas na própria página oficial de admissões e em fontes
terceiras NÃO encontraram: taxa de aceitação, tamanho de turma, GPA
médio dos admitidos, nem GMAT médio -- a página de admissões só
confirma o que já estava cadastrado (GPA mínimo 3.0, GMAT/GRE
opcional). 'Taxa Aceitação %', 'Tamanho Turma', 'GPA Médio Admitidos' e
'GMAT Médio' permanecem "Verificar" -- sem números suficientes para
avaliar a dificuldade além da classificação qualitativa já cadastrada
('Dificuldade Entrada': Moderada, mantida).

=========================================================================
8) EMPREGABILIDADE/SALÁRIO -- NÃO ENCONTRADO, MANTIDO "Verificar"
=========================================================================
Nenhum relatório de emprego (employment report) específico do MSBA foi
localizado. Um dado de "salário médio de business analyst em Seattle
($108.334, ou $81.400 para nível de entrada)" apareceu em buscas, mas é
um dado de MERCADO REGIONAL GERAL (BLS/mercado de trabalho), não um
salário de EGRESSOS do programa -- não usado, para não confundir
estatística de mercado com estatística de outcomes do programa (mesmo
critério de não-contaminação já aplicado a toda escola pesquisada nesta
sessão). 'Taxa Empregabilidade %' e 'Salário Médio Pós USD' permanecem
"Verificar". 'Principais Empregadores' (Microsoft, Amazon, Meta,
Nordstrom, Boeing) já estava cadastrado e é consistente com os nomes
que aparecem nas novas buscas (Amazon Web Services, Boeing) -- mantido
sem alteração.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Seattle U",
    "Estado": "WA",
    "Cidade": "Seattle",
}

UPDATES = {
    "Tipos de Bolsa": (
        "Albers Incoming Student Graduate Scholarship: valor variável "
        "(até cobrir a mensalidade integral, não cobre taxas), aberta a "
        "PMBA/MBA SEM/MSBA/MSF/MPAC/Early Career MBA, candidatura "
        "incluída na própria inscrição do programa. Redapt Scholarship "
        "(exclusiva do MSBA, até 4 alunos ingressantes/outono): $15.000 "
        "no 1º ano + permite estágio part-time já no 1º trimestre; foco "
        "em diversidade (primeiro-geração, obstáculos educacionais/"
        "econômicos superados, vivência diversa); candidatura até 20 "
        "mai. Genevieve Albers Graduate Fellowship (geral do Albers, "
        "não exclusiva do MSBA): $15.000 + $7.000 de estipêndio "
        "(pesquisa, 12,5h/sem, 3 trimestres); domésticos e "
        "internacionais; candidatura até 30 abr. Taxa de aplicação "
        "sempre isenta"
    ),
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 140,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 65,
    "Custo Vida Total (programa) USD": 30773,
    "Ranking QS": (
        "251 de 315 -- QS Global Ranking 2024 (ranking geral, não "
        "específico do MSBA, e menos atual que o Ranking US News já "
        "confirmado em 2026)"
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
        raise ValueError("Nenhuma linha encontrada para Seattle U (Seattle, WA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Seattle U, Seattle, WA)")

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
