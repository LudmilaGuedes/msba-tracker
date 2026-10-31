"""
Data enrichment: atualiza APENAS a linha da 'RIT' (Rochester Institute
of Technology, Saunders College of Business, Rochester, NY) no
msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (rit.edu). A página geral de tuition da
RIT trouxe números internamente INCONSISTENTES ($2.690/crédito para
1-11 créditos vs. "$64.542 por semestre" para 12-18 créditos, que
implicaria >$5.000/crédito -- não bate) e claramente não específicos
deste programa; por isso NÃO foram usados para alterar a tuition já
presente na planilha (que já tinha valores limpos e consistentes:
$1.970/crédito × 30 créditos = $59.100, provavelmente de uma fonte
mais específica de uma rodada de pesquisa anterior a esta sessão).

CORREÇÃO/ATUALIZAÇÃO DE NOME: 'Nome Oficial do Programa' atualizado
para 'MS in Business Analytics and AI' -- o site atual da RIT já não
usa mais apenas "Business Analytics", refletindo o mesmo padrão de
rebranding com IA visto em outras escolas nesta sessão (Buffalo, Pitt).
STEM confirmado com nota oficial de STEM-OPT de até 3 anos.

TUITION: como a RIT é privada com tuition flat (já documentado na
planilha: 'Caminho Tuition Residente' = 'N/A (privada)'), 'Custo/
Crédito (Residente)' e 'Tuition Total (Residente)' foram preenchidos
com os MESMOS valores já confirmados para Intl ($1.970/crédito;
$59.100 total) -- não há distinção de residência numa universidade
privada.

'Fees Adicionais USD' preenchido com ESTIMATIVA de confiança moderada:
$470 (Student Activity Fee) + $630 (Student Health Services Fee) =
$1.100/semestre (valores oficiais, mas de página GERAL da RIT, não
específica deste programa) × 2 semestres assumidos (não confirmado
oficialmente o número exato de semestres do programa de 12 meses) =
$2.200. Documentado como estimativa por depender de uma suposição de
estrutura de semestres.

'GMAT/GRE Obrigatório' refinado com achado IMPORTANTE para o perfil de
busca (aluna internacional com diploma estrangeiro): o teste é
OBRIGATÓRIO especificamente para "international degree holders"
(portadores de diploma obtido fora dos EUA) -- outros candidatos podem
ter dispensa, mas esse não é o caso de uma aluna formada no Brasil.
'GPA Mínimo' preenchido: 3.0 (recomendado, não necessariamente um piso
rígido). 'Testes Inglês Aceitos' refinado com mínimos oficiais: TOEFL
88 (iBT)/4.5 (novo formato), IELTS 6.5, Duolingo 130, PTE Academic 60.

'Documentos Necessários' preenchido com a lista oficial, incluindo um
detalhe ÚTIL: cartas de recomendação são exigidas para TODOS os
programas de pós-graduação da RIT, EXCETO os oferecidos pelo College
of Business -- ou seja, este programa (Saunders College of Business)
NÃO exige cartas de recomendação. 'Taxa Aplicação USD' preenchido:
$65 (dispensada para alunos/ex-alunos da RIT).

'Processo TA/RA' preenchido com achado relevante: assistantships são
competitivas e seletivas, mas EXPLICITAMENTE "do not cover tuition
costs" -- fornecem apenas wages/stipend por trabalho de pesquisa,
ensino ou administrativo; maioria concedida após a matrícula, não no
momento da admissão.

'Taxa Empregabilidade %' preenchido com 86 (86,4% empregados em até 6
meses, dado oficial e específico deste programa, com "knowledge rate"
de 73,3% -- acima da média nacional NACE de 41% -- documentado como
contexto de confiabilidade do dado). 'Salário Médio Pós USD'
preenchido com $84.000 (média; mediana oficial é $67.000, mencionada
aqui pois a diferença é grande e relevante para expectativas
realistas). 'Principais Empregadores' preenchido com uma amostra da
lista oficial (TikTok, EY, Dow Jones & Co., CooperVision, Citibank,
Tata Consultancy Services).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Seguro Saúde (programa) USD' (mandatório confirmado, mas sem valor de
prêmio específico divulgado), 'Tamanho Turma', 'GPA Médio Admitidos',
'GMAT Médio', 'Taxa Aceitação %', 'Ranking US News', 'Ranking QS'
(a página cita rankings de fontes não-tradicionais como "College
Factual" e "TFE Times", que não correspondem ao formato esperado
desses dois campos específicos, então não foram usados), 'Entrevista'
(não mencionado em nenhuma fonte oficial consultada), 'Utilities/mês',
'Custo Vida Total (programa) USD'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.rit.edu/study/business-analytics-ms (nome atual,
  formato, STEM/OPT, currículo, GPA 3.0 recomendado, GMAT/GRE para
  diploma internacional, mínimos de inglês, outcomes: 86% placement,
  $84K salário médio, rankings não-tradicionais)
- https://www.rit.edu/admissions/tuition-and-fees (fees gerais —
  Student Activity Fee, Student Health Services Fee — mas números de
  tuition geral inconsistentes, não usados)
- https://www.rit.edu/admissions/financial-aid/graduate-funding-sources
  (bolsas de mérito 10-40%, RIT Master Plan/Up, assistantships não
  cobrem tuition)
- https://www.rit.edu/admissions/graduate/application-instructions
  (taxa de aplicação $65, documentos exigidos, cartas de recomendação
  dispensadas para College of Business)
- https://www.rit.edu/careerservices/study/business-analytics-and-ai-ms
  (outcomes oficiais e específicos: 86,4% empregabilidade, $84K/$67K
  salário médio/mediano, empregadores)
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
    "Nome da Universidade": "RIT",
    "Estado": "NY",
    "Cidade": "Rochester",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "MS in Business Analytics and AI",
    "Custo/Crédito (Residente) USD": 1970,  # privada, tarifa flat (igual ao Intl)
    "Tuition Total (Residente) USD": 59100,  # privada, tarifa flat (igual ao Intl)
    "Fees Adicionais USD": 2200,  # estimativa: (Student Activity $470 + Health Services $630) x 2 semestres assumidos
    "GMAT/GRE Obrigatório": (
        "OBRIGATÓRIO para portadores de diploma internacional "
        "(obtido fora dos EUA); outros candidatos podem ter dispensa "
        "(critério exato não detalhado)"
    ),
    "GPA Mínimo": "3.0 (recomendado)",
    "Testes Inglês Aceitos": "TOEFL 88 (iBT) / 4.5 (novo formato); IELTS 6.5; Duolingo 130; PTE Academic 60",
    "Taxa Aplicação USD": 65,  # dispensada para alunos/ex-alunos da RIT
    "Documentos Necessários": (
        "Transcripts oficiais de todas as instituições; resume/CV; "
        "personal statement (sem limite de tamanho); GMAT/GRE "
        "(obrigatório para diploma internacional); teste de inglês "
        "(dispensável para países de língua inglesa ou "
        "estudo/trabalho prévio nos EUA); CARTAS DE RECOMENDAÇÃO NÃO "
        "EXIGIDAS (exceção aplicável a todos os programas do College "
        "of Business, incl. este)"
    ),
    "Processo TA/RA": (
        "Competitivo e seletivo; NÃO cobre custos de tuition — "
        "fornece apenas wages/stipend por trabalho de pesquisa, "
        "ensino ou administrativo; maioria concedida após a "
        "matrícula, não no momento da admissão"
    ),
    "Taxa Empregabilidade %": 86,  # 86,4% oficial, específico do programa (6 meses pós-formatura)
    "Salário Médio Pós USD": 84000,  # média oficial; mediana é $67.000 (diferença grande, documentada no docstring)
    "Principais Empregadores": "TikTok, Ernst & Young (EY), Dow Jones & Co., CooperVision, Citibank, Tata Consultancy Services",
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
        raise ValueError("Nenhuma linha encontrada para RIT (Rochester, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (RIT, Rochester, NY)")

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
