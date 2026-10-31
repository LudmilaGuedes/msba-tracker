"""
Data enrichment: atualiza APENAS a linha 'Oklahoma State' (Spears School
of Business / Watson Graduate School of Management, Stillwater, OK) no
msba_programs_database.xlsx.

FONTE: EMAIL DIRETO da Graduate Programs Office (Watson Graduate School
of Management, Spears School of Business) para a usuaria, respondendo
perguntas especificas sobre tuition, bolsas/assistantships, requisitos e
admissao do MS-BAnDS (set/2026). Fonte de PRIMEIRA MAO -- pesa mais que
qualquer agregador de terceiros, mesmo padrao ja usado com o email da UC
Cincinnati nesta mesma sessao de trabalho.

CORRECAO CRITICA DE TUITION -- confirma suspeita ja registrada na
auditoria anterior (etl_fix_verificacao_oficial_lote1.py): o email diz
"The total program cost for full time enrollment for 37 hours at the
current tuition rate is approximately $43,000 US (for non-residents of
Oklahoma). Resident tuition is approximately $19,200 US."
A planilha tinha 'Tuition Total (Intl)' = $20.786 -- exatamente o valor
que a auditoria ja suspeitava ser uma cifra ANUAL de um programa de 21
meses, nao o total. CONFIRMADO: a planilha subestimava o custo
internacional em quase $22.000 (mais que o DOBRO do valor real).
'Custo/Crédito' preenchido por CALCULO DERIVADO ($43.000/37 e
$19.200/37) -- o proprio email usa "37 hours" para o calculo (nao 38,
o total de creditos ja cadastrado), entao o valor por credito e uma
aproximacao, nao uma tarifa oficial publicada por credito. Ambos os
valores da fonte ja vem com a ressalva "approximately".

CORRECOES DE FORMATO/CALENDARIO:
- 'Início' estava "Fall/Spring" -- CORRIGIDO para "Fall": o email e
  explicito -- "The MS-BAnDS program only has fall (August) admission
  for full time enrollment."
- 'Deadline Aplicação' estava "Spring Oct 1" (confuso/incorreto) --
  CORRIGIDO com o calendario real: inscricoes abrem 1/out do ano
  anterior; prioritario 15/jan; final 1/mar; comite comeca a avaliar
  em meados de janeiro.
- 'Rounds Aplicação' preenchido: nao ha rounds formais -- e admissao
  ROLLING dentro da janela de inscricao (entrevista -> oferta de
  pre-admissao -> deposito de $300 em ate 1 semana, reembolsado apos
  ingresso; se nao pago, oferta e revogada e vaga vai pro proximo da
  fila de espera).

CORRECAO -- Experiência Profissional: estava "Não obrigatória", o que e
enganoso para uma candidata internacional especificamente -- o email diz
"Work experience of about two years is EXPECTED for admissions for
FOREIGN students; however, it may be waived for exceptional foreign
candidates". Para domesticos e "recommended but not required" (esse sim
bate com o texto antigo). Corrigido para refletir a distincao.

GMAT/GRE (novo, antes "Verificar"): OBRIGATORIO para internacionais e
para domesticos sem vinculo previo com a OSU, "no exceptions". Scores
minimos RECOMENDADOS (nao eliminatorios por si so): GMAT 600 / GRE 310.
Alunos de graduacao da propria OSU sao dispensados.
IMPORTANTE: o email deixa explicito que os testes de ingles (TOEFL/
IELTS/PTE/Duolingo) NAO SUBSTITUEM o GMAT/GRE -- sao exigencias
independentes, nao intercambiaveis. Nota acrescentada em 'Testes Inglês
Aceitos'.

MATERIAS PRE-REQUISITO e DOCUMENTOS (novos, antes "Verificar"):
sem pre-requisitos formais de curso antes da admissao, mas espera-se
base em Estatistica introdutoria e familiaridade basica com Python/SQL/R
(ambos cursados logo no inicio do programa). Documentos: bacharelado de
4 anos conferido (nao aceita diploma de 3 anos isoladamente, exceto se
somado a um mestrado); transcripts OFICIAIS apenas -- avaliacoes de
terceiros (ex.: WES) NAO sao aceitas nem revisadas; candidatura 100%
pelo sistema online; taxa de aplicacao obrigatoria e nao-dispensavel
(sem vouchers do programa) -- valor exato nao informado no email, por
isso 'Taxa Aplicação USD' permanece "Verificar".

BOLSAS E ASSISTANTSHIPS (novo, antes "Verificar" em quase todos os
campos): numero LIMITADO de Graduate Teaching Assistantships (GTA),
tipicamente concedidas SO APOS o 1o semestre already cursado, limitadas
a 2 SEMESTRES (nao o programa inteiro de ~4 semestres+estagios). Cobrem
APENAS tuition (nao fees) + salario de $1.400/mes + seguro saude
subsidiado (so para o aluno, nao dependentes). Para credenciais
excepcionais: um numero reduzido de GTA + bolsas de ~$1.500 concedidas
JA NO 1o SEMESTRE. Selecao ocorre no fim do verao, de forma competitiva
(GRE/GMAT, GPA de graduacao, anos de experiencia). O proprio email
adverte: "We will not guarantee any financial aid prior to admission;
new students should secure their own financing prior to arrival at
OSU." -- por isso 'Quão Realista (Intl)' foi classificado como "Baixa a
Moderada" (mais conservador que a maioria das outras linhas desta
sessao, que usam "Moderada"): aqui a propria universidade e explicita em
NAO prometer nada antes da chegada, e a cobertura do GTA (so tuition, so
2 semestres) e mais limitada que o padrao ja visto em FSU/USF/UCF/UGA.

STEM/CURRICULO (enriquecido): confirmado CIP code 52.1302 (Business
Statistics, lista IPEDS 2020) -- ja estava "Sim" em STEM Designated
(schema so aceita Sim/Não, entao o CIP code foi anexado a 'Destaques
Currículo'). Tambem adicionada a estrutura de estagio: ate 2 estagios
totalizando ate 7 meses (3 meses no primeiro verao + 4 meses no ultimo
semestre), com opcao excepcional de estagio avancado de ate 11 meses
(jun-nov) conforme necessidade da empresa contratante e do curriculo do
aluno.

NAO ALTERADO (nao endereçado pelo email, permanece como estava):
Taxa Aceitação % (31), Ranking QS (2), Taxa Empregabilidade % (100),
Principais Empregadores, GPA Mínimo/Médio, Tamanho Turma, Salário Médio
Pós, Taxa Aplicação USD, Caminho Tuition Residente (o email nao aborda
elegibilidade a tarifa de residente para internacionais -- diferente do
caso da Florida, nao ha lei estadual ja confirmada nesta sessao para
Oklahoma, entao nao presumido).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Oklahoma State",
    "Estado": "OK",
    "Cidade": "Stillwater",
}

UPDATES = {
    "Tuition Total (Intl) USD": 43000,  # CORRECAO: era 20786 (subestimava quase o dobro)
    "Tuition Total (Residente) USD": 19200,  # era "Verificar"
    "Custo/Crédito (Intl) USD": 1162.16,  # derivado: 43000/37 (email usa 37h no calculo, não 38)
    "Custo/Crédito (Residente) USD": 518.92,  # derivado: 19200/37
    "Início": "Fall",  # CORRECAO: era "Fall/Spring" -- só há admissão Fall p/ full-time
    "Deadline Aplicação": (
        "Somente admissão Fall; inscrições abrem 1/out do ano anterior; "
        "prioritário 15/jan; final 1/mar; comitê começa avaliação em "
        "meados de janeiro"
    ),
    "Rounds Aplicação": (
        "Sem rounds formais -- admissão rolling dentro da janela: "
        "entrevista -> oferta de pré-admissão -> depósito de $300 em "
        "até 1 semana (reembolsado após ingresso); se não pago, oferta "
        "é revogada e vaga vai para a fila de espera"
    ),
    "Experiência Profissional": (
        "Esperada (~2 anos) para candidatos internacionais, com "
        "dispensa possível para candidatos excepcionais; recomendada "
        "mas não obrigatória para domésticos"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório para internacionais e domésticos sem vínculo "
        "prévio com a OSU, sem exceções. Scores mínimos recomendados: "
        "GMAT 600 / GRE 310 (não eliminatórios por si só). Alunos de "
        "graduação da própria OSU são dispensados"
    ),
    "Matérias Pré-Requisito": (
        "Nenhum curso pré-requisito formal antes da admissão; espera-se "
        "base em Estatística introdutória e familiaridade básica com "
        "Python/SQL/R (ambos cursados logo no início do programa)"
    ),
    "Documentos Necessários": (
        "Bacharelado de 4 anos conferido (diploma de 3 anos só é aceito "
        "se somado a um mestrado); transcripts OFICIAIS apenas -- "
        "avaliações de terceiros (ex. WES) não são aceitas nem "
        "revisadas; candidatura 100% pelo sistema online; GMAT/GRE; "
        "taxa de aplicação obrigatória e não-dispensável"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 89/IELTS 6.5/Duolingo 120/PTE 60 -- exigido conforme "
        "idioma de instrução da graduação. IMPORTANTE: não substitui o "
        "requisito de GMAT/GRE -- são exigências independentes, não "
        "intercambiáveis"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Bolsas de ~$1.500 para candidatos excepcionais, concedidas já "
        "no 1º semestre. Seleção de bolsas/GTA ocorre no fim do verão, "
        "de forma competitiva (GRE/GMAT, GPA de graduação, anos de "
        "experiência). Universidade não garante nenhum auxílio antes da "
        "admissão -- orienta o aluno a se planejar financeiramente por "
        "conta própria antes de chegar aos EUA"
    ),
    "TA/RA Disponível": (
        "Sim (GTA), mas número limitado, tipicamente concedido só após "
        "completar o 1º semestre, e por apenas 2 semestres (não o "
        "programa inteiro) -- não garantido"
    ),
    "Processo TA/RA": (
        "GTA concedido tipicamente após o 1º semestre, limitado a 2 "
        "semestres; cobre APENAS tuition (não fees) + salário de "
        "$1.400/mês + seguro saúde subsidiado (só para o aluno). Para "
        "credenciais excepcionais, um número reduzido de GTA é "
        "concedido já no 1º semestre. Seleção no fim do verão, "
        "competitiva (GRE/GMAT, GPA, experiência)"
    ),
    "Graduate Assistantships": (
        "GTA: tuition-only waiver + $1.400/mês + seguro saúde "
        "subsidiado (aluno apenas); limitado a 2 semestres, "
        "tipicamente após o 1º semestre; não garantido antes da "
        "admissão"
    ),
    "Quão Realista (Intl)": (
        "Baixa a Moderada -- a universidade é explícita em não "
        "garantir nenhum auxílio antes da admissão ('secure their own "
        "financing prior to arrival'); GTA cobre só tuition (não fees) "
        "e só por 2 dos ~4 semestres do programa; mais limitado que o "
        "padrão observado em outras escolas pesquisadas nesta sessão"
    ),
    "Destaques Currículo": (
        "38 créditos; STEM (CIP 52.1302, Business Statistics, IPEDS "
        "2020); até 2 estágios somando até 7 meses (3 meses no 1º "
        "verão + 4 meses no último semestre), com opção excepcional de "
        "estágio avançado de até 11 meses conforme necessidade da "
        "empresa/currículo"
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
        raise ValueError("Nenhuma linha encontrada para Oklahoma State (Stillwater, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Oklahoma State, Stillwater, OK)")

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
