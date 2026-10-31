"""
Data enrichment (round 2): atualiza APENAS a linha da 'Fordham' (Gabelli
School of Business, New York, NY) no msba_programs_database.xlsx.

Objetivo: confirmar/corrigir tuition, bolsas, taxa de aplicação (e a
isenção no Round 1) e requisitos de aplicação, a pedido do usuário.

FONTES (todas oficiais, fordham.edu -- as páginas relevantes ficam
atrás de um gateway CAS que redireciona usuários anônimos de volta
para o conteúdo público; acessadas via download direto com cookie jar
persistente para sobreviver ao redirect duplo, em 2026-09-08):
- https://www.fordham.edu/student-financial-services/tuition-and-payments/graduate-tuition/gabelli-school-of-business/
- https://www.fordham.edu/gabelli-school-of-business/academic-programs-and-admissions/graduate-programs/graduate-admissions/financing-your-graduate-education/admissions-scholarships/
- https://www.fordham.edu/gabelli-school-of-business/academic-programs-and-admissions/graduate-programs/graduate-admissions/graduate-admissions-deadlines-/
- https://www.fordham.edu/gabelli-school-of-business/academic-programs-and-admissions/graduate-programs/graduate-admissions/graduate-admissions-requirements/

ACHADOS:

1. TUITION -- CORRIGIDO. Tabela oficial "Graduate Gabelli School of
   Business Tuition and Fees, 2026-2027 Academic Year": tarifa por
   crédito para MS Programs, "Entrants From Fall 2023 and later" =
   $2.303/crédito (o valor anterior na planilha, $2.201, está
   desatualizado -- corresponde a uma tabela de ano anterior).
   'Total Créditos' já confirmado no schema = 30 (bate com o currículo
   oficial: 10 cursos x 3 créditos, 5 obrigatórios + 5 eletivos, 3
   semestres). Novo total: 30 x $2.303 = $69.090 (era $66.030).
   Fordham é privada, sem distinção Intl/Residente -- mesmo valor nos
   dois campos, como já estava.

2. FEES ADICIONAIS -- preenchido (estava 'Verificar'). Taxas
   obrigatórias por termo/únicas da Graduate School (tabela oficial
   2026-2027): General Fee $128/termo, Technology Fee $249/termo,
   Health Insurance Fee $1.890 (termo Fall) + $3.342 (termo Spring,
   inclui Summer), International Student Service Fee $50 (só
   internacionais), Maintenance of Matriculation Fee $51,
   Enrichment Fee (MS Programs - New Students, só 1º termo) $3.939.
   Guardado como texto discriminado (não um único número), já que são
   taxas de naturezas e periodicidades diferentes -- somar tudo num
   único valor exigiria assumir o número exato de termos cursados
   (mesmo princípio de não empilhar estimativa sobre estimativa já
   usado nos enrichments do Baruch).

3. BOLSAS -- CORRIGIDO. A nota anterior ('Merit (69% recebem, $5k-full
   tuition)') misturava dado de outro programa: a página oficial
   confirma merit-based scholarships de $5.000 a $30.000 para
   Specialized Master's Degrees (não "full tuition" -- esse valor é
   específico do Full-time MBA Scholarship, $20.000-full tuition,
   programa diferente). Também existe o Dean's Academic &
   Professional/Excellence & Impact Scholarship, disponível a
   candidatos de Specialized Master's Programs on-site (inclui MSBA):
   valor fixo entre $15.000-$25.000 (definido anualmente). E o Mayor's
   Graduate Scholarship Program: mínimo de $8.000 para quem se
   qualifica e cursa um MS Program incluindo Business Analytics
   (critérios de elegibilidade específicos não detalhados nesta
   página -- provável programa municipal de NY, requer confirmação
   à parte se for relevante). Todos são considerados automaticamente a
   partir da própria aplicação de admissão (não há aplicação separada
   de bolsa). O percentual "69% recebem" não foi encontrado/confirmado
   nesta fonte -- removido por falta de confirmação.

4. TAXA DE APLICAÇÃO -- CORRIGIDA: $150 (era $130, desatualizado).
   CONFIRMADO: Round 1 tem isenção da taxa de aplicação (application
   fee waived) E consideração prioritária para bolsas -- a nota
   anterior na planilha estava correta nesse ponto.

5. DEADLINES/ROUNDS -- atualizados para o ciclo vigente (Fall 2027,
   página oficial de deadlines, tabela "MS Programs (on-campus)"):
   Round 1: 9 de outubro de 2026 (decisão 4 dez 2026) -- taxa isenta +
   bolsa prioritária. Round 2: 15 de janeiro de 2027 (decisão 19 mar
   2027) -- consideração prioritária para bolsas. Round 3: 19 de março
   de 2027 (decisão 14 mai 2027) -- último dia para candidatos
   internacionais que precisam de visto patrocinado do exterior. Round
   4: admissão rolling (candidatos domésticos e internacionais já
   presentes localmente).

6. REQUISITOS DE APLICAÇÃO -- CORRIGIDOS (página oficial "Graduate
   Admissions Requirements"):
   - Carta de recomendação: OPCIONAL para Specialized Master's (1
     acadêmica ou profissional) -- CORREÇÃO: a planilha tinha "2
     cartas de recomendação" como se fossem obrigatórias; isso é
     regra do MBA (1 carta profissional OBRIGATÓRIA), não do MSBA.
   - Ensaios: 2 obrigatórios para MSBA (não "statement of purpose
     500-750 palavras" genérico): Essay 1 (500 palavras, motivação
     para o mestrado / objetivos de carreira, obrigatório para todos)
     + Essay 3 (350 palavras, específico para candidatos de MSAIB,
     MSBA, MSF, MSMI, MSQF -- preparo quantitativo: cursos,
     certificações ou experiências que demonstrem prontidão para
     ambiente academicamente quantitativo). + 1 ensaio opcional (150
     palavras, para explicar gaps ou particularidades do histórico).
   - Transcripts: não-oficial para iniciar a aplicação; oficial
     exigido antes do 1º dia de aula (não na aplicação em si).
   - GMAT/GRE: opcional (já confirmado, sem mudança).
   - Prova de inglês: TOEFL iBT 100 / IELTS 7 / PTE 68 / DET 120 (já
     confirmado, sem mudança).
   - Entrevista: PARTE OBRIGATÓRIA do processo se o candidato for
     convidado pelo comitê de admissão (nem todo candidato é
     convidado -- o comitê pode decidir sem entrevista se já tiver
     informação suficiente). Preenchido campo 'Entrevista' (estava
     'Verificar').

CAMPOS **NÃO** alterados nesta rodada: 'Taxa Empregabilidade %' e
'Salário Médio Pós USD' (já preenchidos com 80% / $90.000 na
planilha) -- NÃO PUDE CONFIRMAR nem contradizer esses valores com uma
fonte oficial específica do MSBA. O único relatório oficial de
outcomes que encontrei acessível (fordham.edu, "2022 Employment
Outcomes") é da turma de 2022 e reporta salário médio agregado de
TODOS os programas de Specialty Masters juntos ($76.880), não
específico de Business Analytics, e não reporta uma taxa de
empregabilidade por programa -- portanto não é uma fonte adequada para
confirmar OU substituir o dado já existente (que pode vir de uma fonte
mais recente/específica não localizada nesta sessão). Ver mensagem ao
usuário para o detalhe dessa divergência.
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
    "Nome da Universidade": "Fordham",
    "Estado": "NY",
    "Cidade": "New York",
}

UPDATES = {
    "Custo/Crédito (Intl) USD": 2303,  # oficial 2026-27, entrants Fall 2023+ (era 2201, desatualizado)
    "Custo/Crédito (Residente) USD": 2303,  # privada, mesma tarifa
    "Tuition Total (Intl) USD": 69090,  # 30 x 2303 (era 66030)
    "Tuition Total (Residente) USD": 69090,
    "Fees Adicionais USD": (
        "General Fee $128/termo; Technology Fee $249/termo; Health "
        "Insurance Fee $1.890 (Fall) + $3.342 (Spring, incl. Summer); "
        "International Student Service Fee $50 (só internacionais); "
        "Maintenance of Matriculation Fee $51; Enrichment Fee (novos "
        "alunos de MS, só 1º termo) $3.939"
    ),
    "Taxa Aplicação USD": 150,  # oficial (era 130, desatualizado); isenta no Round 1
    "Deadline Aplicação": (
        "Fall 2027, Round 1: 9 out 2026 (decisão 4 dez 2026, taxa "
        "isenta + bolsa prioritária); Round 2: 15 jan 2027 (decisão 19 "
        "mar 2027, bolsa prioritária); Round 3: 19 mar 2027 (decisão "
        "14 mai 2027, último prazo p/ internacionais que precisam de "
        "visto do exterior); Round 4: rolling (domésticos e "
        "internacionais já presentes nos EUA)"
    ),
    "Rounds Aplicação": (
        "4 rounds para Fall: R1 (taxa isenta + bolsa prioritária), R2 "
        "(bolsa prioritária), R3 (último prazo p/ visto do exterior), "
        "R4 (rolling)"
    ),
    "Tipos de Bolsa": (
        "Merit-based (Specialized Master's): $5.000-$30.000, "
        "consideração automática via aplicação de admissão (sem "
        "aplicação separada). Dean's Academic & Professional/"
        "Excellence & Impact Scholarship (Specialized Master's "
        "on-site, inclui MSBA): valor fixo $15.000-$25.000/ano. "
        "Mayor's Graduate Scholarship Program: mínimo $8.000 p/ quem "
        "se qualifica em MS Business Analytics (critérios de "
        "elegibilidade não detalhados nesta fonte)."
    ),
    "Documentos Necessários": (
        "Transcripts não-oficiais para iniciar aplicação (oficial só "
        "antes do 1º dia de aula), currículo, carta de recomendação "
        "OPCIONAL (1, acadêmica ou profissional — não é obrigatória "
        "para Specialized Master's), 2 ensaios obrigatórios (Essay 1: "
        "500 palavras, motivação/objetivos; Essay 3, específico p/ "
        "MSBA: 350 palavras, preparo quantitativo) + 1 opcional (150 "
        "palavras), GMAT/GRE opcional, prova de inglês se aplicável "
        "(TOEFL iBT 100/IELTS 7/PTE 68/DET 120)"
    ),
    "Entrevista": (
        "Por convite do comitê de admissão (parte do processo "
        "holístico); o comitê pode decidir sem entrevista se já tiver "
        "informação suficiente — nem todo candidato é convidado"
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
        raise ValueError("Nenhuma linha encontrada para Fordham (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Fordham, New York, NY)")

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
