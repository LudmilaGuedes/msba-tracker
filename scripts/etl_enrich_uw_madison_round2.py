# -*- coding: utf-8 -*-
"""
Data enrichment, 2a rodada (31/out/2026): atualiza APENAS a linha da
'UW-Madison' (Wisconsin School of Business, Madison, WI) no
msba_programs_database.xlsx. Foco: MS in Business Analytics (presencial, STEM).
(Não confundir com 'Dakota State', que também fica em uma cidade chamada Madison, no SD.)

CONFIRMADO SEM ALTERAÇÃO (business.wisc.edu, out/2026):
- US$1.500/crédito x 30 = US$45.000, mesma tarifa para residente e não residente;
  segregated fees US$1.699,48 (prorrateadas com o verão).
- 10 meses, STEM; taxa de aplicação US$75 (+US$6 para internacionais).
- Bolsas: 74% da Class of 2026 recebeu bolsa, oferta média de US$9.000.
  Assistantships sem tuition remission.
- Outcomes da Class of 2025: 89% com oferta em até 6 meses; salário médio
  US$75.629 (só de quem foi empregado nos EUA); empregadores conferem.
- Prazos do ciclo com início em jul/2027: 4 nov, 13 jan, 3 mar, 29 abr (final
  para internacionais), 3 jun (final para domésticos). DET 125, IELTS 7.0,
  TOEFL 92 / 5.0.

CORRIGIDO/PREENCHIDO:
- Tamanho Turma: 62 -> 43 (Class of 2027, Class Profile atual).
- GPA Médio Admitidos: 3.49 -> 3.33 (Class of 2027).
- GPA Mínimo: a página não traz um mínimo geral; o 3.0 vale para a dispensa de
  GMAT/GRE.
- GMAT/GRE: critério de dispensa reescrito conforme a página atual (3.0 +
  Quantitative Readiness Statement + 1 indicador quantitativo; formulário
  enviado 1 semana antes do prazo). Acrescentadas as médias de GRE da Class of
  2027 (V158 / Q167); GMAT Médio continua Verificar porque a escola não divulga.
- Formato/Início/Destaques: bootcamp online de 4 semanas (jul-ago) +
  Foundations (set-dez) + Applications (jan-mai); currículo atualizado.
- Seguro Saúde: Verificar -> 1.572 (SHIP internacional 2026-27, student only, anual).
- Custo de vida (Madison): Aluguel 1.000 (quartos privativos perto do campus
  anunciados por US$975-1.200), Mercado 300, Utilities 120, Transporte 0 (UW
  Bus Pass pago pelas segregated fees), Seguro Residencial 15, Celular 50,
  WiFi 60, Outros 150.
- Custo Vida Total: fórmula do dashboard, soma mensal 1.695 x 10 meses + seguro
  1.572 = 18.522.
- Tipos de Bolsa / Quão Realista: bolsa automática, sem aplicação separada;
  aplicar até o prazo de prioridade; a escola não diz nada sobre elegibilidade de
  internacionais.
- Documentos: transcripts oficiais só depois da admissão, até 1 jul (se o
  diploma ainda não tiver saído, manda o histórico não oficial atualizado); WES
  sugerida; comprovação financeira só depois da admissão; I-20 até 5 jul.
- Entrevista: Verificar -> não mencionada (só consulta opcional).
- Testes de inglês: TOEFL válido por 2 anos (código 1846); PTE não aparece;
  regras de dispensa atualizadas.

MANTIDO COMO Verificar: Ranking US News, GMAT Médio, Taxa Aceitação,
Processo TA/RA (a WSB não descreve o processo).

Fontes (out/2026):
- business.wisc.edu/graduate/masters/business-analytics/ (+ /admissions/,
  /tuition/, /careers/, /curriculum/, Class Profile)
- uhs.wisc.edu: 2026-27 International Student Plan, Snapshot of Benefits and Premiums
- anúncios de quartos privativos em campusareahousing.wisc.edu (out/2026)

Uso: python scripts/etl_enrich_uw_madison_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("UW-Madison", "WI", "Madison")

UPDATES = {
    "Formato": (
        "Presencial: bootcamp online de 4 semanas (jul-ago), depois 2 semestres "
        "100% presenciais em Madison (Foundations set-dez; Applications jan-mai)"
    ),
    "Início": "Summer (só uma turma por ano: bootcamp online em julho; aulas presenciais a partir de setembro)",
    "Tamanho Turma": 43,
    "Destaques Currículo": (
        "Statistics & Programming Bootcamp; Advanced SQL & Data Warehousing; Data Technology; "
        "Machine Learning; Artificial Intelligence; Prescriptive Modeling & Optimization; "
        "Experiments & Causal Methods; Data Visualization; Consulting Capstone com cliente "
        "corporativo real (2o semestre); eletivas (marketing, supply chain, risk, people, health, "
        "text analytics, automação com IA); Python, R, SQL, Snowflake, Databricks, dbt, LLMs/RAG; "
        "QS 2025: #7 entre programas públicos de business analytics nos EUA"
    ),
    "Seguro Saúde (programa) USD": 1572,
    "Aluguel Privativo/mês USD": 1000,
    "Mercado/mês USD": 300,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 120,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 0,
    "Outros Gastos/mês USD": 150,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 1.695 x 10 meses + seguro 1.572.
    "Custo Vida Total (programa) USD": 18522,
    "Tipos de Bolsa": (
        "Bolsa de mérito do programa: consideração automática na aplicação (sem formulário "
        "separado), por avaliação holística; 74% da Class of 2026 recebeu, oferta média de "
        "US$9.000 (Classes 2025 e 2026); aplicar até o prazo de prioridade aumenta a chance. "
        "Assistantships/university appointments dão stipend e seguro, mas sem tuition remission. "
        "FAFSA só para americanos"
    ),
    "Quão Realista (Intl)": (
        "Alta para bolsa parcial: 74% recebem (média US$9 mil, ~20% da tuition) e a escola não "
        "publica restrição a internacionais (também não confirma explicitamente); bolsa integral ou "
        "isenção de tuition não existe"
    ),
    "GPA Mínimo": (
        "Sem mínimo geral publicado; 3.0 cumulativo exigido para pedir dispensa de GMAT/GRE "
        "(GPA internacional equivalente); auto-admissão para bacharéis da UW-Madison com GPA >= 3.4"
    ),
    "GPA Médio Admitidos": 3.33,
    "GMAT/GRE Obrigatório": (
        "Exigido (GMAT, GMAT Focus ou GRE), com dispensa seletiva: GPA 3.0 + Quantitative "
        "Readiness Statement (250 palavras) + pelo menos 1 indicador (graduação em STEM, business, "
        "economia, finanças, contabilidade, matemática, estatística, engenharia ou computação; GPA "
        ">= 3.2 ou B+ em 3 matérias quantitativas; 2+ anos de experiência analítica; CPA/CFA/PE/FRM; "
        "EA/MCAT/LSAT ou GMAT/GRE vencido; MBA Math/GMAC Business Fundamentals/HBS CORe; mestrado). "
        "Formulário de dispensa até 1 semana antes do prazo; resposta em até 10 dias úteis. "
        "GRE médio Class of 2027: Verbal 158, Quant 167"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 92 (antes de 21/jan/26) ou 5.0 (nova escala), feito até 2 anos antes do "
        "início, código 1846; IELTS mín. 7.0; Duolingo mín. 125; PTE não listado; aceita combinar "
        "notas de testes diferentes; dispensa para graduação/mestrado de 8+ semestres lecionado em "
        "inglês ou 2+ anos de trabalho em empresa só de inglês"
    ),
    "Rounds Aplicação": (
        "5 rounds (início jul/2027): 4 nov, 13 jan, 3 mar, 29 abr (final para internacionais), 3 jun "
        "(final para domésticos); aplicações entre prazos são avaliadas depois do próximo prazo; "
        "bolsa: melhor chance aplicando até o prazo de prioridade; o ciclo de jul/2028 deve repetir o "
        "calendário (~nov/2027 a abr/2028)"
    ),
    "Documentos Necessários": (
        "3 essays (máx. 600 palavras cada) + 1 opcional; 3 video essays (1 de 2-3 min na aplicação "
        "+ 2 em até 72h depois de enviar; a análise só começa com os 3); resume; transcripts não "
        "oficiais na aplicação (oficiais só depois da admissão, via Parchment de preferência, até 1 jul; "
        "se o diploma ainda não tiver saído, envia o histórico não oficial atualizado); avaliação WES "
        "sugerida para diploma estrangeiro; GMAT/GRE ou dispensa aprovada; teste de inglês; carta de "
        "recomendação opcional; taxa US$75 (+US$6 internacional; isenção só para militares dos EUA); "
        "comprovação financeira só depois da admissão; pedido de I-20 até 5 jul"
    ),
    "Entrevista": "Não exigida (não mencionada na admissão; só consulta opcional com a equipe; os 3 video essays cumprem esse papel)",
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    unknown = [h for h in UPDATES if h not in col]
    if unknown:
        raise ValueError(f"Colunas inexistentes no schema: {unknown}")

    target_rows = [
        row[0].row
        for row in ws.iter_rows(min_row=2)
        if (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        ) == MATCH
    ]
    if len(target_rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, achei {len(target_rows)}. Abortando.")
    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UW-Madison, Madison, WI)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        if cell.value != new_value:
            changes.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print(f"\nCampos atualizados ({len(changes)}) (Antes -> Depois):")
    for header, old, new in changes:
        print(f"  - {header}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
