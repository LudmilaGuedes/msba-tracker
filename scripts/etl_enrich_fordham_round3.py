"""
Data enrichment (round 3): atualiza APENAS a linha da 'Fordham' (Gabelli
School of Business, New York, NY) -- MS in Business Analytics presencial/STEM.

Match por Nome + Estado + Cidade. Rodar da raiz do repositório:
    python scripts/etl_enrich_fordham_round3.py [--dry-run]

FONTES OFICIAIS (fordham.edu, acessadas em 2026-11-01 via curl com cookie
jar -- as páginas passam por um gateway CAS que devolve o conteúdo público):
- .../graduate-programs/academic-programs/ms-programs/business-analytics/
- .../business-analytics/business-analytics-career-pathways--partners/
- .../graduate-admissions/graduate-admissions-requirements/
- .../graduate-admissions/graduate-admissions-deadlines-/
- .../graduate-admissions/financing-your-graduate-education/admissions-scholarships/
- .../graduate-admissions/financing-your-graduate-education/graduate-assistantships/
- student-financial-services/tuition-and-payments/graduate-tuition/gabelli-school-of-business/

CONFIRMADO SEM MUDANÇA: tuition 2026-27 $2.303/crédito x 30 = $69.090
(privada, mesmo valor intl/residente); fees (General $128 e Technology $249
por termo, International Service $50, Enrichment MS $3.939 no 1º termo,
seguro saúde $1.890 + $3.342); taxa de aplicação $150 (isenta no Round 1);
GMAT/GRE opcional; carta opcional para Specialized Master's; deadlines Fall
2027 (9 out 2026 / 15 jan / 19 mar / rolling); STEM, 30 créditos, 12 meses
full-time, presencial no Lincoln Center.

NÃO ALTERADO POR FALTA DE FONTE: 'Taxa Empregabilidade %' (80) e 'Salário
Médio Pós USD' (90000) vêm de busca web do round 1. A Fordham não publica
relatório de empregabilidade do MSBA; a página oficial só diz que salários
iniciais base ficam entre US$75k e US$150k (90k está dentro dessa faixa).
"""

import sys

import openpyxl

from db_retention import DB_PATH, rotate_backup

SHEET = "MSBA Programs"
MATCH = ("Fordham", "NY", "New York")

UPDATES = {
    "Destaques Currículo": (
        "30 créditos, 12 meses full-time (ou part-time), campus Lincoln Center (Manhattan); "
        "STEM designated (24 meses de STEM OPT); 4 tracks: Machine Learning and AI, Advanced "
        "Analytics, Analytics Consulting, IT Ethics and Cybersecurity; core com Database "
        "Management/SQL, Machine Learning for Business, Web Analytics e Cloud Computing for "
        "Analytics; internships em empresas de NYC; opção de joint MSBA-MSIT; advising do "
        "Student Success Center"
    ),
    "Tipos de Bolsa": (
        "Merit-based (Specialized Master's): US$5.000-30.000, consideração automática com a "
        "aplicação completa -- é preciso marcar a caixa de bolsa no formulário; anunciada junto "
        "com a admissão e não é reavaliada depois do depósito. Dean's Academic Excellence & "
        "Impact Scholarship (Specialized Master's presenciais, inclui MSBA): valor fixo de "
        "US$15.000-25.000, dividido em 2 semestres de fall, recomendado aplicar até o Round 2. "
        "Graduate Assistantship (acadêmica ou administrativa): pago 100% em desconto de tuition, "
        "US$1.755 (5h/semana) a US$7.020 (20h/semana) por semestre. Mayor's Graduate "
        "Scholarship (mín. US$8.000 no MSBA) só vale para o MSBA part-time noturno -- não se "
        "aplica a quem estuda full-time com F-1"
    ),
    "Quão Realista (Intl)": (
        "Alta para bolsa de mérito: US$5.000-30.000 com consideração automática para "
        "internacionais, com prioridade nos Rounds 1 e 2; cobre só parte da tuition (~US$69k). "
        "Graduate assistantship é possível para internacionais (não exige SSN), mas é muito "
        "seletiva, distribuída semestre a semestre (mais vagas no spring) e precisa ser "
        "solicitada a cada termo"
    ),
    "TA/RA Disponível": "Sim",
    "Processo TA/RA": (
        "Graduate Assistant Application aberta a todos os alunos full-time elegíveis, a cada "
        "termo: abre ~2 meses antes do início das aulas e fecha ~1 mês depois (Fall 2026: "
        "27 jul-7 set). Professores e escritórios escolhem no pool até o fim do 2º mês do "
        "termo. Exige 9+ créditos para até 20h/semana (6+ créditos para 10h). Vagas acadêmicas "
        "(ensino, correção, pesquisa) e administrativas"
    ),
    "Graduate Assistantships": (
        "Sim -- compensação 100% em tuition remission por semestre: 5h/semana US$1.755; "
        "7,5h US$2.632; 10h US$3.510; 12,5h US$4.387; 20h US$7.020. Isento de imposto, sem "
        "W-2 e sem exigência de SSN; vagas limitadas, mais numerosas no spring"
    ),
    "Principais Empregadores": (
        "Accenture, Amazon, Bank of America, BlackRock, Citigroup, Deloitte, EY, General "
        "Electric, Goldman Sachs, Google, Intel, IBM, Meta, PepsiCo, Pfizer, Salesforce, "
        "Tesla, Verizon"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 100/IELTS 7.0/PTE 68/Duolingo 120 (notas recomendadas, não mínimas; válidas "
        "por 2 anos; a escola pode exigir entrevista ou teste de ESL)"
    ),
    "Documentos Necessários": (
        "Transcripts não oficiais para iniciar a aplicação (oficial só antes do 1º dia de "
        "aula); avaliação de credenciais WES, ECE ou outro membro NACES para diploma de fora "
        "dos EUA; currículo; carta de recomendação OPCIONAL (1, acadêmica ou profissional); "
        "Essay 1 (500 palavras, motivação e objetivos); Essay 3 específico de MSBA (350 "
        "palavras, preparo quantitativo); essay opcional (150 palavras); GMAT/GRE opcional; "
        "prova de inglês se aplicável (TOEFL iBT 100/IELTS 7/PTE 68/DET 120 recomendados); "
        "quem precisa de visto do exterior deve aplicar até o Round 3"
    ),
    "Entrevista": (
        "Obrigatória para quem é convidado (25-30 min, com staff de admissão, professores, "
        "alunos ou ex-alunos; agendar com 10 dias de antecedência); o comitê pode decidir sem "
        "entrevista se já tiver informação suficiente"
    ),
}


def main() -> None:
    dry = "--dry-run" in sys.argv
    if not dry:
        rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(header)}
    missing = [k for k in UPDATES if k not in col]
    if missing:
        raise SystemExit(f"colunas inexistentes no schema: {missing}")

    nome, estado, cidade = MATCH
    rows = [
        r for r in range(2, ws.max_row + 1)
        if ws.cell(r, col["Nome da Universidade"]).value == nome
        and ws.cell(r, col["Estado"]).value == estado
        and ws.cell(r, col["Cidade"]).value == cidade
    ]
    if len(rows) != 1:
        raise SystemExit(f"esperava 1 linha para {MATCH}, achei {len(rows)}")
    r = rows[0]

    changed = 0
    for field, new in UPDATES.items():
        cell = ws.cell(r, col[field])
        if cell.value == new:
            continue
        changed += 1
        print(f"== {field}\n   ANTES : {cell.value}\n   DEPOIS: {new}\n")
        cell.value = new
    print(f"{changed} campo(s) alterado(s) na linha {r} ({nome}).")

    if dry:
        print("dry-run: nada gravado")
        return
    wb.save(DB_PATH)
    print(f"gravado em {DB_PATH}")


if __name__ == "__main__":
    main()
