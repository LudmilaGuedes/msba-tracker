# -*- coding: utf-8 -*-
"""
Varredura nacional (set/2026): adiciona programas presenciais e STEM de
MSBA, ou equivalentes proximos (MS em analytics numa escola de negocios),
que ainda nao estavam na base e que ACEITAM o Duolingo English Test.
Criterio: programas que nao aceitam Duolingo nao entram.

Metodo: busca por regiao a partir de uma lista de candidatos (programas
conhecidos e resultados de busca), confirmando para cada um: (1) se existe
um MS presencial (nao so online), (2) se e STEM e (3) se aceita Duolingo.
A maior parte dos dados veio de paginas oficiais ou de resumos de busca
dessas paginas. Quando o dado so apareceu em agregadores (Yocket,
Collegedunia, Shiksha etc.), ele foi para uma observacao em texto, e nao
para os campos numericos de tuition. Campos nao encontrados ficam
"Verificar". Empregabilidade, salario e custo de vida NAO foram pesquisados
nesta varredura (ficam "Verificar"), por causa do volume de programas.

EXCLUIDOS nesta varredura, e por que:
- Nao aceitam Duolingo: Univ. of Delaware, Penn State (M-BAN), Georgia
  State (e demais do sistema USG, como Kennesaw State), Cal State East Bay,
  Pepperdine (so TOEFL/IELTS/PTE).
- MSBA so online (sem visto F-1): Indiana/Kelley, Nebraska-Lincoln, Ohio
  State (SMB-A), Villanova, Oklahoma (Price), Northern Illinois,
  Montclair State (online/hibrido).
- Sem MSBA de fato (so MBA, certificado ou MS de outra area): Kentucky
  (Gatton), Auburn, Clemson, Marquette, Saint Louis Univ., Memphis, Rider,
  Colorado State, Rollins, Clarkson (MBA em BA).
- Nao foi possivel confirmar Duolingo e/ou STEM, entao nao entraram:
  UMass Boston, Saint Joseph's Univ., Providence College, Utah State
  (MDATA), Kansas State, Pace (programa correspondente nao identificado),
  Missouri S&T, SDSU.
- Removidos antes por decisao anterior (nao readicionados): NYU,
  Vanderbilt, Northwestern.

Uso: python scripts/etl_add_nationwide_sweep.py (a partir da raiz).
"""

import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
VERIFICAR = "Verificar"

# chaves curtas -> colunas do schema (mapeadas no script de carga)

NEW = [
    dict(nome="Stevens Institute of Technology", estado="NJ", cidade="Hoboken",
         escola="School of Business", tipo="Privada",
         programa="MS in Business Intelligence & Analytics",
         formato="Presencial", tt_intl=None,
         ingles="TOEFL iBT 4 (nova escala) / IELTS 6.0 / Duolingo 110 (mínimos para consideração; recomendados TOEFL 86, IELTS 7)",
         notas="Tuition ~US$42,8-44,2k por ano (agregadores)",
         fontes="stevens.edu/graduate-education/english-language-proficiency-policy"),
    dict(nome="Lehigh University", estado="PA", cidade="Bethlehem",
         escola="College of Business", tipo="Privada",
         programa="MS in Business Analytics & AI",
         formato="Presencial (matrícula full-time obrigatória no Fall e Spring)",
         duracao=11, creditos=34, cc_intl=1660, cc_res=1660, tt_intl=56440, tt_res=56440,
         gpa_medio=3.48,
         ingles="TOEFL, IELTS ou Duolingo (Duolingo 115 para pós-graduação)",
         destaques="10 a 16 meses conforme a trilha (padrão ou estendida); turma 2025: 56% internacionais, média de 1 ano de experiência",
         fontes="business.lehigh.edu MSBA about-program; tuition 2025-26 via busca"),
    dict(nome="Babson College", estado="MA", cidade="Wellesley",
         escola="F.W. Olin Graduate School of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         gmat="Opcional (test-optional para entrada em 2026)",
         ingles="TOEFL, IELTS, PTE ou Duolingo aceitos (mínimos na página de internacionais)",
         notas="Tuition anual citada ~US$56,2k (agregador)",
         fontes="babson.edu graduate admissions FAQ / MSBA page"),
    dict(nome="Brandeis University", estado="MA", cidade="Waltham",
         escola="International Business School (IBS)", tipo="Privada",
         programa="MS in Business Analytics (MSBA)", formato="Presencial",
         duracao=12, tt_intl=53300, tt_res=53300,
         bolsas="Sim",
         tipos_bolsa="Bolsa garantida de US$10.000 para todos os admitidos que começam no Fall 2026; consideração automática para bolsas de mérito adicionais",
         realista="Alta -- bolsa de US$10k garantida para admitidos, mais mérito automático",
         ingles="TOEFL, IELTS ou Duolingo (Duolingo recomendado 130+)",
         destaques="Opções de 12 ou 16 meses",
         fontes="brandeis.edu/global/academics/msba"),
    dict(nome="Boston College", estado="MA", cidade="Chestnut Hill",
         escola="Carroll School of Management", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         cc_intl=2306, cc_res=2306,
         ingles="TOEFL, IELTS, PTE ou Duolingo (Duolingo mín. 130; confirmado em 17/ago/2026)",
         fontes="bc.edu Carroll admission-insights English-Language-Test-Requirement; tuition-and-aid"),
    dict(nome="Worcester Polytechnic Institute (WPI)", estado="MA", cidade="Worcester",
         escola="WPI Business School", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         ingles="TOEFL iBT 90 / TOEFL Essentials 8.5 / IELTS 7 (sub-notas ≥6.5) / Duolingo 115",
         notas="Tuition total citada ~US$34,6k (agregador; WPI anunciou redução de tuition da pós)",
         fontes="wpi.edu business-analytics-ms / international graduate"),
    dict(nome="Suffolk University", estado="MA", cidade="Boston",
         escola="Sawyer Business School", tipo="Privada",
         programa="Master of Science in Business Analytics (MSBA)", formato="Presencial",
         duracao=9, creditos=36, cc_intl=1874, cc_res=1874, tt_intl=67464, tt_res=67464,
         ingles="Duolingo mín. 110 para admissão plena (TOEFL/IELTS também aceitos)",
         destaques="Full-time em até 9 meses ou part-time em 21 meses; duplos diplomas (MSBA/MSF, MSBA/MHA)",
         fontes="suffolk.edu MSBA page"),
    dict(nome="Clark University", estado="MA", cidade="Worcester",
         escola="School of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         duracao=18,
         ingles="TOEFL 80 / IELTS 6.0 / PTE 53 / Duolingo 115",
         notas="Tuition ~US$58,9-60,5k (agregadores); duração citada de 12 a 24 meses",
         fontes="clarku.edu business-analytics-ms; agregadores"),
    dict(nome="Hult International Business School", estado="MA", cidade="Boston",
         escola="Hult International Business School", tipo="Privada",
         programa="Master of Business Analytics & AI", formato="Presencial",
         duracao=12, tt_intl=58500, tt_res=58500,
         ingles="TOEFL, IELTS, PTE ou Duolingo aceitos",
         fontes="hult.edu/masters; agregadores"),
    dict(nome="Bryant University", estado="RI", cidade="Smithfield",
         escola="College of Business", tipo="Privada",
         programa="Master of Science in Business Analytics and AI", formato="Presencial",
         duracao=12, gmat="Não exigido",
         ingles="Duolingo mín. 115 (TOEFL/IELTS também aceitos)",
         destaques="12 ou 18 meses",
         notas="Tuition anual citada ~US$39,5k (agregador)",
         fontes="bryant.edu graduate international students; program page"),
    dict(nome="University of New Hampshire", estado="NH", cidade="Durham",
         escola="Peter T. Paul College of Business and Economics", tipo="Pública",
         programa="MS in Business Analytics", formato="Presencial",
         duracao=9, cc_intl=1075, cc_res=800,
         ingles="IELTS, TOEFL, Duolingo ou PTE aceitos",
         notas="Tuition total citada ~US$38,7k (agregador)",
         fontes="unh.edu program business-analytics; agregadores"),
]

NEW += [
    dict(nome="Hofstra University", estado="NY", cidade="Hempstead",
         escola="Frank G. Zarb School of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial", duracao=18,
         ingles="IELTS, TOEFL, Duolingo ou PTE (Duolingo 120)",
         notas="Tuition citada ~US$24,8-29,7k por ano (agregadores)",
         fontes="hofstra.edu info-systems-business-analytics/ms-business-analytics"),
    dict(nome="Fairfield University", estado="CT", cidade="Fairfield",
         escola="Charles F. Dolan School of Business", tipo="Privada",
         programa="MS in AI and Business Analytics (MSAIBA; antigo MSBA, renomeado em abr/2026)",
         formato="Presencial (também existe online)", duracao=12,
         ingles="Duolingo mín. 105 (TOEFL/IELTS também aceitos)",
         fontes="fairfield.edu Dolan graduate admission; news abr/2026"),
    dict(nome="Quinnipiac University", estado="CT", cidade="Hamden",
         escola="School of Business", tipo="Privada",
         programa="MS in Applied AI and Business Analytics", formato="Presencial (também existe online)",
         inicio="Fall ou Spring", cc_intl=1165, cc_res=1165, gpa_min="3.0",
         ingles="TOEFL 85+ / IELTS 6.5+ / Duolingo 105+",
         prereq="Recomendado: cursos introdutórios de computação e estatística",
         fontes="qu.edu graduate international; catalog.qu.edu; tuition 2026-27 via busca"),
    dict(nome="St. John's University", estado="NY", cidade="Queens",
         escola="Peter J. Tobin College of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         ingles="TOEFL, IELTS ou Duolingo",
         fontes="stjohns.edu business-analytics-master-science"),
    dict(nome="Tulane University", estado="LA", cidade="New Orleans",
         escola="A.B. Freeman School of Business", tipo="Privada",
         programa="Master of Business Analytics and AI (MANA)", formato="Presencial",
         duracao=10, creditos=36, gmat="Opcional (pode reforçar a candidatura e as chances de bolsa)",
         ingles="TOEFL, IELTS, PTE ou Duolingo",
         deadline="Round 1: 1 nov (decisão 15 jan); Round 2: 1 fev (decisão 15 mar); Round 3: 1 mai (decisão 1 jun); rolling depois",
         taxa=125, docs="Formulário online; resume; 1 essay; 1 carta de recomendação; transcripts",
         destaques="2 semestres, ou 3 para incluir estágio",
         fontes="freeman.tulane.edu/graduate/master-business-analytics/admissions"),
    dict(nome="Emory University", estado="GA", cidade="Atlanta",
         escola="Goizueta Business School", tipo="Privada",
         programa="MS in Business Analytics (MSBA)", formato="Presencial",
         duracao=10, tt_intl=82794, tt_res=82794,
         ingles="TOEFL 100 / IELTS 7.0 / PTE 68 / Duolingo aceito",
         destaques="Opções de 10 meses (US$82.794) ou 16 meses (US$90.760), tuition e fees 2026-27",
         fontes="goizueta.emory.edu/msba/admissions/tuition; application"),
    dict(nome="Duke University", estado="NC", cidade="Durham",
         escola="Fuqua School of Business", tipo="Privada",
         programa="Master of Quantitative Management: Business Analytics (MQM:BA)", formato="Presencial",
         duracao=10, inicio="Julho", tt_intl=85600, tt_res=85600,
         ingles="Duolingo (preferido pela Fuqua), TOEFL ou IELTS",
         rounds="Taxa de aplicação ISENTA para quem envia nota do Duolingo",
         notas="Custo total de participação estimado em US$123.452 (turma 2026)",
         fontes="fuqua.duke.edu/programs/mqm-business-analytics (application-instructions, tuition-costs)"),
    dict(nome="Iowa State University", estado="IA", cidade="Ames",
         escola="Ivy College of Business", tipo="Pública",
         programa="MS in Business Analytics", formato="Presencial",
         duracao=15, creditos=30, cc_intl=1848, cc_res=782,
         ingles="Duolingo aceito (TOEFL/IELTS também)",
         notas="Tarifas por crédito da Ivy (+ fees): residente 782, não-residente 1.746, internacional 1.848",
         fontes="grad-college.iastate.edu programs business_analytics; ivybusiness.iastate.edu"),
    dict(nome="UC Riverside", estado="CA", cidade="Riverside",
         escola="School of Business", tipo="Pública",
         programa="Master of Science in Business Analytics (MSBA)", formato="Presencial",
         duracao=9, creditos=48,
         ingles="Duolingo 100 (TOEFL e IELTS também aceitos)",
         destaques="Oferecido com o Departamento de Estatística; 48 unidades (quarters); 3 quarters ou 4 para incluir estágio no verão",
         fontes="business.ucr.edu/msba"),
]

NEW += [
    dict(nome="UNC Charlotte", estado="NC", cidade="Charlotte",
         escola="Belk College of Business / School of Data Science", tipo="Pública",
         programa="MS in Data Science and Business Analytics (DSBA)", formato="Presencial",
         duracao=18, creditos=33,
         ingles="TOEFL, IELTS ou Duolingo (Duolingo mín. 115)",
         notas="Custo do 1º ano citado ~US$25-33k (agregadores)",
         fontes="dsba.charlotte.edu admissions international-applicants"),
    dict(nome="University of Tampa", estado="FL", cidade="Tampa",
         escola="Sykes College of Business", tipo="Privada",
         programa="M.S. in Business Analytics", formato="Presencial (também online)",
         cc_intl=758, cc_res=758,
         ingles="IELTS, TOEFL ou Duolingo",
         fontes="ut.edu graduate-degrees ms-in-business-analytics; agregadores"),
    dict(nome="Chapman University", estado="CA", cidade="Orange",
         escola="Argyros College of Business and Economics", tipo="Privada",
         programa="MS in Business Analytics (MSBA)", formato="Presencial",
         duracao=10, cc_intl=2135, cc_res=2135,
         ingles="Duolingo (DET) aceito, além de TOEFL/IELTS",
         destaques="Programa de 10 meses reformulado em 2025, com IA integrada ao currículo",
         fontes="chapman.edu MSBA page; news.chapman.edu 2025/07/02"),
    dict(nome="University of St. Thomas", estado="MN", cidade="Minneapolis",
         escola="Opus College of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial (confirmar horários/formato com o programa)",
         creditos=30, cc_intl=1489, cc_res=1489, tt_intl=44670, tt_res=44670,
         bolsas="Sim", tipos_bolsa="Dean's Scholarship de até US$9.000 para ingressantes",
         ingles="TOEFL, IELTS ou Duolingo",
         fontes="business.stthomas.edu ms-business-analytics admissions-fees"),
    dict(nome="Case Western Reserve University", estado="OH", cidade="Cleveland",
         escola="Weatherhead School of Management", tipo="Privada",
         programa="Master of Business Analytics and Intelligence (MBAI)", formato="Presencial",
         creditos=30, cc_intl=1725, cc_res=1725, tt_intl=51750, tt_res=51750,
         ingles="TOEFL, IELTS, Duolingo ou PTE (aceitos pela Weatherhead; confirmar com o programa, pois a School of Graduate Studies da CWRU tem regra diferente para outros cursos)",
         destaques="Analytics, machine learning e IA aplicados a operações e marketing",
         fontes="case.edu/weatherhead MBAI tuition-information (entrada Fall 2026)"),
    dict(nome="University of Dayton", estado="OH", cidade="Dayton",
         escola="School of Business Administration", tipo="Privada",
         programa="Master of Business Analytics (MBAN)", formato="Presencial (também online)",
         duracao=12,
         ingles="IELTS, TOEFL, Duolingo ou PTE (Duolingo 110)",
         notas="Tuition citada ~US$29-37k por ano (agregadores)",
         fontes="udayton.edu business-analytics-mban; agregadores"),
    dict(nome="Gonzaga University", estado="WA", cidade="Spokane",
         escola="School of Business Administration", tipo="Privada",
         programa="M.S. in Business Analytics (MSBA)", formato="Presencial",
         duracao=9, creditos=30, cc_intl=1058, cc_res=1058, tt_intl=31740, tt_res=31740,
         ingles="TOEFL, IELTS, iTEP, Duolingo ou PTE",
         destaques="30 créditos em dois semestres; credenciamento AACSB; trilha de negócios e opção acelerada",
         fontes="gonzaga.edu school-of-business-administration graduate msba (tuition-financial-aid)"),
    dict(nome="Loyola University Chicago", estado="IL", cidade="Chicago",
         escola="Quinlan School of Business", tipo="Privada",
         programa="MS in Business Data Analytics", formato="Presencial",
         duracao=12, cc_intl=1760, cc_res=1760,
         ingles="TOEFL 90 / IELTS 6.5 / PTE 61 / Duolingo 115 (mínimos da pós da Quinlan)",
         fontes="luc.edu/quinlan STEM business programs; msinbusinessanalytics"),
]

NEW += [
    dict(nome="University of New Haven", estado="CT", cidade="West Haven",
         escola="Pompea College of Business", tipo="Privada",
         programa="M.S. in Business Analytics (renomeado para AI in Business and Analytics)", formato="Presencial",
         ingles="IELTS, TOEFL, Duolingo ou PTE",
         notas="Tuition do 1º ano citada ~US$21,8k (agregador)",
         fontes="newhaven.edu business graduate-programs business-analytics; catalog.newhaven.edu"),
    dict(nome="UMass Lowell", estado="MA", cidade="Lowell",
         escola="Manning School of Business", tipo="Pública",
         programa="Master of Science in Business Analytics", formato="Presencial (também online/híbrido)",
         duracao=12, creditos=30,
         ingles="Duolingo aceito (mín. citado 135 por agregador -- confirmar); TOEFL/IELTS também",
         notas="Custo do 1º ano citado ~US$28,4k (agregador); 12 a 24 meses",
         fontes="uml.edu/msb operations-info-systems msba; agregadores"),
    dict(nome="Adelphi University", estado="NY", cidade="Garden City",
         escola="Robert B. Willumstad School of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial", creditos=30,
         bolsas="Sim", tipos_bolsa="Graduate Merit Scholarships de US$3.000 a US$10.000 por ano, renováveis",
         ingles="TOEFL 80 / IELTS 6.5 / Duolingo 110",
         notas="Tuition citada de US$25,6k a US$44,7k por ano (agregadores divergem)",
         fontes="adelphi.edu program graduate business-analytics; agregadores"),
    dict(nome="LIU Post (Long Island University)", estado="NY", cidade="Brookville",
         escola="School of Business", tipo="Privada",
         programa="MS in Data Analytics & Strategic Business Intelligence", formato="Presencial (aulas à noite e fins de semana)",
         creditos=30, ingles="Duolingo 100 (TOEFL/IELTS também)",
         destaques="Mistura de data science, TI e negócios; full-time ou part-time",
         fontes="liu.edu/post MS-data-analytics-strategic-business-intelligence"),
    dict(nome="Xavier University", estado="OH", cidade="Cincinnati",
         escola="Williams College of Business", tipo="Privada",
         programa="Master of Science in Business Analytics", formato="Presencial",
         cc_intl=966, cc_res=966,
         ingles="Duolingo 105+ (TOEFL/IELTS também)",
         fontes="xavier.edu master-science-business-analytics; tuition-fees graduate-current 2026-27"),
    dict(nome="Virginia Commonwealth University (VCU)", estado="VA", cidade="Richmond",
         escola="School of Business", tipo="Pública",
         programa="Master of Decision Analytics (MDA)", formato="Presencial (aulas à noite; há também formato de fim de semana)",
         ingles="IELTS, TOEFL, PTE ou Duolingo",
         destaques="Concentrações: accounting analytics, data science in business, financial analytics, marketing analytics, production & operations analytics",
         fontes="business.vcu.edu master-of-decision-analytics-evening-format; bulletin.vcu.edu"),
    dict(nome="University of San Diego", estado="CA", cidade="San Diego",
         escola="Knauss School of Business", tipo="Privada",
         programa="MS in Business Analytics", formato="Presencial",
         duracao=10, creditos=36, cc_intl=1725, cc_res=1725, tt_intl=62100, tt_res=62100,
         ingles="IELTS, TOEFL ou Duolingo (Duolingo 120)",
         destaques="Programa presencial de 10 meses; turmas com 0 a 6,5 anos de experiência",
         fontes="sandiego.edu business graduate business-analytics; agregadores"),
]

KEY_TO_HEADER = {
    "nome": "Nome da Universidade", "estado": "Estado", "cidade": "Cidade",
    "escola": "Business School", "tipo": "Tipo (Pub/Priv)",
    "programa": "Nome Oficial do Programa", "formato": "Formato",
    "duracao": "Duração (meses)", "inicio": "Início", "destaques": "Destaques Currículo",
    "creditos": "Total Créditos", "cc_intl": "Custo/Crédito (Intl) USD",
    "cc_res": "Custo/Crédito (Residente) USD", "tt_intl": "Tuition Total (Intl) USD",
    "tt_res": "Tuition Total (Residente) USD", "bolsas": "Bolsas Disponíveis",
    "tipos_bolsa": "Tipos de Bolsa", "realista": "Quão Realista (Intl)",
    "gpa_min": "GPA Mínimo", "gpa_medio": "GPA Médio Admitidos",
    "gmat": "GMAT/GRE Obrigatório", "prereq": "Matérias Pré-Requisito",
    "ingles": "Testes Inglês Aceitos", "deadline": "Deadline Aplicação",
    "taxa": "Taxa Aplicação USD", "rounds": "Rounds Aplicação",
    "docs": "Documentos Necessários",
}
IGNORED_KEYS = {"notas", "fontes"}


def build_row(d, headers):
    row = {h: VERIFICAR for h in headers}
    row["STEM Designated"] = "Sim"
    for k, v in d.items():
        if k in IGNORED_KEYS or v is None:
            continue
        row[KEY_TO_HEADER[k]] = v
    if d.get("notas"):
        base = row["Destaques Currículo"]
        obs = "Obs.: " + d["notas"]
        row["Destaques Currículo"] = obs if base == VERIFICAR else f"{base}; {obs}"
    return row


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}
    unknown = set(KEY_TO_HEADER.values()) - set(headers)
    if unknown:
        raise ValueError(f"Colunas inexistentes no schema: {unknown}")

    existing = {
        (r[col["Nome da Universidade"] - 1], r[col["Estado"] - 1], r[col["Cidade"] - 1])
        for r in ws.iter_rows(min_row=2, values_only=True)
    }

    ref = ws.cell(row=2, column=1)
    font = Font(name=ref.font.name, sz=ref.font.sz)
    side = Side(style="thin", color="00D0D0D0")
    border = Border(left=side, right=side, top=side, bottom=side)
    align = Alignment(vertical="top", wrap_text=True)
    fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")

    added, skipped = [], []
    for d in NEW:
        key = (d["nome"], d["estado"], d["cidade"])
        if key in existing:
            skipped.append(key)
            continue
        row = build_row(d, headers)
        r = ws.max_row + 1
        for h, c in col.items():
            cell = ws.cell(row=r, column=c, value=row[h])
            cell.font, cell.border, cell.alignment = font, border, align
            if row[h] == VERIFICAR:
                cell.fill = fill
        added.append((r, d))
        existing.add(key)

    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(headers)).column_letter}{ws.max_row}"

    total = ws.max_row - 1
    for srow in wb[SHEET_SUMMARY].iter_rows(min_row=2):
        if srow[0].value in ("Total de Programas", "Programas Presenciais STEM"):
            srow[1].value = total
        if srow[0].value == "Data de Compilação":
            stamp = datetime.now().strftime("%d/%m/%Y")
            srow[1].value = f"{srow[1].value} + {len(added)} programas (varredura nacional) em {stamp}"

    wb.save(DB_PATH)
    print()
    print(f"{len(added)} programas adicionados (total na base: {total}):")
    for r, d in added:
        print(f"  linha {r}: {d['nome']} ({d['cidade']}, {d['estado']}) -- {d['programa']}")
    if skipped:
        print()
        print(f"Já existiam (pulados): {skipped}")


if __name__ == "__main__":
    main()
