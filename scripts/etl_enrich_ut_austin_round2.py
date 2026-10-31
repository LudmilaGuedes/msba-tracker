# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, out/2026): atualiza APENAS a linha da 'UT Austin'
(McCombs School of Business, Austin, TX) no msba_programs_database.xlsx.

Foco: MS Business Analytics and AI (MSBAi), presencial em Austin, STEM.
A rodada 1 (etl_enrich_ut_austin.py, set/2026) ja tinha auditado a maior
parte da linha; esta rodada reconfere tudo nas paginas oficiais e corrige o
que mudou ou estava errado.

CONFIRMADO SEM ALTERACAO: nome do programa; STEM; presencial, 10 meses, 36
creditos; tuition total US$55.600 (residente) / US$59.700 (nao-residente,
que e a tarifa de internacionais), por semestre: residente 10.600 Summer +
21.500 Fall + 21.500 Spring, nao-residente 11.900 + 22.900 + 22.900, mais a
taxa de pre-matricula de US$2.000; turma 64, GPA medio 3.55, GMAT medio 655
(Class of 2027); GMAT/GRE opcional; 4 rounds (Round 2, 12/jan/2027, e o
ultimo para internacionais); taxa de aplicacao US$90 (intl) / US$65
(domestico); 2 cartas de recomendacao; Kira (ate 30 min) + entrevista com
selecionados; TA/RA nao disponivel; bolsas de merito limitadas e
automaticas; empregabilidade e salario (relatorio Class of 2025, ainda o
mais recente); QS 12 (2026; a edicao 2027 nao foi encontrada).

CORRIGIDO/PREENCHIDO:
- Fees Adicionais: estava 2000, mas os US$2.000 sao a taxa de pre-matricula,
  que ja esta somada na tuition total (contava duas vezes). O que fica de
  fora da tuition, segundo a pagina oficial, e o seguro saude internacional
  e as taxas do ISSS (valor do ISSS nao publicado).
- Seguro Saude (programa): estava Verificar. UT SHIP 2026-27 (ISSS):
  Fall US$1.204 + Spring US$1.505 + Summer US$903 = US$3.612. Estimativa:
  a turma comeca no Summer e os precos de 2027-28 ainda nao sairam.
- Custo Vida Total (programa): 27.640 -> 30.252, recalculado com a mesma
  formula do dashboard (gastos mensais 2.664 x 10 meses + seguro 3.612).
- Ranking US News: estava Verificar. A pagina do programa cita o US News #9
  em "Best Business School Business Analytics Dept." (ranking de
  especialidade, nao de MS especificamente).
- Testes Ingles Aceitos: a Graduate School agora lista DET 115 como minimo
  oficial (junto com TOEFL 79 e IELTS 6.5). A isencao estava errada: vale
  so para bacharelado nos EUA ou em pais da lista de qualifying countries;
  mestrado NAO isenta (a pagina diz isso explicitamente).
- Rounds Aplicacao: a isencao da taxa de aplicacao nao vale para
  internacionais (Graduate School: "Fee waivers are not available for
  international citizens").
- Tipos de Bolsa: acrescentado que o valor das bolsas nao e publicado, que
  a pagina nao diz se internacionais sao elegiveis, e que ha emprestimos
  para internacionais (alguns sem co-signer americano).
- Documentos Necessarios: a pagina pede historico profissional (employment
  history), nao resume; 3 ensaios de ate 500 palavras (2 obrigatorios + 1
  opcional); transcripts de todas as faculdades cursadas.

MANTIDO COMO Verificar: Custo/Credito (Intl e Residente), porque o McCombs
cobra por semestre, sem tarifa por credito publicada; Taxa de Aceitacao %
(nao divulgada).

Fontes (out/2026):
- mccombs.utexas.edu/.../ms-business-analytics-and-artificial-intelligence/on-campus/admissions/tuition-financial-aid/
- mccombs.utexas.edu/.../on-campus/admissions/application-process/
- mccombs.utexas.edu/.../on-campus/admissions/faq
- mccombs.utexas.edu/.../ms-business-analytics-and-artificial-intelligence/on-campus/ (rankings)
- mccombs.utexas.edu/.../ms-business-analytics-on-campus/admissions/class-stats/
- graduate.utexas.edu/admissions/apply/international (TOEFL 79, IELTS 6.5,
  DET 115, isencoes)
- gradschool.utexas.edu/how-to-apply (fee waiver nao vale para internacionais)
- global.utexas.edu/isss/advising-services/insurance/costs (UT SHIP 2026-27)

Uso: python scripts/etl_enrich_ut_austin_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
VERIFICAR = "Verificar"

MATCH = ("UT Austin", "TX", "Austin")

UPDATES = {
    "Fees Adicionais USD": (
        "Taxa de pré-matrícula de US$2.000 já incluída na tuition total; à parte: "
        "seguro saúde internacional (ver coluna de seguro) e taxas do ISSS (valor não "
        "publicado), além de livros, software e laptop"
    ),
    "Seguro Saúde (programa) USD": 3612,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 2.664 x 10 meses + seguro 3.612.
    "Custo Vida Total (programa) USD": 30252,
    "Ranking US News": (
        "9 (Best Business School Business Analytics Dept., ranking de especialidade "
        "da US News citado pelo McCombs)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL, IELTS ou Duolingo (DET); mínimos da Graduate School: TOEFL iBT 79 / "
        "IELTS Academic 6.5 / DET 115; o programa espera TOEFL 100 ou IELTS 7.0 "
        "(sem referência publicada para DET); isenção só para bacharelado nos EUA ou "
        "em país da lista de qualifying countries -- mestrado não isenta"
    ),
    "Rounds Aplicação": (
        "4 rounds; Round 2 (12 jan 2027) é o último para internacionais; taxa US$90 "
        "intl / US$65 doméstico; isenção da taxa não disponível para internacionais"
    ),
    "Tipos de Bolsa": (
        "Merit (número limitado; consideração automática, sem inscrição separada; "
        "valores não publicados e a página não diz se internacionais são elegíveis); "
        "sem TA/RA; empréstimos para internacionais possíveis (alguns sem co-signer "
        "americano)"
    ),
    "Documentos Necessários": (
        "Aplicação McCombs (Step 1) + aplicação da Graduate School (Step 2); "
        "transcripts oficiais de todas as faculdades cursadas; histórico profissional "
        "(employment history); 3 ensaios de até 500 palavras (2 obrigatórios + 1 "
        "opcional); 2 cartas de recomendação (pelo menos 1 acadêmica); GMAT/GRE "
        "opcional; TOEFL/IELTS/DET (internacionais); taxa de US$90 (intl) / US$65 "
        "(doméstico); video assessment Kira (Step 3, até 30 min)"
    ),
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
    print(f"Linha alvo confirmada: {row_idx} (UT Austin, Austin, TX)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        changes.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print("\nCampos atualizados (Antes -> Depois):")
    for header, old, new in changes:
        print(f"  - {header}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
