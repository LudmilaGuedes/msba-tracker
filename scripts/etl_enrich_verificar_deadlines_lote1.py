# -*- coding: utf-8 -*-
"""
Data enrichment em LOTE: pesquisa a fundo (Deep Research via subagentes
paralelos, fontes oficiais .edu) de TODAS as 13 linhas que estavam com
"Deadline Aplicação" = "Verificar" no banco, para confirmar as datas reais
de aplicacao.

Pedido da usuaria: "faca uma pesquisa a fundo de todos os programas que
estao com verificar [no deadline], tenha certeza das datas de aplicacao
deles".

Resultado: 10 de 13 programas tiveram deadline confirmado (ou fortemente
inferido) em fonte oficial e foram atualizados. 3 programas permanecem
"Verificar" porque a pesquisa oficial nao encontrou uma data especifica
publicada mesmo apos multiplas buscas -- nesses casos o campo "Rounds
Aplicacao" foi enriquecido com o contexto encontrado (modelo de admissao,
regras confirmadas, contato direto recomendado) em vez de inventar uma
data:
  - UCF (Professional): trilha "Professional MSM" nao tem calendario
    proprio publicado (distinto do MSBA padrao da UCF, que ja tem deadline
    confirmado em outra linha). Pagina oficial remete a contato direto.
  - Southern Connecticut State University (SCSU): admissao rolling
    confirmada oficialmente, mas nenhuma data especifica do STEM MSBA
    encontrada.
  - Texas A&M University-Texarkana: catalogo oficial diz "program
    specific" sem listar a data; unica regra confirmada e generica (docs
    de internacionais 45 dias antes do 1o dia de aula).

Tambem corrigidos/preenchidos 3 valores de "Início" que a mesma pesquisa
esclareceu:
  - UC San Diego (Rady): programa inicia no VERAO, nao no outono (estava
    "Fall" por engano -- corrigido para "Summer").
  - Santa Clara University: confirmado Fall, meados de setembro.
  - Rensselaer Polytechnic Institute (RPI): confirmado Fall (mes exato
    nao encontrado com uma fonte oficial especifica de MSBA).
  - Southern Connecticut State University (SCSU): confirmado Fall e
    Spring (linguagem oficial geral da pos-graduacao).

Fontes por escola (todas .edu, verificadas por WebFetch/WebSearch):
- UC Davis (GSM): gsm.ucdavis.edu/master-science-business-analytics/admissions
- UC San Diego (Rady): rady.ucsd.edu/programs/ms-business-analytics/admissions/
- Michigan State (Broad): broad.msu.edu/masters/business-data-science-analytics/admissions/
  -- CONFIANCA MODERADA: pagina renderizada via JavaScript, datas vieram
  de snippets indexados pelo buscador do proprio dominio, nao de leitura
  direta do HTML confirmada. Registrado com essa ressalva no proprio
  campo "Rounds Aplicacao".
- Binghamton (SUNY): gograd.binghamton.edu/portal/Program_Information
- UConn: analytics.business.uconn.edu/ (secao "Important Dates")
  -- Fall 2027 ainda nao publicado; usado o prazo do ciclo Fall 2026
  (29/jul/2026) como padrao anual aproximado, marcado com "~".
- GWU: business.gwu.edu/admissions/masters-application-guide
- U of Central Oklahoma: uco.edu/programs/jcgs/cob/isom/business-analytics
- Illinois Tech: iit.edu/admissions-aid/graduate-admission/admission-dates-and-deadlines
  -- calendario geral de pos-graduacao; nao ha pagina especifica do
  Stuart School/MSBA com calendario proprio.
- Santa Clara: scu.edu/business/ms/ms-business-analytics/ -- Fall 2026
  confirmado; Fall 2027 ainda nao publicado (mesma ressalva "~").
- RPI: gradadmissions.rpi.edu/deadlines-and-requirements/lally-school-management-application-deadlines

Uso: python etl_enrich_verificar_deadlines_lote1.py (executar de dentro
de scripts/, mesma convencao dos demais etl_enrich_*.py).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

# Cada entrada: (Nome da Universidade, Estado, Cidade) -> {campo: novo_valor}
SCHOOL_UPDATES = {
    ("UC Davis", "CA", "Davis/San Francisco"): {
        "Deadline Aplicação": (
            "Round 1: 15/out/2026; Round 2: 15/nov/2026; Round 3: 15/jan/2027; "
            "Round 4: 15/mar/2027; Round 5: 15/mai/2027"
        ),
        "Rounds Aplicação": (
            "5 rounds (ciclo de ingresso ago/2027); Round 4 (15/mar/2027) é o "
            "prazo prioritário recomendado para internacionais, mas inscrições "
            "em qualquer round até o R5 são aceitas para todos os candidatos"
        ),
    },
    ("UC San Diego", "CA", "La Jolla"): {
        "Início": "Summer",
        "Deadline Aplicação": (
            "Round 1: 1/out/2026; Round 2: 15/nov/2026; Round 3: 15/jan/2027; "
            "Round 4: 1/abr/2027; Round 5: 15/mai/2027"
        ),
        "Rounds Aplicação": (
            "5 rounds (ciclo de ingresso Summer 2027 -- programa da Rady inicia "
            "no verão, não no outono); Round 4 (1/abr/2027) é o prazo final "
            "para internacionais -- Round 5 (15/mai/2027) é exclusivo para "
            "candidatos domésticos"
        ),
    },
    ("Michigan State", "MI", "East Lansing"): {
        "Deadline Aplicação": "Fall 2027: 1/abr/2027; Spring 2027: 1/out/2026",
        "Rounds Aplicação": (
            "Admissão por prioridade/rolling (programa 100% online, "
            "assíncrono); sem rounds numerados. Após os prazos acima, "
            "inscrições continuam sendo aceitas conforme disponibilidade de "
            "vagas. ATENÇÃO -- confiança moderada: dados extraídos de "
            "resultados de busca indexados do domínio oficial broad.msu.edu; "
            "a página usa JavaScript e não foi possível confirmar por leitura "
            "direta do HTML em 3 tentativas. Recomenda-se checagem manual "
            "antes de decidir."
        ),
    },
    ("Binghamton SUNY", "NY", "Binghamton"): {
        "Deadline Aplicação": "Fellowship Clifford D. Clark: 15/jan; recomendado (consideração plena): 15/mar",
        "Rounds Aplicação": (
            "Admissão contínua (rolling), somente Fall (único intake); não é "
            "sistema de rounds. Datas confirmadas para o ciclo Fall 2026 "
            "(gograd.binghamton.edu) -- Fall 2027 deve seguir padrão anual "
            "semelhante, mas ainda não republicado; confirmar mais perto da "
            "data. Sem prazo internacional separado publicado; recomendação "
            "geral é aplicar cedo dado o processo de visto."
        ),
    },
    ("UConn", "CT", "Hartford/Stamford"): {
        "Deadline Aplicação": (
            "Fall: ~fim de jul/2027 (padrão anual, ainda não publicado "
            "oficialmente); Spring 2027: 16/dez/2026 (confirmado)"
        ),
        "Rounds Aplicação": (
            "Datas fixas por semestre (Fall/Spring/Summer), sem rounds "
            "numerados. Spring 2027 confirmado oficialmente (início das "
            "aulas 19/jan/2027). Fall 2027 ainda não publicado -- usado como "
            "referência o prazo do ciclo Fall 2026 (29/jul/2026) por padrão "
            "anual, não é uma data oficial confirmada para 2027. Candidatos "
            "que precisam de visto só podem entrar via Fall ou Spring (não "
            "Summer); processo de visto leva ~3 meses após admissão formal, "
            "então aplicar perto do prazo é arriscado para quem precisa de "
            "visto."
        ),
    },
    ("GWU", "DC", "Washington"): {
        "Deadline Aplicação": "1/nov/2026",
        "Rounds Aplicação": (
            "Deadline única por ciclo (não é sistema de rounds), somente "
            "Fall; ciclo Fall 2027 confirmado na página oficial de "
            "admissões. Internacionais que precisam de visto são apenas "
            "recomendados a aplicar cedo (mesmo prazo que domésticos, sem "
            "data separada)."
        ),
    },
    ("U of Central Oklahoma", "OK", "Edmond"): {
        "Deadline Aplicação": "Geral: 15/abr; Doméstico: 15/mai; Internacional: 15/jun",
        "Rounds Aplicação": (
            "Datas fixas escalonadas (não é sistema de rounds). Confirmado "
            "oficialmente: Fall Admission (geral) 15/abr; Domestic Admission "
            "15/mai; Space Available 15/mai; International Admission 15/jun "
            "-- atenção: o prazo internacional é o MAIS TARDIO dos quatro, ao "
            "contrário do padrão comum de outras escolas (confirmado "
            "exatamente assim na página oficial)."
        ),
    },
    ("Illinois Tech", "IL", "Chicago"): {
        "Deadline Aplicação": (
            "Round 1: 31/jan/2027 (priority); Round 2: 15/abr/2027 (priority); "
            "Round 3: 15/jun/2027 (internacional fora dos EUA) / 15/jul/2027 "
            "(doméstico e internacional já nos EUA)"
        ),
        "Rounds Aplicação": (
            "3 prazos (2 priority + 1 regular); sem taxa de aplicação "
            "(política permanente, já confirmada). Deadline de documentos "
            "para I-20 (visto): 30/jun/2027 (fora dos EUA) / 31/jul/2027 (já "
            "nos EUA). Datas confirmadas no calendário geral de pós-graduação "
            "da Illinois Tech -- não há página específica do Stuart "
            "School/MSBA com calendário próprio; assume-se que o MSBA segue o "
            "calendário geral da universidade."
        ),
    },
    ("Santa Clara University", "CA", "Santa Clara"): {
        "Início": "Fall (meados de setembro)",
        "Deadline Aplicação": "Prioritário: 1/abr; Final: 1/jun",
        "Rounds Aplicação": (
            "Prioritário + final (não é sistema de rounds numerados); admite "
            "Fall e Spring. Confirmado para o ciclo Fall 2026 -- Fall 2027 "
            "ainda não publicado, mas deve seguir o mesmo padrão anual. "
            "Internacionais (F-1) são recomendados a se inscrever até o prazo "
            "PRIORITÁRIO para dar tempo ao processamento do I-20; após o "
            "prazo final, aplicações de F-1 não são mais aceitas (domésticos "
            "podem seguir aplicando conforme vaga disponível)."
        ),
    },
    ("Rensselaer Polytechnic Institute (RPI)", "NY", "Troy"): {
        "Início": "Fall",
        "Deadline Aplicação": (
            "Round 1: 1/nov/2026; Round 2: 15/dez/2026; Round 3: 1/fev/2027; "
            "Round 4: 1/mar/2027; Round 5: 15/abr/2027"
        ),
        "Rounds Aplicação": (
            "5 rounds (ciclo Fall 2027) + admissão rolling após 15/abr/2027; "
            "Round 1 não cobra taxa de aplicação. Internacionais são apenas "
            "recomendados a aplicar nos primeiros rounds (R1-R2) por causa do "
            "processamento de visto/I-20, sem prazo numerado separado. "
            "Confirmado na página oficial de admissões da pós-graduação da "
            "RPI -- apresentada como calendário geral do Lally School "
            "(MS/MBA), não confirmada como exclusiva do MSBA. Mês exato de "
            "início do programa (Fall) não confirmado com fonte específica."
        ),
    },
    ("Southern Connecticut State University (SCSU)", "CT", "New Haven"): {
        "Início": "Fall, Spring",
        # Deadline Aplicação permanece "Verificar" -- ver nota no topo do arquivo.
        "Rounds Aplicação": (
            "Verificar -- admissão contínua (rolling) confirmada oficialmente "
            "(southernct.edu/grad), mas nenhuma data específica encontrada "
            "para o STEM MSBA; a página do programa não lista prazo. Um "
            "prazo geral de 1/mai apareceu em resultados de busca mas não "
            "pôde ser confirmado como aplicável a este programa "
            "especificamente -- não usado. Sem distinção de prazo "
            "internacional x doméstico publicada. Recomenda-se contato "
            "direto: graddocs@southernct.edu."
        ),
    },
    ("Texas A&M University-Texarkana (A&M-Texarkana)", "TX", "Texarkana"): {
        # Deadline Aplicação permanece "Verificar" -- ver nota no topo do arquivo.
        "Rounds Aplicação": (
            "Verificar -- não há rounds nem prazo fixo publicado (catálogo "
            "oficial diz prazos 'program specific', sem data listada para "
            "este programa). Regra confirmada para internacionais: "
            "documentos devem chegar ao Office of International Studies até "
            "45 dias antes do primeiro dia de aula do semestre "
            "(catalog.tamut.edu). Taxa de aplicação: $30 doméstico / $50 "
            "internacional. Recomenda-se contato direto: "
            "internationalprograms@tamut.edu ou graduateadmissions@tamut.edu "
            "para a data exata do próximo ciclo."
        ),
    },
}

# UCF (Professional), Orlando/FL: nenhuma atualização de data -- a pesquisa
# não encontrou calendário publicado para a trilha "Professional" (distinta
# do MSBA padrão da UCF). Só o campo "Rounds Aplicação" é enriquecido com o
# contexto encontrado, para não deixar a linha sem nenhum registro da
# tentativa de pesquisa.
SCHOOL_UPDATES[("UCF (Professional)", "FL", "Orlando")] = {
    "Rounds Aplicação": (
        "Verificar -- não foi encontrado prazo específico publicado para a "
        "trilha 'Professional' (Professional MSM, trilha Business "
        "Analytics), distinta do MSBA padrão da UCF (que tem prazo próprio "
        "já registrado nesta base, na linha 'UCF'). Página oficial "
        "(business.ucf.edu) diz apenas 'Application deadlines vary by track "
        "-- contact Office of Professional Development (407-823-2185) or "
        "Graduate Business Programs (407-235-3917)'. Recomenda-se contato "
        "direto antes de aplicar."
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    all_fields = set()
    for updates in SCHOOL_UPDATES.values():
        all_fields.update(updates.keys())
    unknown_cols = [h for h in all_fields if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    no_fill = PatternFill(fill_type=None)

    report = []
    not_found = []

    for (nome, estado, cidade), updates in SCHOOL_UPDATES.items():
        target_rows = [
            row[0].row
            for row in ws.iter_rows(min_row=2)
            if row[col_uni - 1].value == nome
            and row[col_estado - 1].value == estado
            and row[col_cidade - 1].value == cidade
        ]
        if len(target_rows) == 0:
            not_found.append(nome)
            continue
        if len(target_rows) > 1:
            raise ValueError(f"Mais de uma linha corresponde a '{nome}' (linhas {target_rows}). Abortando por seguranca.")

        row_idx = target_rows[0]
        before = {}
        after = {}
        for header, new_value in updates.items():
            col_idx = header_to_col[header]
            cell = ws.cell(row=row_idx, column=col_idx)
            before[header] = cell.value
            cell.value = new_value
            cell.fill = no_fill
            after[header] = new_value
        report.append((nome, row_idx, before, after))

    if not_found:
        raise ValueError(f"Universidades nao encontradas na planilha (abortando, nada foi salvo): {not_found}")

    wb.save(DB_PATH)

    print(f"\n{len(report)} faculdades atualizadas:\n")
    for nome, row_idx, before, after in report:
        print(f"=== {nome} (linha {row_idx}) ===")
        for header in after:
            print(f"  - {header}:")
            print(f"      Antes: {before[header]!r}")
            print(f"      Depois: {after[header]!r}")
        print()


if __name__ == "__main__":
    main()
