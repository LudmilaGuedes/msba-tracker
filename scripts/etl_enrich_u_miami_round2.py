"""
Data enrichment (round 2): atualiza a linha da 'University of Miami'
(Miami Herbert Business School, MSBA) no msba_programs_database.xlsx,
a pedido do usuário: detalhar como cada bolsa funciona, elegibilidade,
valores, quão realista é, confirmar taxa de aplicação (e se em algum
momento fica de graça), e confirmar empregabilidade/salário.

NOTA: o domínio oficial migrou de herbert.miami.edu para
mbs.miami.edu (redirect 301 confirmado em todas as páginas
acessadas) -- registrado aqui porque outras linhas desta planilha
(documentadas em rounds anteriores) ainda citam URLs herbert.miami.edu
que hoje redirecionam.

FONTES (oficiais, mbs.miami.edu, via WebFetch/WebSearch em
2026-09-08):
- https://go.miami.edu/admit-faq
- https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/cost-of-attendance.html
- https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/requirements.html
- https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/index.html

CONFIRMADO SEM MUDANÇA (já batia exato):
- Taxa de Aplicação: $100 -- CONFIRMADO oficial.
- Taxa Empregabilidade: 94% -- CONFIRMADO oficial, e desta vez
  confirmado ser ESPECÍFICO do MSBA (coorte 2024, "94% of students
  earning jobs within six months of graduation"), não um agregado de
  outro programa -- boa notícia, o dado já era confiável.
- GMAT/GRE Obrigatório ("Opcional, recomendado para bolsa por mérito
  máxima"): CONFIRMADO -- a página oficial confirma textualmente que
  "submitting an unofficial test score can be beneficial for securing
  the maximum merit-based scholarship" -- ou seja, o dado que já
  estava na planilha já capturava essa nuance corretamente.
- Testes de Inglês (TOEFL 94/IELTS 7.0/Duolingo 125): CONFIRMADO
  oficial, sem mudança.
- Entrevista (por convite, virtual, ~30 min): CONFIRMADO oficial.
- Quão Realista (Intl) -- "Alta (74% dos admitidos no Fall 2025
  receberam bolsa por mérito; +$10 milhões concedidos aos
  mestrados/doutorados em 2025)": CONFIRMADO exato, mesmos números
  encontrados agora.

COMO FUNCIONA CADA BOLSA (resposta detalhada à pergunta do usuário):
1. BOLSA DE MÉRITO PADRÃO (a que vale pra você): automática, SEM
   inscrição separada -- toda candidata admitida é automaticamente
   avaliada. Baseada em histórico acadêmico, experiência profissional
   e potencial de liderança. Valor: de PARCIAL a TUITION INTEGRAL
   ("range from partial to full-tuition scholarships") -- não é um
   percentual fixo (diferente da Baylor, que é sempre 65%) nem uma
   faixa em dólares publicada (diferente da Fordham) -- varia caso a
   caso. Submeter GMAT/GRE (mesmo não-oficial) AJUDA a conseguir o
   valor MÁXIMO dessa bolsa -- é o único alavanque concreto e
   documentado que a candidata controla.
   ELEGIBILIDADE: você é elegível automaticamente ao se candidatar,
   sem nenhum requisito extra além da aplicação normal.
   QUÃO REALISTA: Alta -- 74% dos admitidos no Fall 2025 receberam
   essa bolsa, e a escola concedeu +$10 milhões no total (agregado de
   todos os mestrados/doutorados, não só MSBA).
2. UKRAINIAN GRADUATE BUSINESS SCHOLARSHIP: bolsa de tuition integral
   (2 disponíveis no total, uma para MBA e uma para um Specialized
   Master's) -- EXCLUSIVA para candidatos impactados pela guerra na
   Ucrânia. NÃO SE APLICA a candidatas brasileiras (não são elegíveis por nacionalidade/
   situação). Citado aqui só para contexto de que a escola tem bolsas
   de tuition integral para casos específicos, não generalizáveis.
3. Fellowships nomeadas (Emery Means Findley Jr., Merrill Lynch &
   Co. Foundation em International Finance, James W. McLamore, E.
   Bruce McLaughlin & Cynthia M. Swol em Marketing): são para
   candidatos "outstanding" de programas de negócio em geral -- não
   há detalhe de valor, processo de aplicação, ou se cobrem o MSBA
   especificamente. Não incluídas na estimativa de "quão realista"
   por falta de dado concreto aplicável.

TAXA DE APLICAÇÃO -- FICA DE GRAÇA EM ALGUM MOMENTO? Resposta: SIM, mas
não por round/deadline -- por participação em evento. "Miami Herbert
offers a limited number of application fee waivers. To qualify,
applicants must register and attend one of our in-person or virtual
informational events." Ou seja: participar de um webinar/evento
informativo oficial do programa dá chance de isenção da taxa de $100
-- é limitado ("limited number"), não é garantido, e não está
vinculado a nenhum round específico de aplicação.

ENRIQUECIDO -- 'Rounds Aplicação' (estava 'Verificar'): a página de
requisitos não lista rounds numerados, só dois deadlines fixos por
semestre de entrada (Fall: 1 jun; Spring: 1 fev-15 mar, já
confirmados no campo 'Deadline Aplicação'). Documentada a ausência de
estrutura de rounds e a condição de isenção de taxa via evento.

NÃO CONFIRMADO -- 'Salário Médio Pós USD' permanece 'Verificar': a
própria página de outcomes do MSBA (mbs.miami.edu/.../business-
analytics/index.html) só publica a taxa de emprego (94%), SEM
nenhuma cifra de salário médio/mediano para o programa. Não há
relatório de salário específico do MSBA publicado por esta escola.
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
    "Nome da Universidade": "University of Miami",
    "Estado": "FL",
    "Cidade": "Coral Gables",
}

UPDATES = {
    "Tipos de Bolsa": (
        "Merit-based automático (sem inscrição separada), baseado em "
        "histórico acadêmico, experiência profissional e potencial de "
        "liderança. Valor: de PARCIAL a TUITION INTEGRAL (não é "
        "percentual fixo nem faixa em dólares publicada -- varia caso "
        "a caso). Submeter GMAT/GRE (mesmo não-oficial) ajuda a "
        "conseguir o valor MÁXIMO. Também existem 2 bolsas de tuition "
        "integral exclusivas para candidatos afetados pela guerra na "
        "Ucrânia (não aplicável) e fellowships nomeadas sem valor/"
        "processo divulgado."
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados -- 2 deadlines fixos por semestre de "
        "entrada (ver Deadline Aplicação). Taxa de $100 pode ser "
        "isenta (número limitado de isenções) para quem participa de "
        "um evento informativo oficial (webinar/sessão presencial) "
        "antes de aplicar -- não é isenção automática nem vinculada a "
        "round/deadline específico."
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
        raise ValueError("Nenhuma linha encontrada para University of Miami. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Miami)")

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
