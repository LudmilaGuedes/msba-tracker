"""
Data enrichment: atualiza APENAS a linha da 'UGA' (Terry College of
Business, Athens, GA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM. Deep
research em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida local.

ACHADO CRITICO DA AUDITORIA: 'Duração (meses)' estava cadastrada como 12,
mas múltiplas fontes (UGA Today/news.uga.edu, página oficial do programa via
busca, brochure MSBA) descrevem consistentemente um programa de 10 MESES
(2 semestres, Fall->Spring, 33 créditos -- créditos já corretos). CORRIGIDO.

Tuition: terry.uga.edu bloqueou acesso automatizado (HTTP 403) em todas as
páginas tentadas, então a tuition veio direto da página oficial da UGA
Bursar (busfin.uga.edu) para 2026-27, que lista o MSBA com tarifa "E-rate"
diferenciada por residência: Residente $745/h (flat $8.940/semestre p/
12+ créditos), Não-residente doméstico $1.544/h ($18.528/semestre), Fora do
país (internacional) $1.589/h ($19.068/semestre). Como o MSBA é um programa
full-time em coorte fechada (~16-17 créditos/semestre, bem acima do piso de
12 para a tarifa flat), foi usada a tarifa FLAT de tempo integral x 2
semestres, não o cálculo por crédito x 33 (uma fonte secundária citava
$25.380 usando essa segunda abordagem com a tarifa residente do ano
anterior -- mantido apenas como nota, não usado, por vir de ano-base
diferente e já embutir uma metodologia distinta).

Assistantships: uma fonte dizia que "até 90% da turma de MSBA de DOIS ANOS"
recebe Graduate Assistantship com tuition reduzida a $25/semestre -- mas
"dois anos" contradiz a duração de 10 meses confirmada para o MSBA, sugerindo
que a fonte pode estar confundindo o MSBA com outro programa de mestrado da
Terry. Por isso esse dado NÃO foi usado para alterar 'TA/RA Disponível',
'Processo TA/RA' ou 'Graduate Assistantships' (mantidos como estavam).

Custo de vida: 'Utilities/mês' e 'Transporte/mês' foram preenchidos com
fontes razoavelmente confiáveis (COA oficial da UGA Financial Aid para
transporte; blog especializado em moradia estudantil de Athens para
utilities, dentro de uma faixa plausível) -- com isso, e usando os 6
componentes mensais que já estavam na planilha + o seguro-saúde do
programa (já cadastrado), foi possível calcular 'Custo Vida Total
(programa) USD' pela primeira vez.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://news.uga.edu/business-analytics-degree/ (duração 10 meses, 2
  semestres, foco em ML/AI/SQL/R/Python)
- https://busfin.uga.edu/finance-division/bursar/students-parents/bursar-tuition-fees/2026-2027/
  (tarifas 2026-27 por residência para o MSBA)
- Busca por admissions: GPA mínimo 3.0; GMAT ou GRE aceitos, sem nota
  mínima, dispensa possível; TOEFL/IELTS só p/ não-nativos; pré-requisitos
  de estatística/cálculo/programação; experiência de estágio recomendada
- Busca por assistantships: existe programa de waiver que permite a
  internacionais pagar tarifa doméstica (in-state), critérios não
  detalhados publicamente
- osfa.uga.edu/costs/ (COA oficial 2025-26): transporte $1.528/ano
  acadêmico (9 meses) -> ~$170/mês
- ramblerathens.com (blog especializado em moradia estudantil de Athens):
  utilities entre $15-70/mês dependendo de uso e nº de moradores -- usado
  ponto médio-alto ($50) por ser mais consistente com "quarto individual"
  (menos divisão de custos com colegas de apê) do que com repúblicas
  grandes
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
    "Nome da Universidade": "UGA",
    "Estado": "GA",
    "Cidade": "Athens",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 10,  # CORRECAO: planilha tinha 12; múltiplas fontes confirmam 10 meses (2 semestres, Fall->Spring)
    "Custo/Crédito (Residente) USD": 745,  # tarifa E-rate oficial 2026-27
    "Custo/Crédito (Intl) USD": 1589,  # tarifa "out-of-country" (internacional) oficial 2026-27
    "Tuition Total (Residente) USD": 17880,  # tarifa flat de tempo integral ($8.940/semestre) x 2 semestres
    "Tuition Total (Intl) USD": 38136,  # tarifa flat de tempo integral internacional ($19.068/semestre) x 2 semestres
    "Utilities/mês USD": 50,  # aprox., faixa $15-70 encontrada para apartamentos estudantis em Athens
    "Transporte/mês USD": 170,  # $1.528/ano acadêmico (COA oficial UGA Financial Aid) / 9 meses
    "Custo Vida Total (programa) USD": 19700,  # (850+350+15+50+50+60+170+150) x 10 meses + 2.750 (seguro saúde, já cadastrado)
    "Caminho Tuition Residente": (
        "Existe programa de waiver que permite a estudantes "
        "internacionais pagar apenas a tarifa doméstica (in-state), mas "
        "com critérios específicos não detalhados publicamente"
    ),
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Ambos aceitos (GMAT ou GRE); sem nota mínima definida; dispensa "
        "de teste disponível para candidatos selecionados"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL ou IELTS -- exigido apenas para não-falantes nativos de "
        "inglês"
    ),  # detalhamento: planilha tinha só "TOEFL/IELTS" vago
    "Matérias Pré-Requisito": (
        "Estatística introdutória e cálculo em nível universitário; "
        "conhecimento de programação em pelo menos 1 linguagem de alto "
        "nível (R, Python, SQL, C, Java); 1+ estágio ou experiência "
        "profissional recomendada"
    ),
    "Documentos Necessários": (
        "Transcripts, GMAT ou GRE (sem nota mínima, dispensa possível "
        "para candidatos selecionados), TOEFL/IELTS (não-nativos), "
        "histórico de estágio/experiência profissional (recomendado)"
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
        raise ValueError("Nenhuma linha encontrada para UGA (Athens, GA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UGA, Athens, GA)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
