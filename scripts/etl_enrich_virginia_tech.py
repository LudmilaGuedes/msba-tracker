"""
Data enrichment: atualiza APENAS a linha da 'Virginia Tech' (Pamplin College
of Business, Blacksburg, VA) no msba_programs_database.xlsx.

Foco exclusivo: M.S. in Business Administration - concentracao em Business
Analytics (MSBA-BA), presencial, 30 creditos, 12 meses. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/ingles, e custo de vida em Blacksburg.

ACHADO CRITICO DA AUDITORIA: 'STEM Designated' estava 'Sim' na planilha.
Multiplas fontes independentes (busca web + paginas oficiais adjacentes,
ago/2026) confirmam que o MSBA-BA de Virginia Tech NAO E STEM-certified
-- CIP code 52.0201 (Business Administration and Management, General),
que nao consta na lista STEM do DHS. CORRIGIDO para 'Nao'. Isso e
especialmente relevante dado que STEM designation e requisito obrigatorio
do perfil da usuaria (36 meses de OPT).

'Nome Oficial do Programa' estava uma paráfrase ('MS in Business Admin -
BA concentration'); CORRIGIDO para o nome oficial usado pela propria
Virginia Tech: 'M.S. in Business Administration - Business Analytics
(MSBA-BA)'.

'Inicio' estava 'Fall'; a pagina oficial de curriculo (msba.vt.edu/ba/
curriculum.html) mostra que o programa comeca no termo Summer II (curso
de Managerial Statistics + orientacao obrigatoria), seguido por Fall e
Spring -- CORRIGIDO para 'Summer II'.

'Ranking QS' estava preenchido com 20. Nao foi encontrada nenhuma fonte
que confirme um ranking QS especifico para o MSBA-BA de Virginia Tech;
o valor 20 parece ter sido confundido com o ranking #20 do TechGuide
(2026) para programas de Master's in Business Analytics, uma fonte
diferente da QS. CORRIGIDO de volta para 'Verificar' ate uma fonte QS
real ser encontrada.

Tuition (fonte oficial: Bursar's Office, "2026-2027 Tuition & Fee Rates
with COA" PDF, tabela "Graduate Blacksburg Campus" + linha de
"Masters of Science in Business Administration- Bus Analytics (MSBA-BA)
Supplemental Fee: $175.00 per credit hour"):
  - Tuition (Residente VA): $928.25/credito
  - Tuition (Nao-VA/Internacional): $1.865,50/credito
  - Estes sao os valores de TUITION BASE (sem a supplemental fee), que e
    o que os campos 'Custo/Credito' e 'Tuition Total' do schema
    representam (multiplicado por 30 creditos).
  - A supplemental fee de $175/credito ($5.250 no total do programa) e
    igual para todos independente de residencia -- registrada em 'Fees
    Adicionais USD'. Fees adicionais de biblioteca/comprehensive/CFEF
    (que variam por semestre e carga de creditos, e nao sao decomponiveis
    com confianca num unico valor de "programa inteiro" a partir da
    tabela publicada) foram DELIBERADAMENTE EXCLUIDAS deste campo, mesmo
    criterio ja aplicado no script da NC State: nao inventar uma
    granularidade que a propria universidade nao oferece como total
    unico.

Custo de vida: 'Transporte/mes' preenchido com $0 -- Blacksburg Transit e
gratuito para estudantes de Virginia Tech mediante Hokie Passport (fonte:
ridebt.org/vt). 'Seguro Saude (programa)' preenchido com $3.993 (taxa
padrao nao-subsidiada do Virginia Tech Student Medical Insurance Plan
2026-27: Fall $2.012 + Spring $1.981; fonte: risk.vt.edu/smi/plancosts).
'Utilities/mes' permanece 'Verificar' -- a unica fonte encontrada e uma
media de apartamento inteiro (~$149/mes), sem uma forma confiavel de
ratear para um quarto individual. Com isso, 'Custo Vida Total (programa)
USD' tambem permanece 'Verificar' -- falta 1 dos 9 componentes.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://msba.vt.edu/ba/about.html (programa, formato, capstone,
  taxa de emprego/salario/bonus da turma de 2025: 97,5% / $82k / $5k)
- https://msba.vt.edu/ba/admission.html (GPA minimo 3.0, GMAT/GRE nao
  exigido, TOEFL/IELTS, deadlines, entrevista obrigatoria, taxa $75)
- https://msba.vt.edu/ba/curriculum.html (estrutura por termo: Summer II
  -> Fall -> Spring/Summer I capstone; 30 creditos = 15 core + 15
  analytics)
- Busca web multipla confirmando CIP 52.0201 e "not STEM certified"
  (msba-BA nao tem designacao STEM/OPT extension)
- https://www.bursar.vt.edu/.../2026-2027%20Tuition%20%20Fee%20Rates%20
  with%20COA.pdf (tabelas oficiais de tuition e supplemental fees)
- https://risk.vt.edu/smi/plancosts.html (Student Medical Insurance)
- https://ridebt.org/vt (Blacksburg Transit gratuito p/ alunos VT)
- Graduate School VT: padrao minimo TOEFL 80 (sub-scores >=16, exames
  a partir de 21/jan/2026) / IELTS 6.5 (sub-scores >=6.5) -- pagina do
  MSBA-BA nao especifica minimo proprio, entao o padrao institucional
  foi usado
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
    "Nome da Universidade": "Virginia Tech",
    "Estado": "VA",
    "Cidade": "Blacksburg",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "M.S. in Business Administration - Business Analytics (MSBA-BA)",  # CORRECAO: nome oficial completo, era paráfrase abreviada
    "STEM Designated": "Não",  # CORRECAO CRITICA: planilha tinha "Sim"; CIP 52.0201 nao e STEM-designated, confirmado por multiplas fontes
    "Início": "Summer II",  # CORRECAO: planilha tinha "Fall"; curriculo oficial mostra inicio em Summer II (curso + orientacao) antes do Fall
    "Custo/Crédito (Intl) USD": 1866,  # oficial 2026-27: $1.865,50/credito (tuition base, Graduate Blacksburg, nao-VA)
    "Custo/Crédito (Residente) USD": 928,  # oficial 2026-27: $928,25/credito (tuition base, Graduate Blacksburg, VA resident)
    "Tuition Total (Intl) USD": 55965,  # 30 creditos x $1.865,50
    "Tuition Total (Residente) USD": 27848,  # 30 creditos x $928,25 (arredondado)
    "Fees Adicionais USD": 5250,  # MSBA-BA Supplemental Fee: $175/credito x 30 (igual p/ todos, independe de residencia)
    "Seguro Saúde (programa) USD": 3993,  # VT Student Medical Insurance 2026-27, taxa padrao nao-subsidiada (Fall $2.012 + Spring $1.981)
    "Transporte/mês USD": 0,  # Blacksburg Transit e gratuito para alunos de Virginia Tech (Hokie Passport)
    "Ranking QS": VERIFICAR,  # CORRECAO: valor "20" nao confirmado como ranking QS real; parece confusao com ranking #20 do TechGuide 2026
    "Taxa Empregabilidade %": 97.5,  # taxa de colocacao oficial, turma de 2025
    "Salário Médio Pós USD": 82000,  # salario mediano oficial, turma de 2025 (+ bonus mediano de assinatura de $5k, nao capturado em campo separado)
    "Principais Empregadores": (
        "Patrocinadores/parceiros do capstone corporativo: NASA, Leidos, "
        "Lockheed Martin, Advance Auto Parts, Norfolk Southern"
    ),
    "GPA Mínimo": 3.0,  # "minimum cumulative GPA of 3.0 or higher" (excecoes possiveis com outros pontos fortes)
    "GMAT/GRE Obrigatório": "Não obrigatório (não exigido desde o ciclo 2022-23)",
    "Testes Inglês Aceitos": (
        "TOEFL (mín. 80, nenhuma sub-nota <16) ou IELTS (mín. 6.5, "
        "nenhuma sub-nota <6.5) — padrão institucional (Graduate School); "
        "dispensado se graduação em instituição de ensino em inglês"
    ),
    "Matérias Pré-Requisito": (
        "Não há pré-requisitos formais obrigatórios; recomenda-se noção "
        "básica de estatística, programação e banco de dados — "
        "deficiências não impedem admissão mas devem ser corrigidas "
        "antes do início do programa"
    ),
    "Taxa Aplicação USD": 75,
    "Entrevista": "Sim (obrigatória para candidatos qualificados, após revisão da aplicação online)",
    "Rounds Aplicação": (
        "Contínuo (avaliação mensal); Early Action 1º nov; deadline final "
        "internacional 1º fev; deadline final doméstico 1º mar"
    ),
    "Deadline Aplicação": "Early Action 1º Nov; Final Intl 1º Fev; Final Doméstico 1º Mar",
    "Processo TA/RA": (
        "Sem processo formal centralizado no MSBA-BA; parte da turma "
        "recebe assistantships via Center for Business Analytics, outros "
        "departamentos/centros do Pamplin College ou outras unidades da "
        "universidade"
    ),
    "Graduate Assistantships": (
        "Disponíveis para parte da turma a cada ano, via Center for "
        "Business Analytics ou outras unidades de Virginia Tech; não é "
        "garantido nem estruturado como via automática"
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
        raise ValueError("Nenhuma linha encontrada para Virginia Tech (Blacksburg, VA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Virginia Tech, Blacksburg, VA)")

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
