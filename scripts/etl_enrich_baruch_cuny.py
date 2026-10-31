"""
Data enrichment: atualiza APENAS a linha da 'Baruch CUNY' (Zicklin
School of Business, New York, NY) no msba_programs_database.xlsx.

LIMITAÇÃO SEVERA DE ACESSO (documentada para transparência): o domínio
oficial do programa (zicklin.baruch.cuny.edu, em qualquer variante —
com/sem www, http/https) retornou erro de certificado SSL
("unable to verify the first certificate") em TODAS as tentativas de
acesso via WebFetch nesta rodada, tornando a página oficial do
programa, currículo, admissões e FAQ completamente inacessíveis.
Tentativas de bypass via Wayback Machine também falharam (domínio
bloqueado para esta ferramenta). O orçamento de busca web da sessão
também está esgotado (200/200 WebSearch). Como resultado, este
enrichment ficou limitado majoritariamente a dados de tuition
confirmáveis via o domínio oficial cuny.edu (tabela de tuition
university-wide), que É acessível. A maioria dos campos de admissão,
currículo, outcomes e custo de vida detalhado permanece 'Verificar'
por falta de fonte acessível nesta rodada -- não por ausência de dado
existente.

TUITION: CORRIGIDO -- achado importante. A tabela oficial de tuition da
CUNY (cuny.edu, 2026-27) mostra que Baruch tem categorias de tuition
SEPARADAS para MBA ($1.110/crédito não-residente) vs. "Masters Degree"
padrão ($855/crédito não-residente / $470/crédito residente
part-time / $5.545/semestre residente full-time), e a própria página
confirma explicitamente que MBA e MPA/MIA têm linhas de tuition
PRÓPRIAS, diferentes da categoria "Masters Degree" padrão, que é a
aplicável ao MS in Business Analytics. O valor anterior na planilha
('Tuition Total (Intl)' = $40.316) bate muito mais de perto com a
tarifa de MBA (36 × $1.110 = $39.960, próximo dos $40.316) do que com
a tarifa correta de "Masters Degree" (36 × $855 = $30.780) -- forte
indício de que o dado antigo confundiu a tarifa do MSBA com a do MBA
da mesma escola. CORRIGIDO para a tarifa correta de "Masters Degree":
'Custo/Crédito (Intl)' = $855, 'Custo/Crédito (Residente)' = $470
(tarifa part-time, usada por ser a única taxa por-crédito residente
disponível; a tarifa full-time é uma tarifa fixa de $5.545/semestre,
não por-crédito). 'Tuition Total' (36 créditos, já confirmado no
schema): Intl = 36 × $855 = $30.780 (CORREÇÃO SIGNIFICATIVA: quase
$10.000 mais barato que o valor anterior); Residente = 36 × $470 =
$16.920 (cálculo por crédito, usado por não haver confirmação da
estrutura exata de semestres do programa -- se o programa usa 3
semestres full-time de 12 créditos, o total via tarifa fixa seria
$16.635, próximo mas não idêntico; a diferença é pequena e documentada
aqui, sem impacto material).

'Fees Adicionais USD' e a maioria dos demais campos que dependem da
página oficial do programa (duração exata, currículo, admissões,
outcomes, seguro saúde, bolsas) PERMANECEM 'Verificar' -- calcular fees
exigiria assumir o número de semestres do programa, o que não foi
possível confirmar sem acesso à página oficial (evitando empilhar
estimativa sobre estimativa, princípio já estabelecido nesta sessão).

'Caminho Tuition Residente' preenchido com nota geral sobre residência
em NY/CUNY (mesmo padrão de outros estados públicos já tratados nesta
sessão): estabelecer residência NY para fins de tuition normalmente
exige domicílio prévio de 1 ano, o que é inviável para uma aluna
internacional em F-1 num programa curto -- portanto a tarifa
não-residente/internacional é a relevante na prática.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.cuny.edu/financial-aid/tuition-and-college-costs/tuition-fees/
  (tabela oficial de tuition CUNY 2026-27, com categorias separadas
  para MBA, MPA/MIA e "Masters Degree" padrão, fees obrigatórias
  gerais)
- https://www.baruch.cuny.edu/academic-degree-programs/ (confirma
  existência de versão presencial E versão 100% online do MS in
  Business Analytics/AI, mas sem detalhes de currículo/admissão)
- zicklin.baruch.cuny.edu (página oficial do programa) -- INACESSÍVEL
  nesta rodada devido a erro de certificado SSL em todas as variantes
  de URL testadas
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
    "Nome da Universidade": "Baruch CUNY",
    "Estado": "NY",
    "Cidade": "New York",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 855,  # oficial CUNY 2026-27, categoria "Masters Degree" padrão (não MBA)
    "Custo/Crédito (Residente) USD": 470,  # oficial CUNY 2026-27, tarifa part-time por crédito
    "Tuition Total (Intl) USD": 30780,  # CORREÇÃO: 36 x $855 (era $40.316, aparentemente usando tarifa de MBA por engano)
    "Tuition Total (Residente) USD": 16920,  # 36 x $470
    "Caminho Tuition Residente": (
        "Não aplicável na prática para internacionais em programa "
        "curto — NY/CUNY exige domicílio prévio de 1 ano para status "
        "de residente; tarifa não-residente/internacional é a "
        "relevante"
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
        raise ValueError("Nenhuma linha encontrada para Baruch CUNY (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baruch CUNY, New York, NY)")

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
