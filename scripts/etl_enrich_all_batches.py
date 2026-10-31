# -*- coding: utf-8 -*-
"""
Data enrichment em massa (set/2026): aplica os lotes pesquisados em paralelo
(scripts/enrichment_batches/output_batch_NN.json) ao msba_programs_database.xlsx.

Cada lote foi produzido por um agente de pesquisa seguindo
scripts/enrichment_batches/INSTRUCOES.md (fontes oficiais, schema de 54
colunas). As fontes e observacoes de cada universidade ficam nos proprios
JSONs. Este script so valida e aplica; nao pesquisa nada.

Validacoes antes de gravar:
- match por Nome da Universidade + Estado + Cidade (exatamente 1 linha);
- cabecalho existe no schema; nunca altera Nome/Estado/Cidade;
- colunas numericas so recebem numero ou "Verificar" (texto com digitos
  vira numero-lixo no dashboard -- ver etl_fix_texto_em_colunas_numericas.py);
- Mercado/mes USD nunca e alterado (valor fixo de 300, decisao da usuaria);
- enums (Tipo, STEM, Bolsas Disponiveis) so com valores permitidos;
- valores vazios/None viram "Verificar".
Valores rejeitados sao listados no fim e NAO sao gravados.

Uso: python scripts/etl_enrich_all_batches.py [--dry-run] (a partir da raiz).
"""

import glob
import json
import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
BATCH_GLOB = "scripts/enrichment_batches/output_batch_*.json"
KEY_COLS = ("Nome da Universidade", "Estado", "Cidade")
LOCKED = set(KEY_COLS) | {"Mercado/mês USD"}

NUMERIC_COLS = {
    "Duração (meses)", "Tamanho Turma", "Total Créditos",
    "Ranking US News", "Ranking QS", "Taxa Empregabilidade %",
    "GPA Mínimo", "GPA Médio Admitidos", "GMAT Médio", "Taxa Aceitação %",
}
ENUMS = {
    "Tipo (Pub/Priv)": {"Pública", "Privada"},
    "STEM Designated": {"Sim", "Não"},
    "Bolsas Disponíveis": {"Sim", "Não"},
}

# Correcoes manuais sobre a linha da University of San Diego (etl_enrich_usd.py
# gravou Mercado/mes pesquisado e um texto com digitos em Ranking US News).
MANUAL = [
    ({"Nome da Universidade": "University of San Diego", "Estado": "CA", "Cidade": "San Diego"},
     {"Mercado/mês USD": 300, "Ranking US News": "Verificar"}),
]


def is_numeric_col(header):
    return header in NUMERIC_COLS or header.endswith("USD")


def validate(header, value):
    """Devolve (valor_normalizado, motivo_rejeicao_ou_None)."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return "Verificar", None
    if isinstance(value, bool):
        return None, "booleano"
    if is_numeric_col(header):
        if isinstance(value, (int, float)):
            return (int(value) if float(value).is_integer() else value), None
        if isinstance(value, str):
            s = value.strip()
            if s == "Verificar":
                return s, None
            try:
                num = float(s.replace(",", ""))
                return (int(num) if num.is_integer() else num), None
            except ValueError:
                if any(ch.isdigit() for ch in s):
                    return None, "texto com dígitos em coluna numérica"
                return s, None
        return None, f"tipo {type(value).__name__}"
    if header in ENUMS and value not in ENUMS[header]:
        return None, f"fora do enum {sorted(ENUMS[header])}"
    if isinstance(value, (list, dict)):
        return None, f"tipo {type(value).__name__}"
    return value, None


def main():
    dry_run = "--dry-run" in sys.argv
    batches = sorted(glob.glob(BATCH_GLOB))
    entries = []
    for path in batches:
        with open(path, encoding="utf-8") as fh:
            for item in json.load(fh):
                entries.append((path, item))
    print(f"{len(batches)} lotes, {len(entries)} universidades")

    if not dry_run:
        rotate_backup()
        print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    def find_row(match):
        rows = [
            r[0].row for r in ws.iter_rows(min_row=2)
            if all(r[col[k] - 1].value == match.get(k) for k in KEY_COLS)
        ]
        return rows[0] if len(rows) == 1 else None

    rejected, unmatched, total_changed = [], [], 0
    det = {"sim": [], "nao": [], "nao_confirmado": []}
    alerts = []

    work = [(p, it["match"], it.get("updates", {}), it) for p, it in entries]
    work += [("MANUAL", m, u, {}) for m, u in MANUAL]

    for path, match, updates, item in work:
        name = match.get("Nome da Universidade")
        row_idx = find_row(match)
        if row_idx is None:
            unmatched.append((path, match))
            continue
        if item:
            det.setdefault(item.get("det_aceito", "nao_confirmado"), []).append(
                f"{name}" + (f" ({item['det_nota_minima']})" if item.get("det_nota_minima") else ""))
            obs = item.get("observacoes") or ""
            if "ALERTA" in obs:
                alerts.append((name, obs))

        changed = []
        for header, raw in updates.items():
            if header not in col:
                rejected.append((name, header, raw, "coluna inexistente"))
                continue
            if header in LOCKED and path != "MANUAL":
                rejected.append((name, header, raw, "coluna bloqueada"))
                continue
            value, why = validate(header, raw)
            if why:
                rejected.append((name, header, raw, why))
                continue
            cell = ws.cell(row=row_idx, column=col[header])
            if cell.value == value:
                continue
            changed.append((header, cell.value, value))
            cell.value = value
            cell.fill = PatternFill(fill_type=None)

        total_changed += len(changed)
        print(f"\n=== {name} (linha {row_idx}): {len(changed)} campos")
        for header, before, after in changed:
            print(f"  - {header}: {before!r} -> {after!r}")

    if not dry_run:
        wb.save(DB_PATH)

    print(f"\n{'[DRY RUN] ' if dry_run else ''}TOTAL: {total_changed} campos alterados")
    print(f"\nDET aceito ({len(det.get('sim', []))}): {', '.join(det.get('sim', []))}")
    print(f"DET NÃO aceito ({len(det.get('nao', []))}): {', '.join(det.get('nao', []))}")
    print(f"DET não confirmado ({len(det.get('nao_confirmado', []))}): {', '.join(det.get('nao_confirmado', []))}")
    if alerts:
        print("\nALERTAS:")
        for name, obs in alerts:
            print(f"  - {name}: {obs}")
    if unmatched:
        print("\nSEM MATCH (não aplicados):")
        for path, m in unmatched:
            print(f"  - {path}: {m}")
    if rejected:
        print("\nVALORES REJEITADOS (não gravados):")
        for name, header, raw, why in rejected:
            print(f"  - {name} | {header} = {raw!r} ({why})")


if __name__ == "__main__":
    main()
