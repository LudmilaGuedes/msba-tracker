import { VERIFICAR, type CellValue } from "../types/program";
import type { ColumnType } from "../types/columns";

export function isVerificar(value: CellValue | undefined | null): boolean {
  if (value === undefined || value === null) return true;
  return typeof value === "string" && value.trim().toLowerCase() === VERIFICAR.toLowerCase();
}

/** Converte um CellValue em número, ou null se não for numérico (ex.: "Verificar"). */
export function parseNumeric(value: CellValue | undefined | null): number | null {
  if (value === undefined || value === null) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const cleaned = value.replace(/[^0-9.\-]/g, "");
  if (cleaned === "" || cleaned === "-") return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

export function formatCellValue(value: CellValue | undefined | null, type: ColumnType): string {
  if (isVerificar(value)) return "Verificar";
  if (value === undefined || value === null || value === "") return "—";

  switch (type) {
    case "currency": {
      const n = parseNumeric(value);
      if (n === null) return String(value);
      return n.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
    }
    case "percent": {
      const n = parseNumeric(value);
      if (n === null) return String(value);
      return `${n}%`;
    }
    case "number": {
      const n = parseNumeric(value);
      if (n === null) return String(value);
      return n.toLocaleString("en-US");
    }
    default:
      return String(value);
  }
}

export function formatUSD(n: number | null): string {
  if (n === null) return "—";
  return n.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
}
