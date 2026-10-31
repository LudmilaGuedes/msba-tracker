import * as XLSX from "xlsx";
import { COLUMNS } from "../types/columns";
import type { RawProgram } from "../types/program";
import { programId } from "./programId";

const SOURCE_URL = `${import.meta.env.BASE_URL}data/msba_programs_database.xlsx`;
const SHEET_NAME = "MSBA Programs";

/**
 * Baixa a planilha original (public/data) e converte cada linha em um
 * RawProgram, mapeando cabeçalhos PT -> chaves camelCase de COLUMNS.
 * O id vem de nome + estado + cidade (programId), então continua o mesmo
 * quando linhas são removidas ou reordenadas na planilha.
 */
export async function loadProgramsFromExcel(): Promise<RawProgram[]> {
  const res = await fetch(SOURCE_URL);
  if (!res.ok) {
    throw new Error(`Não foi possível carregar a planilha (${res.status}).`);
  }
  const buffer = await res.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const sheetName = workbook.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if (rows.length === 0) return [];

  const headerRow = rows[0].map((h) => String(h).trim());
  const headerIndex: Record<string, number> = {};
  headerRow.forEach((h, i) => {
    headerIndex[h] = i;
  });

  const dataRows = rows.slice(1).filter((row) => row.some((cell) => String(cell).trim() !== ""));

  return dataRows.map((row) => {
    const program: RawProgram = { id: "" };
    for (const col of COLUMNS) {
      const idx = headerIndex[col.header];
      const raw = idx === undefined ? "" : row[idx];
      program[col.key] = typeof raw === "number" ? raw : String(raw ?? "").trim();
    }
    program.id = programId(program.universidade, program.estado, program.cidade);
    return program;
  });
}
