// Lê um arquivo de backup (mesmo formato do export "Excel": 55 colunas do
// schema + Minha Nota/Meus Comentários/Minhas Pastas) e devolve linhas
// estruturadas, prontas pra DataContext.restoreFromBackup() casar contra os
// programas atuais e reconstruir o overlay do usuário.
import * as XLSX from "xlsx";
import { COLUMNS } from "../types/columns";
import type { CellValue } from "../types/program";

export interface BackupRow {
  universidade: string;
  estado: string;
  cidade: string;
  /** Os campos do schema presentes no backup, já mapeados PT -> chave camelCase. */
  values: Record<string, CellValue>;
  rating: number;
  comments: string;
  folderNames: string[];
}

const SHEET_NAME = "MSBA Programs";
const RATING_HEADER = "Minha Nota (0-5)";
const COMMENTS_HEADER = "Meus Comentários";
const FOLDERS_HEADER = "Minhas Pastas";

export async function parseBackupFile(file: File): Promise<BackupRow[]> {
  const buffer = await file.arrayBuffer();
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
    const values: Record<string, CellValue> = {};
    for (const col of COLUMNS) {
      const idx = headerIndex[col.header];
      // Backup de antes de a coluna existir: deixa sem valor, pra não virar
      // uma edição que apaga o dado atual da planilha.
      if (idx === undefined) continue;
      const raw = row[idx];
      values[col.key] = typeof raw === "number" ? raw : String(raw ?? "").trim();
    }

    const ratingIdx = headerIndex[RATING_HEADER];
    const commentsIdx = headerIndex[COMMENTS_HEADER];
    const foldersIdx = headerIndex[FOLDERS_HEADER];
    const ratingRaw = ratingIdx === undefined ? 0 : Number(row[ratingIdx]);
    const foldersRaw = foldersIdx === undefined ? "" : String(row[foldersIdx] ?? "");

    return {
      universidade: String(values.universidade ?? ""),
      estado: String(values.estado ?? ""),
      cidade: String(values.cidade ?? ""),
      values,
      rating: Number.isFinite(ratingRaw) ? ratingRaw : 0,
      comments: commentsIdx === undefined ? "" : String(row[commentsIdx] ?? ""),
      folderNames: foldersRaw
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean),
    };
  });
}
