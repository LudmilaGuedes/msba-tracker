// Backup automático: escreve o mesmo export "Excel" (55 colunas + Minha
// Nota/Meus Comentários/Minhas Pastas) direto no arquivo escolhido pelo
// usuário via File System Access API, sobrescrevendo-o a cada 24h — sem
// depender de servidor, já que o app não tem um. Só funciona em
// Chrome/Edge; navegadores sem suporte caem para o export manual normal
// (ExportButton).
import * as XLSX from "xlsx";
import { COLUMNS } from "../types/columns";
import type { Folder, MergedProgram } from "../types/program";
import { clearBackupHandle, loadBackupHandle, storeBackupHandle } from "./fileHandleDb";

const LAST_BACKUP_KEY = "msba-tracker:last-backup-at:v1";
export const BACKUP_INTERVAL_MS = 24 * 60 * 60 * 1000;

export function isFileSystemAccessSupported(): boolean {
  return typeof window !== "undefined" && "showSaveFilePicker" in window;
}

function folderNames(program: MergedProgram, folders: Folder[]): string {
  return program.folderIds
    .map((id) => folders.find((f) => f.id === id)?.name)
    .filter(Boolean)
    .join(", ");
}

function buildWorkbookBlob(programs: MergedProgram[], folders: Folder[]): Blob {
  const rows = programs.map((p) => {
    const row: Record<string, string | number> = {};
    for (const col of COLUMNS) row[col.header] = p.values[col.key] ?? "";
    row["Minha Nota (0-5)"] = p.rating;
    row["Meus Comentários"] = p.comments;
    row["Minhas Pastas"] = folderNames(p, folders);
    return row;
  });
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "MSBA Programs");
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as Uint8Array<ArrayBuffer>;
  // Empacotado num Blob (em vez de escrever o Uint8Array direto) porque o
  // overload de FileSystemWritableFileStream.write() do lib DOM instalado
  // exige ArrayBufferView<ArrayBuffer>, e o array tipado que o SheetJS
  // devolve é genérico sobre ArrayBufferLike — Blob evita esse atrito de tipos.
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function chooseBackupFile(): Promise<FileSystemFileHandle> {
  const handle = await window.showSaveFilePicker({
    suggestedName: "msba-backup.xlsx",
    types: [
      {
        description: "Planilha Excel",
        accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
      },
    ],
  });
  await storeBackupHandle(handle);
  return handle;
}

export function getBackupHandle(): Promise<FileSystemFileHandle | null> {
  return loadBackupHandle();
}

export async function forgetBackupFile(): Promise<void> {
  await clearBackupHandle();
  localStorage.removeItem(LAST_BACKUP_KEY);
}

export async function hasReadWritePermission(handle: FileSystemFileHandle, requestIfNeeded: boolean): Promise<boolean> {
  const opts = { mode: "readwrite" as const };
  const state = await handle.queryPermission(opts);
  if (state === "granted") return true;
  if (!requestIfNeeded) return false;
  const requested = await handle.requestPermission(opts);
  return requested === "granted";
}

export async function writeBackupNow(handle: FileSystemFileHandle, programs: MergedProgram[], folders: Folder[]): Promise<void> {
  const blob = buildWorkbookBlob(programs, folders);
  const writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();
  localStorage.setItem(LAST_BACKUP_KEY, String(Date.now()));
}

export function getLastBackupAt(): number | null {
  const raw = localStorage.getItem(LAST_BACKUP_KEY);
  return raw ? Number(raw) : null;
}

export function isBackupDue(): boolean {
  const last = getLastBackupAt();
  if (last === null) return true;
  return Date.now() - last >= BACKUP_INTERVAL_MS;
}
