import * as XLSX from "xlsx";
import { COLUMNS } from "../types/columns";
import type { Folder, MergedProgram } from "../types/program";
import type { StrategyDataStore } from "../types/strategy";
import { classifyLeadTimeBucket, LEAD_TIME_BUCKET_LABELS, parseRounds } from "./strategyParsing";

function folderNames(program: MergedProgram, folders: Folder[]): string {
  return program.folderIds
    .map((id) => folders.find((f) => f.id === id)?.name)
    .filter(Boolean)
    .join(", ");
}

function strategyColumns(program: MergedProgram, strategyData: StrategyDataStore): Record<string, string | number> {
  const s = strategyData[program.id];
  if (!s) return {};
  const roundText = parseRounds(program.values.deadlineAplicacao)[0].dateText;
  const bucket = classifyLeadTimeBucket(program.values.deadlineAplicacao);
  return {
    "Estratégia: Janela do Round": LEAD_TIME_BUCKET_LABELS[bucket],
    "Estratégia: Data de Aplicação (1º round)": roundText,
    "Estratégia: Taxa de Aplicação ($)": s.applicationFee ?? "",
    "Estratégia: OBS Taxa": s.feeNotes ?? "",
    "Estratégia: GMAT Vou Enviar": s.gmatWillSend ? "Sim" : "Não",
    "Estratégia: Data Prevista Resultado": s.resultDate ?? "",
  };
}

function timestampSuffix(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
}

export function exportToExcel(programs: MergedProgram[], folders: Folder[], strategyData: StrategyDataStore = {}) {
  const rows = programs.map((p) => {
    const row: Record<string, string | number> = {};
    for (const col of COLUMNS) {
      row[col.header] = p.values[col.key] ?? "";
    }
    row["Minha Nota (0-5)"] = p.rating;
    row["Meus Comentários"] = p.comments;
    row["Minhas Pastas"] = folderNames(p, folders);
    Object.assign(row, strategyColumns(p, strategyData));
    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "MSBA Programs");
  XLSX.writeFile(workbook, `msba-meus-dados-${timestampSuffix()}.xlsx`);
}

export function exportToJson(programs: MergedProgram[], folders: Folder[]) {
  const payload = {
    exportedAt: new Date().toISOString(),
    folders,
    programs: programs.map((p) => ({
      id: p.id,
      rating: p.rating,
      comments: p.comments,
      folderIds: p.folderIds,
      editedFields: Array.from(p.editedKeys),
      data: p.values,
    })),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `msba-meus-dados-${timestampSuffix()}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
