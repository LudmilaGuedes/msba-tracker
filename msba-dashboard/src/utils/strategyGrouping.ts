import type { Folder, MergedProgram } from "../types/program";
import { STRATEGY_ELIGIBLE_FOLDER_NAMES } from "../types/strategy";
import {
  classifyLeadTimeBucket,
  estimateDate,
  parseRounds,
  LEAD_TIME_BUCKET_ORDER,
  type LeadTimeBucket,
} from "./strategyParsing";

export function normalize(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "");
}

const ELIGIBLE_NORMALIZED = STRATEGY_ELIGIBLE_FOLDER_NAMES.map(normalize);

/** true quando o programa está em pelo menos uma das pastas de elegibilidade da Estratégia de Aplicação. */
export function isStrategyEligible(program: MergedProgram, folders: Folder[]): boolean {
  return program.folderIds.some((id) => {
    const folder = folders.find((f) => f.id === id);
    return folder && ELIGIBLE_NORMALIZED.includes(normalize(folder.name));
  });
}

/** true quando o programa está marcado com a pasta de nome `name` (case/acento-insensitive). */
export function hasFolderNamed(program: MergedProgram, folders: Folder[], name: string): boolean {
  const target = normalize(name);
  return program.folderIds.some((id) => {
    const folder = folders.find((f) => f.id === id);
    return folder && normalize(folder.name) === target;
  });
}

export interface StrategyEntry {
  program: MergedProgram;
  /** Texto do primeiro round/deadline do programa — sempre assumimos que o usuário aplica no primeiro round disponível. */
  roundText: string;
  bucket: LeadTimeBucket;
  dateEstimate: Date | null;
}

/** Monta uma entrada por programa, sempre baseada no primeiro round/deadline (Round 1, Early Action, "Prioritário"...). */
export function buildStrategyEntries(programs: MergedProgram[]): StrategyEntry[] {
  return programs.map((program) => {
    const roundText = parseRounds(program.values.deadlineAplicacao)[0].dateText;
    return {
      program,
      roundText,
      bucket: classifyLeadTimeBucket(program.values.deadlineAplicacao),
      dateEstimate: estimateDate(roundText),
    };
  });
}

/** Agrupa entradas pela janela do calendário do primeiro deadline (Round 1 set-dez / Round 2 jan-mai / Round 3 jun-ago / sem data), ordenando cada grupo pela data estimada do deadline. */
export function groupEntriesByBucket(entries: StrategyEntry[]): { bucket: LeadTimeBucket; entries: StrategyEntry[] }[] {
  return LEAD_TIME_BUCKET_ORDER.map((bucket) => ({
    bucket,
    entries: entries
      .filter((e) => e.bucket === bucket)
      .sort((a, b) => {
        if (a.dateEstimate && b.dateEstimate) return a.dateEstimate.getTime() - b.dateEstimate.getTime();
        if (a.dateEstimate) return -1;
        if (b.dateEstimate) return 1;
        return String(a.program.values.universidade ?? "").localeCompare(String(b.program.values.universidade ?? ""));
      }),
  })).filter((g) => g.entries.length > 0);
}
