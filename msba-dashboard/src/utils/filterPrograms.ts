import type { MergedProgram } from "../types/program";
import type { FilterState, NumberRange } from "../types/filters";
import { parseNumeric } from "./format";
import { getRealTotalCost } from "./derived";

function inRange(value: number | null, range: NumberRange): boolean {
  if (range.min === null && range.max === null) return true;
  if (value === null) return false;
  if (range.min !== null && value < range.min) return false;
  if (range.max !== null && value > range.max) return false;
  return true;
}

export function filterPrograms(programs: MergedProgram[], filters: FilterState): MergedProgram[] {
  const search = filters.search.trim().toLowerCase();
  const city = filters.city.trim().toLowerCase();

  return programs.filter((p) => {
    const v = p.values;

    if (search) {
      const haystack = `${v.universidade} ${v.nomePrograma} ${v.businessSchool}`.toLowerCase();
      if (!haystack.includes(search)) return false;
    }

    if (filters.states.length > 0 && !filters.states.includes(String(v.estado))) return false;

    if (city && !String(v.cidade).toLowerCase().includes(city)) return false;

    if (!inRange(parseNumeric(v.duracaoMeses), filters.duration)) return false;

    if (!inRange(parseNumeric(v.tuitionTotalIntl), filters.tuition)) return false;

    if (!inRange(getRealTotalCost(p), filters.totalCost)) return false;

    if (filters.onlyFiveStars && p.rating < 5) return false;

    if (filters.onlyWithNotes && !p.comments.trim()) return false;

    if (filters.onlyStem && v.stemDesignated !== "Sim") return false;

    if (filters.folderId && !p.folderIds.includes(filters.folderId)) return false;

    return true;
  });
}
