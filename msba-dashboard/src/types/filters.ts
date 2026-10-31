export interface NumberRange {
  min: number | null;
  max: number | null;
}

export interface FilterState {
  search: string;
  states: string[];
  city: string;
  duration: NumberRange;
  tuition: NumberRange;
  totalCost: NumberRange;
  onlyFiveStars: boolean;
  onlyWithNotes: boolean;
  onlyStem: boolean;
  folderId: string | null;
}

export const EMPTY_RANGE: NumberRange = { min: null, max: null };

export const DEFAULT_FILTERS: FilterState = {
  search: "",
  states: [],
  city: "",
  duration: { ...EMPTY_RANGE },
  tuition: { ...EMPTY_RANGE },
  totalCost: { ...EMPTY_RANGE },
  onlyFiveStars: false,
  onlyWithNotes: false,
  onlyStem: false,
  folderId: null,
};

export function hasActiveFilters(f: FilterState): boolean {
  return (
    f.search.trim() !== "" ||
    f.states.length > 0 ||
    f.city.trim() !== "" ||
    f.duration.min !== null ||
    f.duration.max !== null ||
    f.tuition.min !== null ||
    f.tuition.max !== null ||
    f.totalCost.min !== null ||
    f.totalCost.max !== null ||
    f.onlyFiveStars ||
    f.onlyWithNotes ||
    f.onlyStem ||
    f.folderId !== null
  );
}
