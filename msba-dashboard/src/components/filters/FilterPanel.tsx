import { useMemo } from "react";
import { RotateCcw, Search } from "lucide-react";
import type { MergedProgram, Folder } from "../../types/program";
import { DEFAULT_FILTERS, hasActiveFilters, type FilterState } from "../../types/filters";
import { RangeInput } from "./RangeInput";

interface FilterPanelProps {
  programs: MergedProgram[];
  folders: Folder[];
  filters: FilterState;
  onChange: (filters: FilterState) => void;
  resultCount: number;
}

export function FilterPanel({ programs, folders, filters, onChange, resultCount }: FilterPanelProps) {
  const states = useMemo(
    () => Array.from(new Set(programs.map((p) => String(p.values.estado)).filter(Boolean))).sort(),
    [programs],
  );

  const toggleState = (state: string) => {
    onChange({
      ...filters,
      states: filters.states.includes(state) ? filters.states.filter((s) => s !== state) : [...filters.states, state],
    });
  };

  return (
    <aside className="flex w-full shrink-0 flex-col gap-5 rounded-2xl border border-slate-200 bg-white p-4 lg:w-72">
      <div className="flex items-center justify-between">
        <h2 className="text-sm font-semibold text-slate-800">Filtros</h2>
        {hasActiveFilters(filters) && (
          <button
            onClick={() => onChange(DEFAULT_FILTERS)}
            className="flex items-center gap-1 text-xs font-medium text-indigo-600 hover:text-indigo-700"
          >
            <RotateCcw size={12} /> Limpar
          </button>
        )}
      </div>

      <p className="-mt-3 text-xs text-slate-400">{resultCount} de {programs.length} programas</p>

      <div className="relative">
        <Search size={15} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={filters.search}
          onChange={(e) => onChange({ ...filters, search: e.target.value })}
          placeholder="Buscar universidade ou programa..."
          className="w-full rounded-lg border border-slate-200 py-2 pl-8 pr-3 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </div>

      <div>
        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Cidade</div>
        <input
          value={filters.city}
          onChange={(e) => onChange({ ...filters, city: e.target.value })}
          placeholder="Ex.: Boston, Austin..."
          className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </div>

      <div>
        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">
          Estado {filters.states.length > 0 && `(${filters.states.length})`}
        </div>
        <div className="flex max-h-40 flex-wrap gap-1.5 overflow-y-auto pr-1">
          {states.map((state) => (
            <button
              key={state}
              onClick={() => toggleState(state)}
              className={`rounded-full border px-2.5 py-1 text-xs font-medium transition-colors ${
                filters.states.includes(state)
                  ? "border-indigo-500 bg-indigo-500 text-white"
                  : "border-slate-200 text-slate-600 hover:border-indigo-300"
              }`}
            >
              {state}
            </button>
          ))}
        </div>
      </div>

      <RangeInput
        label="Duração (meses)"
        value={filters.duration}
        onChange={(duration) => onChange({ ...filters, duration })}
      />

      <RangeInput
        label="Tuition Total (Intl)"
        value={filters.tuition}
        onChange={(tuition) => onChange({ ...filters, tuition })}
        prefix="$"
      />

      <RangeInput
        label="Custo Total Real (tuition + vida)"
        value={filters.totalCost}
        onChange={(totalCost) => onChange({ ...filters, totalCost })}
        prefix="$"
      />

      <div>
        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Pasta</div>
        <select
          value={filters.folderId ?? ""}
          onChange={(e) => onChange({ ...filters, folderId: e.target.value || null })}
          className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        >
          <option value="">Todas</option>
          {folders.map((f) => (
            <option key={f.id} value={f.id}>
              {f.name}
            </option>
          ))}
        </select>
      </div>

      <div className="flex flex-col gap-2 border-t border-slate-100 pt-3">
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input
            type="checkbox"
            checked={filters.onlyFiveStars}
            onChange={(e) => onChange({ ...filters, onlyFiveStars: e.target.checked })}
            className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-400"
          />
          Apenas 5 estrelas
        </label>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input
            type="checkbox"
            checked={filters.onlyWithNotes}
            onChange={(e) => onChange({ ...filters, onlyWithNotes: e.target.checked })}
            className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-400"
          />
          Com anotações
        </label>
        <label className="flex items-center gap-2 text-sm text-slate-600">
          <input
            type="checkbox"
            checked={filters.onlyStem}
            onChange={(e) => onChange({ ...filters, onlyStem: e.target.checked })}
            className="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-400"
          />
          Somente STEM designated
        </label>
      </div>
    </aside>
  );
}
