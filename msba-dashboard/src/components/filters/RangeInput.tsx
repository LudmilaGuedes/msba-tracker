import type { NumberRange } from "../../types/filters";

interface RangeInputProps {
  label: string;
  value: NumberRange;
  onChange: (value: NumberRange) => void;
  placeholder?: [string, string];
  prefix?: string;
}

export function RangeInput({ label, value, onChange, placeholder = ["Mín", "Máx"], prefix }: RangeInputProps) {
  return (
    <div>
      <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">{label}</div>
      <div className="flex items-center gap-2">
        <div className="relative flex-1">
          {prefix && <span className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-xs text-slate-400">{prefix}</span>}
          <input
            type="number"
            value={value.min ?? ""}
            onChange={(e) => onChange({ ...value, min: e.target.value === "" ? null : Number(e.target.value) })}
            placeholder={placeholder[0]}
            className={`w-full rounded-lg border border-slate-200 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100 ${prefix ? "pl-6 pr-2" : "px-2"}`}
          />
        </div>
        <span className="text-slate-300">–</span>
        <div className="relative flex-1">
          {prefix && <span className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-xs text-slate-400">{prefix}</span>}
          <input
            type="number"
            value={value.max ?? ""}
            onChange={(e) => onChange({ ...value, max: e.target.value === "" ? null : Number(e.target.value) })}
            placeholder={placeholder[1]}
            className={`w-full rounded-lg border border-slate-200 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100 ${prefix ? "pl-6 pr-2" : "px-2"}`}
          />
        </div>
      </div>
    </div>
  );
}
