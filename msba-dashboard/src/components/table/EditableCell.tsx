import { useEffect, useRef, useState } from "react";
import { Pencil } from "lucide-react";
import type { CellValue } from "../../types/program";
import type { ColumnType } from "../../types/columns";
import { formatCellValue, isVerificar } from "../../utils/format";

interface EditableCellProps {
  value: CellValue;
  type: ColumnType;
  edited: boolean;
  options?: string[];
  onSave: (value: CellValue) => void;
}

export function EditableCell({ value, type, edited, options, onSave }: EditableCellProps) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(String(value ?? ""));
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [editing]);

  const commit = () => {
    setEditing(false);
    const trimmed = draft.trim();
    if (trimmed === String(value ?? "")) return;
    if ((type === "number" || type === "currency" || type === "percent") && trimmed !== "") {
      const n = Number(trimmed.replace(/[^0-9.\-]/g, ""));
      onSave(Number.isFinite(n) ? n : trimmed);
    } else {
      onSave(trimmed);
    }
  };

  if (editing && type === "enum" && options) {
    return (
      <select
        autoFocus
        defaultValue={isVerificar(value) ? "" : String(value ?? "")}
        onChange={(e) => {
          setEditing(false);
          onSave(e.target.value === "" ? "Verificar" : e.target.value);
        }}
        onBlur={() => setEditing(false)}
        onClick={(e) => e.stopPropagation()}
        className="w-full rounded-md border border-indigo-300 bg-white px-2 py-1 text-sm shadow-sm outline-none ring-2 ring-indigo-100"
      >
        <option value="">Verificar</option>
        {options.map((opt) => (
          <option key={opt} value={opt}>
            {opt}
          </option>
        ))}
      </select>
    );
  }

  if (editing) {
    return (
      <input
        ref={inputRef}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit();
          if (e.key === "Escape") {
            setDraft(String(value ?? ""));
            setEditing(false);
          }
        }}
        onClick={(e) => e.stopPropagation()}
        className="w-full rounded-md border border-indigo-300 bg-white px-2 py-1 text-sm shadow-sm outline-none ring-2 ring-indigo-100"
      />
    );
  }

  const verificar = isVerificar(value);

  return (
    <button
      type="button"
      onClick={(e) => {
        e.stopPropagation();
        setDraft(String(value ?? ""));
        setEditing(true);
      }}
      className={`group flex w-full items-center justify-between gap-1.5 rounded-md px-2 py-1 text-left text-sm transition-colors hover:bg-slate-100 ${
        verificar ? "bg-amber-50 text-amber-700" : "text-slate-700"
      }`}
      title="Clique para editar"
    >
      <span className="truncate">{formatCellValue(value, type)}</span>
      <span className="flex shrink-0 items-center gap-1">
        {edited && <span className="h-1.5 w-1.5 rounded-full bg-indigo-500" title="Editado por você" />}
        <Pencil size={12} className="opacity-0 text-slate-400 group-hover:opacity-100" />
      </span>
    </button>
  );
}
