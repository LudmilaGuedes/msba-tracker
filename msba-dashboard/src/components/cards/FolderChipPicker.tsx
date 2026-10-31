import { useEffect, useRef, useState } from "react";
import { Check, FolderPlus } from "lucide-react";
import type { Folder } from "../../types/program";

interface FolderChipPickerProps {
  folders: Folder[];
  selectedIds: string[];
  onToggle: (folderId: string) => void;
  compact?: boolean;
}

export function FolderChipPicker({ folders, selectedIds, onToggle, compact = false }: FolderChipPickerProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, [open]);

  const selected = folders.filter((f) => selectedIds.includes(f.id));

  return (
    <div className="relative inline-block" ref={ref}>
      <button
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          setOpen((v) => !v);
        }}
        className="flex flex-wrap items-center gap-1 rounded-lg px-1.5 py-1 hover:bg-slate-100"
      >
        {selected.length === 0 && (
          <span className="flex items-center gap-1 text-xs text-slate-400">
            <FolderPlus size={14} /> {!compact && "Adicionar"}
          </span>
        )}
        {selected.map((f) => (
          <span
            key={f.id}
            className="rounded-full px-2 py-0.5 text-[11px] font-medium text-white"
            style={{ backgroundColor: f.color }}
          >
            {f.name}
          </span>
        ))}
      </button>

      {open && (
        <div
          className="absolute left-0 top-full z-20 mt-1 w-56 rounded-xl border border-slate-200 bg-white p-1.5 shadow-lg"
          onClick={(e) => e.stopPropagation()}
        >
          {folders.length === 0 && (
            <div className="px-2 py-2 text-xs text-slate-400">Crie uma pasta no gerenciador de pastas.</div>
          )}
          {folders.map((f) => {
            const isSelected = selectedIds.includes(f.id);
            return (
              <button
                key={f.id}
                onClick={() => onToggle(f.id)}
                className="flex w-full items-center gap-2 rounded-lg px-2 py-1.5 text-left text-sm hover:bg-slate-50"
              >
                <span className="h-2.5 w-2.5 rounded-full" style={{ backgroundColor: f.color }} />
                <span className="flex-1 truncate text-slate-700">{f.name}</span>
                {isSelected && <Check size={14} className="text-indigo-600" />}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
