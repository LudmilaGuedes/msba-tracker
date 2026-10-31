import { useEffect, useRef, useState } from "react";
import { ChevronDown, Download, FileJson, FileSpreadsheet } from "lucide-react";
import type { Folder, MergedProgram } from "../types/program";
import { exportToExcel, exportToJson } from "../utils/exportData";
import { useStrategy } from "../context/StrategyContext";

interface ExportButtonProps {
  programs: MergedProgram[];
  folders: Folder[];
}

export function ExportButton({ programs, folders }: ExportButtonProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const { strategyData } = useStrategy();

  useEffect(() => {
    if (!open) return;
    const onClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, [open]);

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-2 rounded-lg bg-indigo-600 px-3.5 py-2 text-sm font-medium text-white shadow-sm hover:bg-indigo-700"
      >
        <Download size={16} /> Exportar Meus Dados <ChevronDown size={14} />
      </button>
      {open && (
        <div className="absolute right-0 top-full z-20 mt-1.5 w-56 rounded-xl border border-slate-200 bg-white p-1.5 shadow-lg">
          <button
            onClick={() => {
              exportToExcel(programs, folders, strategyData);
              setOpen(false);
            }}
            className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"
          >
            <FileSpreadsheet size={16} className="text-emerald-600" /> Excel (.xlsx)
          </button>
          <button
            onClick={() => {
              exportToJson(programs, folders);
              setOpen(false);
            }}
            className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"
          >
            <FileJson size={16} className="text-amber-600" /> JSON
          </button>
        </div>
      )}
    </div>
  );
}
