import { useEffect, useRef, useState } from "react";
import { Check, Clock, HardDriveDownload, RefreshCw, TriangleAlert, X } from "lucide-react";
import type { Folder, MergedProgram } from "../types/program";
import { useAutoBackup } from "../hooks/useAutoBackup";

interface AutoBackupButtonProps {
  programs: MergedProgram[];
  folders: Folder[];
}

function formatRelative(ts: number | null): string {
  if (ts === null) return "ainda sem backup";
  const minutes = Math.floor((Date.now() - ts) / 60000);
  if (minutes < 1) return "agora mesmo";
  if (minutes < 60) return `há ${minutes} min`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `há ${hours}h`;
  return `há ${Math.floor(hours / 24)}d`;
}

export function AutoBackupButton({ programs, folders }: AutoBackupButtonProps) {
  const { status, fileName, lastBackupAt, configure, grantPermission, backupNow, disable } = useAutoBackup(programs, folders);
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

  if (status === "unsupported") {
    return (
      <span
        className="hidden items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-xs text-slate-400 md:flex"
        title="Backup automático precisa do Chrome ou Edge."
      >
        <HardDriveDownload size={14} /> Backup automático indisponível
      </span>
    );
  }

  if (status === "unconfigured" || status === "checking") {
    return (
      <button
        onClick={configure}
        disabled={status === "checking"}
        className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50"
      >
        <HardDriveDownload size={16} /> Configurar backup automático
      </button>
    );
  }

  if (status === "needs-permission") {
    return (
      <button
        onClick={grantPermission}
        className="flex items-center gap-1.5 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-sm font-medium text-amber-700 hover:bg-amber-100"
      >
        <TriangleAlert size={16} /> Reativar backup em "{fileName}"
      </button>
    );
  }

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm font-medium text-emerald-700 hover:bg-emerald-100"
      >
        <Check size={16} /> Backup: {formatRelative(lastBackupAt)}
      </button>
      {open && (
        <div className="absolute right-0 top-full z-20 mt-1.5 w-64 rounded-xl border border-slate-200 bg-white p-1.5 shadow-lg">
          <div className="px-3 py-2 text-xs text-slate-500">
            Salvando automaticamente a cada 24h em
            <br />
            <span className="font-medium text-slate-700">{fileName}</span>
          </div>
          <button
            onClick={() => {
              backupNow();
              setOpen(false);
            }}
            className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"
          >
            <RefreshCw size={16} className="text-indigo-600" /> Salvar backup agora
          </button>
          <button
            onClick={() => {
              configure();
              setOpen(false);
            }}
            className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"
          >
            <Clock size={16} className="text-slate-500" /> Trocar arquivo de backup
          </button>
          <button
            onClick={() => {
              disable();
              setOpen(false);
            }}
            className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-left text-sm text-red-600 hover:bg-red-50"
          >
            <X size={16} /> Desativar backup automático
          </button>
        </div>
      )}
    </div>
  );
}
