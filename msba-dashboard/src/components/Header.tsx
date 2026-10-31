import { FolderKanban, LayoutGrid, Table2, BarChart3, Plus } from "lucide-react";
import type { Folder, MergedProgram } from "../types/program";
import { ExportButton } from "./ExportButton";
import { AutoBackupButton } from "./AutoBackupButton";
import { RestoreBackupButton } from "./RestoreBackupButton";

export type ViewMode = "table" | "cards";

interface HeaderProps {
  view: ViewMode;
  onViewChange: (view: ViewMode) => void;
  showChart: boolean;
  onToggleChart: () => void;
  onOpenFolders: () => void;
  onOpenAddProgram: () => void;
  programs: MergedProgram[];
  allPrograms: MergedProgram[];
  folders: Folder[];
}

export function Header({
  view,
  onViewChange,
  showChart,
  onToggleChart,
  onOpenFolders,
  onOpenAddProgram,
  programs,
  allPrograms,
  folders,
}: HeaderProps) {
  return (
    <header className="sticky top-0 z-10 border-b border-slate-200 bg-white/90 backdrop-blur">
      <div className="mx-auto flex max-w-[1600px] flex-wrap items-center gap-3 px-6 py-4">
        <div className="mr-auto">
          <h1 className="text-lg font-bold text-slate-900">MSBA Tracker</h1>
          <p className="text-xs text-slate-400">Compare, edite e organize seus programas de MSBA</p>
        </div>

        <button
          onClick={onToggleChart}
          className={`flex items-center gap-1.5 rounded-lg border px-3 py-2 text-sm font-medium transition-colors ${
            showChart ? "border-indigo-500 bg-indigo-50 text-indigo-700" : "border-slate-200 text-slate-600 hover:bg-slate-50"
          }`}
        >
          <BarChart3 size={16} /> Gráfico
        </button>

        <button
          onClick={onOpenFolders}
          className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
        >
          <FolderKanban size={16} /> Pastas
        </button>

        <div className="flex items-center rounded-lg border border-slate-200 p-1">
          <button
            onClick={() => onViewChange("cards")}
            className={`rounded-md p-1.5 ${view === "cards" ? "bg-indigo-600 text-white" : "text-slate-500 hover:bg-slate-100"}`}
            aria-label="Visualização em cards"
          >
            <LayoutGrid size={16} />
          </button>
          <button
            onClick={() => onViewChange("table")}
            className={`rounded-md p-1.5 ${view === "table" ? "bg-indigo-600 text-white" : "text-slate-500 hover:bg-slate-100"}`}
            aria-label="Visualização em tabela"
          >
            <Table2 size={16} />
          </button>
        </div>

        <button
          onClick={onOpenAddProgram}
          className="flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-2 text-sm font-medium text-white shadow-sm hover:bg-slate-800"
        >
          <Plus size={16} /> Adicionar Faculdade
        </button>

        <ExportButton programs={programs} folders={folders} />
        <AutoBackupButton programs={allPrograms} folders={folders} />
        <RestoreBackupButton />
      </div>
    </header>
  );
}
