import { useEffect, useMemo, useState } from "react";
import { DataProvider, useData } from "./context/DataContext";
import { StrategyProvider } from "./context/StrategyContext";
import { ToastProvider } from "./context/ToastContext";
import { DEFAULT_FILTERS, type FilterState } from "./types/filters";
import { filterPrograms } from "./utils/filterPrograms";
import type { MergedProgram } from "./types/program";
import { Header, type ViewMode } from "./components/Header";
import { FilterPanel } from "./components/filters/FilterPanel";
import { DataTable } from "./components/table/DataTable";
import { ProgramGrid } from "./components/cards/ProgramGrid";
import { ProgramDetailModal } from "./components/ProgramDetailModal";
import { FolderManager } from "./components/FolderManager";
import { AddProgramModal } from "./components/AddProgramModal";
import { CostVsSalaryChart } from "./components/charts/CostVsSalaryChart";
import { ApplicationStrategyTab } from "./components/strategy/ApplicationStrategyTab";
import { Loader2, LayoutDashboard, ListChecks } from "lucide-react";

type MainTab = "dashboard" | "strategy";

function Dashboard() {
  const { loading, error, programs, folders, deleteCustomProgram } = useData();
  const [mainTab, setMainTab] = useState<MainTab>("dashboard");
  const [view, setView] = useState<ViewMode>("cards");
  const [filters, setFilters] = useState<FilterState>(DEFAULT_FILTERS);
  const [detailProgram, setDetailProgram] = useState<MergedProgram | null>(null);
  const [foldersOpen, setFoldersOpen] = useState(false);
  const [addProgramOpen, setAddProgramOpen] = useState(false);
  const [pendingDetailId, setPendingDetailId] = useState<string | null>(null);
  const [showChart, setShowChart] = useState(false);

  const filtered = useMemo(() => filterPrograms(programs, filters), [programs, filters]);

  useEffect(() => {
    if (!pendingDetailId) return;
    const found = programs.find((p) => p.id === pendingDetailId);
    if (found) {
      setDetailProgram(found);
      setPendingDetailId(null);
    }
  }, [pendingDetailId, programs]);

  if (loading) {
    return (
      <div className="flex h-screen flex-col items-center justify-center gap-3 text-slate-500">
        <Loader2 className="animate-spin" size={28} />
        <p className="text-sm">Carregando programas...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex h-screen flex-col items-center justify-center gap-2 px-4 text-center text-slate-600">
        <p className="text-lg font-semibold text-red-600">Não foi possível carregar os dados</p>
        <p className="text-sm text-slate-500">{error}</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <Header
        view={view}
        onViewChange={setView}
        showChart={showChart}
        onToggleChart={() => setShowChart((v) => !v)}
        onOpenFolders={() => setFoldersOpen(true)}
        onOpenAddProgram={() => setAddProgramOpen(true)}
        programs={filtered}
        allPrograms={programs}
        folders={folders}
      />

      <div className="mx-auto flex max-w-[1600px] gap-1 px-6 pt-4">
        <button
          onClick={() => setMainTab("dashboard")}
          className={`flex items-center gap-1.5 rounded-t-lg border-b-2 px-4 py-2 text-sm font-medium transition-colors ${
            mainTab === "dashboard"
              ? "border-indigo-600 text-indigo-700"
              : "border-transparent text-slate-500 hover:text-slate-700"
          }`}
        >
          <LayoutDashboard size={15} /> Programas
        </button>
        <button
          onClick={() => setMainTab("strategy")}
          className={`flex items-center gap-1.5 rounded-t-lg border-b-2 px-4 py-2 text-sm font-medium transition-colors ${
            mainTab === "strategy"
              ? "border-indigo-600 text-indigo-700"
              : "border-transparent text-slate-500 hover:text-slate-700"
          }`}
        >
          <ListChecks size={15} /> Estratégia de Aplicação
        </button>
      </div>

      {mainTab === "dashboard" ? (
        <main className="mx-auto flex max-w-[1600px] flex-col gap-5 px-6 py-6 lg:flex-row">
          <FilterPanel
            programs={programs}
            folders={folders}
            filters={filters}
            onChange={setFilters}
            resultCount={filtered.length}
          />

          <div className="flex-1 space-y-5">
            {showChart && <CostVsSalaryChart programs={filtered} />}

            {view === "cards" ? (
              <ProgramGrid programs={filtered} folders={folders} onOpenDetail={setDetailProgram} />
            ) : (
              <DataTable programs={filtered} folders={folders} onOpenDetail={setDetailProgram} />
            )}
          </div>
        </main>
      ) : (
        <main className="mx-auto max-w-[1600px] px-6 py-6">
          <ApplicationStrategyTab onOpenDetail={setDetailProgram} />
        </main>
      )}

      <ProgramDetailModal
        program={detailProgram}
        folders={folders}
        onClose={() => setDetailProgram(null)}
        onDelete={(id) => {
          deleteCustomProgram(id);
          setDetailProgram(null);
        }}
      />
      <FolderManager open={foldersOpen} onClose={() => setFoldersOpen(false)} programs={programs} />
      <AddProgramModal
        open={addProgramOpen}
        onClose={() => setAddProgramOpen(false)}
        onCreated={(id) => {
          setAddProgramOpen(false);
          setPendingDetailId(id);
        }}
      />
    </div>
  );
}

export default function App() {
  return (
    <ToastProvider>
      <DataProvider>
        <StrategyProvider>
          <Dashboard />
        </StrategyProvider>
      </DataProvider>
    </ToastProvider>
  );
}
