import { Building2, Clock, GraduationCap, MapPin, MessageSquare, TrendingUp } from "lucide-react";
import type { Folder, MergedProgram } from "../../types/program";
import { formatCellValue, isVerificar } from "../../utils/format";
import { StarRating } from "../common/StarRating";
import { FolderChipPicker } from "./FolderChipPicker";
import { useData } from "../../context/DataContext";
import { ResearchPromptButton } from "../common/ResearchPromptButton";

interface ProgramCardProps {
  program: MergedProgram;
  folders: Folder[];
  onOpenDetail: (program: MergedProgram) => void;
}

function Stat({ icon, label, value, verificar }: { icon: React.ReactNode; label: string; value: string; verificar?: boolean }) {
  return (
    <div className="flex items-center gap-2">
      <span className="text-slate-400">{icon}</span>
      <div className="min-w-0">
        <div className="text-[11px] uppercase tracking-wide text-slate-400">{label}</div>
        <div className={`truncate text-sm font-medium ${verificar ? "text-amber-600" : "text-slate-800"}`}>{value}</div>
      </div>
    </div>
  );
}

export function ProgramCard({ program, folders, onOpenDetail }: ProgramCardProps) {
  const { setRating, toggleProgramFolder } = useData();
  const v = program.values;

  return (
    <div
      onClick={() => onOpenDetail(program)}
      className="group flex cursor-pointer flex-col gap-3 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:-translate-y-0.5 hover:border-indigo-200 hover:shadow-md"
    >
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <h3 className="truncate text-base font-semibold text-slate-900">{v.universidade}</h3>
          <p className="truncate text-sm text-slate-500">{v.nomePrograma}</p>
        </div>
        {v.stemDesignated === "Sim" && (
          <span className="shrink-0 rounded-full bg-emerald-50 px-2 py-0.5 text-[11px] font-semibold text-emerald-700">
            STEM
          </span>
        )}
      </div>

      <div className="flex items-center gap-1.5 text-xs text-slate-500">
        <MapPin size={13} />
        <span>
          {v.cidade}, {v.estado}
        </span>
        <span className="text-slate-300">•</span>
        <Building2 size={13} />
        <span className="truncate">{v.tipo}</span>
      </div>

      <div className="grid grid-cols-2 gap-3 border-t border-slate-100 pt-3">
        <Stat
          icon={<Clock size={15} />}
          label="Duração"
          value={formatCellValue(v.duracaoMeses, "number") + (isVerificar(v.duracaoMeses) ? "" : " meses")}
          verificar={isVerificar(v.duracaoMeses)}
        />
        <Stat
          icon={<GraduationCap size={15} />}
          label="Tuition (Intl)"
          value={formatCellValue(v.tuitionTotalIntl, "currency")}
          verificar={isVerificar(v.tuitionTotalIntl)}
        />
        <Stat
          icon={<TrendingUp size={15} />}
          label="Empregabilidade"
          value={formatCellValue(v.taxaEmpregabilidade, "percent")}
          verificar={isVerificar(v.taxaEmpregabilidade)}
        />
        <Stat
          icon={<TrendingUp size={15} />}
          label="Salário Pós"
          value={formatCellValue(v.salarioMedioPos, "currency")}
          verificar={isVerificar(v.salarioMedioPos)}
        />
      </div>

      <div
        className="flex items-center justify-between gap-2 border-t border-slate-100 pt-3"
        onClick={(e) => e.stopPropagation()}
      >
        <StarRating value={program.rating} onChange={(r) => setRating(program.id, r)} size={16} />
        {program.comments && <MessageSquare size={15} className="text-indigo-400" />}
      </div>

      <div className="flex items-center justify-between gap-2" onClick={(e) => e.stopPropagation()}>
        <FolderChipPicker
          folders={folders}
          selectedIds={program.folderIds}
          onToggle={(fid) => toggleProgramFolder(program.id, fid)}
        />
        <ResearchPromptButton program={program} compact />
      </div>
    </div>
  );
}
