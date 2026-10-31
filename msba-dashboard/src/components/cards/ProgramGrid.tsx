import type { Folder, MergedProgram } from "../../types/program";
import { ProgramCard } from "./ProgramCard";

interface ProgramGridProps {
  programs: MergedProgram[];
  folders: Folder[];
  onOpenDetail: (program: MergedProgram) => void;
}

export function ProgramGrid({ programs, folders, onOpenDetail }: ProgramGridProps) {
  if (programs.length === 0) {
    return (
      <div className="rounded-2xl border border-dashed border-slate-300 bg-white py-16 text-center text-slate-400">
        Nenhum programa encontrado com os filtros atuais.
      </div>
    );
  }

  return (
    <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
      {programs.map((program) => (
        <ProgramCard key={program.id} program={program} folders={folders} onOpenDetail={onOpenDetail} />
      ))}
    </div>
  );
}
