import { ArrowDown, ArrowUp, ArrowUpDown, ExternalLink, MessageSquare } from "lucide-react";
import { useMemo, useState } from "react";
import { COLUMN_MAP } from "../../types/columns";
import type { Folder, MergedProgram } from "../../types/program";
import { useData } from "../../context/DataContext";
import { EditableCell } from "./EditableCell";
import { StarRating } from "../common/StarRating";
import { parseNumeric } from "../../utils/format";
import { FolderChipPicker } from "../cards/FolderChipPicker";
import { ResearchPromptButton } from "../common/ResearchPromptButton";

const TABLE_COLUMN_KEYS = [
  "universidade",
  "estado",
  "cidade",
  "duracaoMeses",
  "tuitionTotalIntl",
  "custoVidaTotal",
  "rankingUsNews",
  "taxaEmpregabilidade",
  "salarioMedioPos",
  "dificuldadeEntrada",
] as const;

type SortKey = (typeof TABLE_COLUMN_KEYS)[number] | "rating";
type SortDirection = "asc" | "desc";

interface DataTableProps {
  programs: MergedProgram[];
  folders: Folder[];
  onOpenDetail: (program: MergedProgram) => void;
}

export function DataTable({ programs, folders, onOpenDetail }: DataTableProps) {
  const { updateCell, setRating, toggleProgramFolder } = useData();
  const [sortKey, setSortKey] = useState<SortKey>("universidade");
  const [sortDir, setSortDir] = useState<SortDirection>("asc");

  const sorted = useMemo(() => {
    const copy = [...programs];
    copy.sort((a, b) => {
      const av = sortKey === "rating" ? a.rating : a.values[sortKey];
      const bv = sortKey === "rating" ? b.rating : b.values[sortKey];
      const an = parseNumeric(av);
      const bn = parseNumeric(bv);

      let cmp: number;
      if (an !== null && bn !== null) {
        cmp = an - bn;
      } else if (an !== null) {
        cmp = -1; // conhecidos antes de "Verificar"
      } else if (bn !== null) {
        cmp = 1;
      } else {
        cmp = String(av ?? "").localeCompare(String(bv ?? ""));
      }
      return sortDir === "asc" ? cmp : -cmp;
    });
    return copy;
  }, [programs, sortKey, sortDir]);

  const toggleSort = (key: SortKey) => {
    if (key === sortKey) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir("asc");
    }
  };

  const SortIcon = ({ col }: { col: SortKey }) => {
    if (col !== sortKey) return <ArrowUpDown size={13} className="text-slate-300" />;
    return sortDir === "asc" ? <ArrowUp size={13} className="text-indigo-600" /> : <ArrowDown size={13} className="text-indigo-600" />;
  };

  return (
    <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm">
      <table className="w-full min-w-[1200px] border-collapse text-sm">
        <thead>
          <tr className="border-b border-slate-200 bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
            {TABLE_COLUMN_KEYS.map((key) => (
              <th key={key} className="whitespace-nowrap px-3 py-3 text-left font-semibold">
                <button className="flex items-center gap-1 hover:text-slate-900" onClick={() => toggleSort(key)}>
                  {COLUMN_MAP[key].header}
                  <SortIcon col={key} />
                </button>
              </th>
            ))}
            <th className="whitespace-nowrap px-3 py-3 text-left font-semibold">
              <button className="flex items-center gap-1 hover:text-slate-900" onClick={() => toggleSort("rating")}>
                Nota <SortIcon col="rating" />
              </button>
            </th>
            <th className="whitespace-nowrap px-3 py-3 text-left font-semibold">Pastas</th>
            <th className="whitespace-nowrap px-3 py-3 text-left font-semibold">Notas</th>
            <th className="px-3 py-3" />
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {sorted.map((program) => (
            <tr key={program.id} className="transition-colors hover:bg-slate-50/80">
              {TABLE_COLUMN_KEYS.map((key) => (
                <td key={key} className="px-2 py-1.5" style={{ minWidth: COLUMN_MAP[key].width }}>
                  <EditableCell
                    value={program.values[key]}
                    type={COLUMN_MAP[key].type}
                    options={COLUMN_MAP[key].options}
                    edited={program.editedKeys.has(key)}
                    onSave={(v) => updateCell(program.id, key, v)}
                  />
                </td>
              ))}
              <td className="px-3 py-1.5">
                <StarRating value={program.rating} onChange={(r) => setRating(program.id, r)} size={15} />
              </td>
              <td className="px-3 py-1.5">
                <FolderChipPicker
                  folders={folders}
                  selectedIds={program.folderIds}
                  onToggle={(fid) => toggleProgramFolder(program.id, fid)}
                  compact
                />
              </td>
              <td className="px-3 py-1.5 text-slate-400">
                {program.comments ? <MessageSquare size={16} className="text-indigo-500" /> : null}
              </td>
              <td className="px-3 py-1.5">
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => onOpenDetail(program)}
                    className="flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium text-indigo-600 hover:bg-indigo-50"
                  >
                    Detalhes <ExternalLink size={13} />
                  </button>
                  <ResearchPromptButton program={program} compact />
                </div>
              </td>
            </tr>
          ))}
          {sorted.length === 0 && (
            <tr>
              <td colSpan={TABLE_COLUMN_KEYS.length + 4} className="px-4 py-10 text-center text-slate-400">
                Nenhum programa encontrado com os filtros atuais.
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
