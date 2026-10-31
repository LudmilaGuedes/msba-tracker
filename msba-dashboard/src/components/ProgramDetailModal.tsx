import { COLUMN_GROUPS, COLUMNS } from "../types/columns";
import type { Folder, MergedProgram } from "../types/program";
import { useData } from "../context/DataContext";
import { Modal } from "./common/Modal";
import { StarRating } from "./common/StarRating";
import { EditableCell } from "./table/EditableCell";
import { FolderChipPicker } from "./cards/FolderChipPicker";
import { RoundBadge } from "./common/RoundBadge";
import { classifyLeadTimeBucket, parseRounds } from "../utils/strategyParsing";
import { RotateCcw, Trash2 } from "lucide-react";

interface ProgramDetailModalProps {
  program: MergedProgram | null;
  folders: Folder[];
  onClose: () => void;
  onDelete: (programId: string) => void;
}

export function ProgramDetailModal({ program, folders, onClose, onDelete }: ProgramDetailModalProps) {
  const { updateCell, setRating, setComments, toggleProgramFolder, resetProgramEdits } = useData();

  if (!program) return null;

  return (
    <Modal
      open={!!program}
      onClose={onClose}
      widthClass="max-w-4xl"
      title={
        <div>
          <div className="flex items-center gap-2">
            {String(program.values.universidade)}
            {program.isCustom && (
              <span className="rounded-full bg-indigo-50 px-2 py-0.5 text-[11px] font-semibold text-indigo-600">
                Adicionado manualmente
              </span>
            )}
          </div>
          <div className="text-sm font-normal text-slate-500">{String(program.values.nomePrograma)}</div>
        </div>
      }
    >
      <div className="mb-6 flex flex-wrap items-center gap-6 rounded-xl bg-slate-50 p-4">
        <div>
          <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Sua nota</div>
          <StarRating value={program.rating} onChange={(r) => setRating(program.id, r)} size={22} />
        </div>
        <div>
          <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Suas pastas</div>
          <FolderChipPicker
            folders={folders}
            selectedIds={program.folderIds}
            onToggle={(fid) => toggleProgramFolder(program.id, fid)}
          />
        </div>
        <div className="ml-auto flex items-center gap-2">
          {program.editedKeys.size > 0 && (
            <button
              onClick={() => resetProgramEdits(program.id)}
              className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-100"
            >
              <RotateCcw size={13} /> Restaurar dados originais
            </button>
          )}
          {program.isCustom && (
            <button
              onClick={() => {
                if (confirm("Excluir esta faculdade adicionada manualmente? Esta ação não pode ser desfeita.")) {
                  onDelete(program.id);
                }
              }}
              className="flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-1.5 text-xs font-medium text-red-600 hover:bg-red-50"
            >
              <Trash2 size={13} /> Excluir programa
            </button>
          )}
        </div>
      </div>

      <div className="mb-6">
        <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-slate-400">
          Meus comentários
        </label>
        <textarea
          defaultValue={program.comments}
          onBlur={(e) => setComments(program.id, e.target.value)}
          placeholder="Ex.: Gostei do campus, mas a taxa de aplicação é alta..."
          rows={3}
          className="w-full resize-none rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </div>

      <div className="space-y-6">
        {COLUMN_GROUPS.map((group) => (
          <div key={group}>
            <h4 className="mb-2 text-sm font-semibold text-slate-700">{group}</h4>
            <div className="grid grid-cols-1 gap-x-4 gap-y-2 rounded-xl border border-slate-100 p-3 sm:grid-cols-2">
              {COLUMNS.filter((c) => c.group === group).map((col) => (
                <div key={col.key} className="min-w-0">
                  <div className="mb-0.5 text-[11px] text-slate-400">{col.header}</div>
                  {col.key === "deadlineAplicacao" && (
                    <div className="mb-1 flex flex-wrap items-center gap-1.5">
                      <span className="text-sm font-medium text-slate-700">
                        {parseRounds(program.values.deadlineAplicacao)[0].dateText}
                      </span>
                      <RoundBadge
                        bucket={classifyLeadTimeBucket(program.values.deadlineAplicacao)}
                      />
                    </div>
                  )}
                  <EditableCell
                    value={program.values[col.key]}
                    type={col.type}
                    options={col.options}
                    edited={program.editedKeys.has(col.key)}
                    onSave={(v) => updateCell(program.id, col.key, v)}
                  />
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </Modal>
  );
}
