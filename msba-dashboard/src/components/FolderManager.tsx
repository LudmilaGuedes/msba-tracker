import { useState } from "react";
import { Check, Pencil, Plus, Trash2, X } from "lucide-react";
import { Modal } from "./common/Modal";
import { useData } from "../context/DataContext";
import { FOLDER_COLORS } from "../constants/folderColors";
import type { MergedProgram } from "../types/program";

interface FolderManagerProps {
  open: boolean;
  onClose: () => void;
  programs: MergedProgram[];
}

export function FolderManager({ open, onClose, programs }: FolderManagerProps) {
  const { folders, createFolder, renameFolder, deleteFolder } = useData();
  const [newName, setNewName] = useState("");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");

  const countFor = (folderId: string) => programs.filter((p) => p.folderIds.includes(folderId)).length;

  const handleCreate = () => {
    const name = newName.trim();
    if (!name) return;
    createFolder(name);
    setNewName("");
  };

  return (
    <Modal open={open} onClose={onClose} title="Minhas pastas" widthClass="max-w-md">
      <div className="mb-4 flex gap-2">
        <input
          value={newName}
          onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleCreate()}
          placeholder="Ex.: Top Escolhas, Plano B, Aplicar Hoje..."
          className="flex-1 rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
        <button
          onClick={handleCreate}
          className="flex items-center gap-1 rounded-lg bg-indigo-600 px-3 py-2 text-sm font-medium text-white hover:bg-indigo-700"
        >
          <Plus size={16} /> Criar
        </button>
      </div>

      <div className="space-y-1.5">
        {folders.length === 0 && (
          <p className="py-6 text-center text-sm text-slate-400">Você ainda não criou nenhuma pasta.</p>
        )}
        {folders.map((f) => (
          <div key={f.id} className="flex items-center gap-2 rounded-lg border border-slate-100 px-3 py-2">
            <span className="h-3 w-3 shrink-0 rounded-full" style={{ backgroundColor: f.color }} />
            {editingId === f.id ? (
              <>
                <input
                  autoFocus
                  value={editingName}
                  onChange={(e) => setEditingName(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      renameFolder(f.id, editingName.trim() || f.name);
                      setEditingId(null);
                    }
                    if (e.key === "Escape") setEditingId(null);
                  }}
                  className="flex-1 rounded-md border border-indigo-200 px-2 py-1 text-sm outline-none"
                />
                <button
                  onClick={() => {
                    renameFolder(f.id, editingName.trim() || f.name);
                    setEditingId(null);
                  }}
                  className="text-emerald-600 hover:text-emerald-700"
                >
                  <Check size={16} />
                </button>
                <button onClick={() => setEditingId(null)} className="text-slate-400 hover:text-slate-600">
                  <X size={16} />
                </button>
              </>
            ) : (
              <>
                <span className="flex-1 truncate text-sm text-slate-700">{f.name}</span>
                <span className="text-xs text-slate-400">{countFor(f.id)} programa(s)</span>
                <button
                  onClick={() => {
                    setEditingId(f.id);
                    setEditingName(f.name);
                  }}
                  className="text-slate-400 hover:text-indigo-600"
                >
                  <Pencil size={14} />
                </button>
                <button
                  onClick={() => {
                    if (confirm(`Remover a pasta "${f.name}"? Os programas não serão excluídos.`)) deleteFolder(f.id);
                  }}
                  className="text-slate-400 hover:text-red-600"
                >
                  <Trash2 size={14} />
                </button>
              </>
            )}
          </div>
        ))}
      </div>

      <div className="mt-4 flex flex-wrap gap-1.5">
        {FOLDER_COLORS.map((c) => (
          <span key={c} className="h-4 w-4 rounded-full" style={{ backgroundColor: c }} />
        ))}
      </div>
    </Modal>
  );
}
