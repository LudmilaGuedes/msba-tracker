import { useState } from "react";
import { Modal } from "./common/Modal";
import { COLUMN_GROUPS, COLUMNS } from "../types/columns";
import { useData } from "../context/DataContext";
import type { CellValue } from "../types/program";

interface AddProgramModalProps {
  open: boolean;
  onClose: () => void;
  onCreated: (programId: string) => void;
}

const EMPTY_FORM: Record<string, string> = Object.fromEntries(COLUMNS.map((c) => [c.key, ""]));

export function AddProgramModal({ open, onClose, onCreated }: AddProgramModalProps) {
  const { addCustomProgram } = useData();
  const [form, setForm] = useState<Record<string, string>>(EMPTY_FORM);
  const [error, setError] = useState<string | null>(null);

  const setField = (key: string, value: string) => setForm((prev) => ({ ...prev, [key]: value }));

  const handleClose = () => {
    setForm(EMPTY_FORM);
    setError(null);
    onClose();
  };

  const handleSubmit = () => {
    if (!form.universidade.trim()) {
      setError("Informe ao menos o nome da universidade.");
      return;
    }
    const created = addCustomProgram(form as Record<string, CellValue>);
    setForm(EMPTY_FORM);
    setError(null);
    onCreated(created.id);
  };

  return (
    <Modal open={open} onClose={handleClose} title="Adicionar Faculdade Manualmente" widthClass="max-w-4xl">
      <p className="mb-5 text-sm text-slate-500">
        Preencha o que souber. Campos deixados em branco são salvos automaticamente como <strong>"Verificar"</strong>.
      </p>

      {error && (
        <div className="mb-4 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">{error}</div>
      )}

      <div className="space-y-6">
        {COLUMN_GROUPS.map((group) => (
          <div key={group}>
            <h4 className="mb-2 text-sm font-semibold text-slate-700">{group}</h4>
            <div className="grid grid-cols-1 gap-x-4 gap-y-3 rounded-xl border border-slate-100 p-3 sm:grid-cols-2">
              {COLUMNS.filter((c) => c.group === group).map((col) => (
                <div key={col.key} className="min-w-0">
                  <label className="mb-1 block text-[11px] text-slate-400">
                    {col.header}
                    {col.key === "universidade" && <span className="text-red-500"> *</span>}
                  </label>
                  {col.type === "enum" && col.options ? (
                    <select
                      value={form[col.key]}
                      onChange={(e) => setField(col.key, e.target.value)}
                      className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
                    >
                      <option value="">Verificar</option>
                      {col.options.map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  ) : col.type === "longtext" ? (
                    <textarea
                      value={form[col.key]}
                      onChange={(e) => setField(col.key, e.target.value)}
                      placeholder="Verificar"
                      rows={2}
                      className="w-full resize-none rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
                    />
                  ) : (
                    <input
                      type="text"
                      value={form[col.key]}
                      onChange={(e) => setField(col.key, e.target.value)}
                      placeholder="Verificar"
                      className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
                    />
                  )}
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      <div className="mt-6 flex justify-end gap-2 border-t border-slate-100 pt-4">
        <button
          onClick={handleClose}
          className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
        >
          Cancelar
        </button>
        <button
          onClick={handleSubmit}
          className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-700"
        >
          Salvar Novo Programa
        </button>
      </div>
    </Modal>
  );
}
