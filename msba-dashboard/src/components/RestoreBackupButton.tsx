import { useRef } from "react";
import { Upload } from "lucide-react";
import { useData } from "../context/DataContext";
import { useToast } from "../context/ToastContext";
import { parseBackupFile } from "../utils/restoreBackup";

// Restaura notas/comentários/pastas/edições a partir de um arquivo de
// backup (o mesmo formato gerado pelo "Exportar Meus Dados" ou pelo backup
// automático). Usa um <input type="file"> comum de propósito — diferente
// do backup automático, isso precisa funcionar mesmo se o navegador não
// suporta a File System Access API, e mesmo se o handle salvo no
// IndexedDB tiver sido perdido junto com o resto dos dados.
export function RestoreBackupButton() {
  const { restoreFromBackup } = useData();
  const { showToast } = useToast();
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;

    const confirmed = confirm(
      `Restaurar a partir de "${file.name}"?\n\nIsso vai substituir suas notas, comentários e pastas atuais pelos dados salvos nesse arquivo, para os programas que baterem.`,
    );
    if (!confirmed) return;

    try {
      const rows = await parseBackupFile(file);
      if (rows.length === 0) {
        showToast("O arquivo não tem nenhuma linha de dados reconhecível.");
        return;
      }
      const summary = restoreFromBackup(rows);
      const parts = [`${summary.matchedOriginal} da planilha original`];
      if (summary.matchedCustom > 0) parts.push(`${summary.matchedCustom} adicionados por você`);
      if (summary.recreated > 0) parts.push(`${summary.recreated} recriados do zero`);
      showToast(`Restaurado: ${parts.join(", ")}.`);
    } catch (err) {
      console.warn("Não foi possível restaurar o backup.", err);
      showToast("Não foi possível ler esse arquivo como backup.");
    }
  };

  return (
    <>
      <button
        onClick={() => inputRef.current?.click()}
        className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
      >
        <Upload size={16} /> Restaurar backup
      </button>
      <input ref={inputRef} type="file" accept=".xlsx" onChange={handleFile} className="hidden" />
    </>
  );
}
