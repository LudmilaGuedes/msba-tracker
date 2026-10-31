import { useCallback, useEffect, useRef, useState } from "react";
import type { Folder, MergedProgram } from "../types/program";
import { useToast } from "../context/ToastContext";
import {
  chooseBackupFile,
  forgetBackupFile,
  getBackupHandle,
  getLastBackupAt,
  hasReadWritePermission,
  isBackupDue,
  isFileSystemAccessSupported,
  writeBackupNow,
} from "../utils/autoBackup";

export type AutoBackupStatus = "unsupported" | "unconfigured" | "needs-permission" | "ready" | "checking";

// Confere a cada 15min (enquanto o app está aberto) se já passou 24h desde
// o último backup — não é uma tarefa agendada do SO, só roda com a aba viva.
const CHECK_INTERVAL_MS = 15 * 60 * 1000;

export function useAutoBackup(programs: MergedProgram[], folders: Folder[]) {
  const { showToast } = useToast();
  const [status, setStatus] = useState<AutoBackupStatus>(isFileSystemAccessSupported() ? "checking" : "unsupported");
  const [fileName, setFileName] = useState<string | null>(null);
  const [lastBackupAt, setLastBackupAt] = useState<number | null>(getLastBackupAt());
  const handleRef = useRef<FileSystemFileHandle | null>(null);

  // Sempre reflete os dados mais recentes sem precisar recriar os callbacks/intervalos.
  const dataRef = useRef({ programs, folders });
  dataRef.current = { programs, folders };

  const runBackup = useCallback(
    async (handle: FileSystemFileHandle, silent: boolean) => {
      try {
        await writeBackupNow(handle, dataRef.current.programs, dataRef.current.folders);
        setLastBackupAt(getLastBackupAt());
        setStatus("ready");
        if (!silent) showToast(`Backup salvo em "${handle.name}".`);
      } catch (err) {
        console.warn("Falha ao gravar backup automático.", err);
        if (!silent) showToast("Não foi possível salvar o backup agora.");
      }
    },
    [showToast],
  );

  useEffect(() => {
    if (!isFileSystemAccessSupported()) return;
    let cancelled = false;
    (async () => {
      const handle = await getBackupHandle();
      if (cancelled) return;
      if (!handle) {
        setStatus("unconfigured");
        return;
      }
      handleRef.current = handle;
      setFileName(handle.name);
      const granted = await hasReadWritePermission(handle, false);
      if (cancelled) return;
      if (!granted) {
        setStatus("needs-permission");
        return;
      }
      setStatus("ready");
      if (isBackupDue()) runBackup(handle, true);
    })();
    return () => {
      cancelled = true;
    };
  }, [runBackup]);

  useEffect(() => {
    if (!isFileSystemAccessSupported()) return;
    const id = setInterval(() => {
      if (handleRef.current && status === "ready" && isBackupDue()) {
        runBackup(handleRef.current, true);
      }
    }, CHECK_INTERVAL_MS);
    return () => clearInterval(id);
  }, [status, runBackup]);

  const configure = useCallback(async () => {
    try {
      const handle = await chooseBackupFile();
      handleRef.current = handle;
      setFileName(handle.name);
      setStatus("ready");
      await runBackup(handle, false);
    } catch (err) {
      if ((err as DOMException)?.name !== "AbortError") {
        console.warn("Não foi possível configurar o backup automático.", err);
      }
    }
  }, [runBackup]);

  const grantPermission = useCallback(async () => {
    if (!handleRef.current) return;
    const granted = await hasReadWritePermission(handleRef.current, true);
    if (granted) {
      setStatus("ready");
      if (isBackupDue()) runBackup(handleRef.current, true);
    }
  }, [runBackup]);

  const backupNow = useCallback(async () => {
    if (!handleRef.current) return;
    await runBackup(handleRef.current, false);
  }, [runBackup]);

  const disable = useCallback(async () => {
    await forgetBackupFile();
    handleRef.current = null;
    setFileName(null);
    setLastBackupAt(null);
    setStatus("unconfigured");
  }, []);

  return { status, fileName, lastBackupAt, configure, grantPermission, backupNow, disable };
}
