import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from "react";
import { loadProgramsFromExcel } from "../utils/dataParser";
import {
  loadCustomPrograms,
  loadFolders,
  loadUserData,
  saveCustomPrograms,
  saveFolders,
  saveUserData,
} from "../utils/storage";
import { COLUMNS, NUMERIC_TYPES } from "../types/columns";
import { FOLDER_COLORS } from "../constants/folderColors";
import { computeCustoVidaTotal, CUSTO_VIDA_TRIGGER_KEYS } from "../utils/derived";
import type { BackupRow } from "../utils/restoreBackup";
import type { CellValue, Folder, MergedProgram, ProgramUserData, RawProgram, UserDataStore } from "../types/program";

export interface RestoreSummary {
  total: number;
  matchedOriginal: number;
  matchedCustom: number;
  recreated: number;
}

interface DataContextValue {
  loading: boolean;
  error: string | null;
  rawPrograms: RawProgram[];
  programs: MergedProgram[];
  folders: Folder[];
  updateCell: (programId: string, key: string, value: CellValue) => void;
  setRating: (programId: string, rating: number) => void;
  setComments: (programId: string, comments: string) => void;
  createFolder: (name: string, color?: string) => Folder;
  ensureFolders: (names: string[]) => void;
  renameFolder: (folderId: string, name: string) => void;
  deleteFolder: (folderId: string) => void;
  toggleProgramFolder: (programId: string, folderId: string) => void;
  resetProgramEdits: (programId: string) => void;
  addCustomProgram: (fields: Record<string, CellValue>) => RawProgram;
  deleteCustomProgram: (programId: string) => void;
  restoreFromBackup: (rows: BackupRow[]) => RestoreSummary;
}

const DataContext = createContext<DataContextValue | null>(null);

function mergeProgram(raw: RawProgram, userData: ProgramUserData | undefined, isCustom: boolean): MergedProgram {
  const values: Record<string, CellValue> = { ...raw };
  const editedKeys = new Set<string>();
  if (userData?.edits) {
    for (const [key, value] of Object.entries(userData.edits)) {
      values[key] = value;
      editedKeys.add(key);
    }
  }
  return {
    id: raw.id,
    raw,
    values,
    editedKeys,
    rating: userData?.rating ?? 0,
    comments: userData?.comments ?? "",
    folderIds: userData?.folderIds ?? [],
    isCustom,
  };
}

export function DataProvider({ children }: { children: ReactNode }) {
  const [rawPrograms, setRawPrograms] = useState<RawProgram[]>([]);
  const [customPrograms, setCustomPrograms] = useState<RawProgram[]>([]);
  const [userData, setUserData] = useState<UserDataStore>({});
  const [folders, setFolders] = useState<Folder[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setUserData(loadUserData());
    setFolders(loadFolders());
    setCustomPrograms(loadCustomPrograms());
    loadProgramsFromExcel()
      .then(setRawPrograms)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  useEffect(() => {
    if (!loading) saveUserData(userData);
  }, [userData, loading]);

  useEffect(() => {
    if (!loading) saveFolders(folders);
  }, [folders, loading]);

  useEffect(() => {
    if (!loading) saveCustomPrograms(customPrograms);
  }, [customPrograms, loading]);

  const patchUserData = useCallback((programId: string, patch: Partial<ProgramUserData>) => {
    setUserData((prev) => ({
      ...prev,
      [programId]: {
        ...prev[programId],
        ...patch,
        updatedAt: Date.now(),
      },
    }));
  }, []);

  const updateCell = useCallback(
    (programId: string, key: string, value: CellValue) => {
      setUserData((prev) => {
        const nextEdits: Record<string, CellValue> = { ...prev[programId]?.edits, [key]: value };

        // Editar um componente de custo de vida (aluguel, mercado, seguro,
        // utilities, etc.) ou a duração do programa recalcula
        // automaticamente 'Custo Vida Total (programa)' — mesmo ponto único
        // por onde toda edição (tabela e modal de detalhe) passa.
        if ((CUSTO_VIDA_TRIGGER_KEYS as readonly string[]).includes(key)) {
          const raw = rawPrograms.find((p) => p.id === programId) ?? customPrograms.find((p) => p.id === programId);
          if (raw) {
            const merged = { ...raw, ...nextEdits };
            const total = computeCustoVidaTotal(merged);
            if (total !== null) nextEdits.custoVidaTotal = total;
          }
        }

        return {
          ...prev,
          [programId]: {
            ...prev[programId],
            edits: nextEdits,
            updatedAt: Date.now(),
          },
        };
      });
    },
    [rawPrograms, customPrograms],
  );

  const setRating = useCallback(
    (programId: string, rating: number) => patchUserData(programId, { rating }),
    [patchUserData],
  );

  const setComments = useCallback(
    (programId: string, comments: string) => patchUserData(programId, { comments }),
    [patchUserData],
  );

  const createFolder = useCallback((name: string, color?: string): Folder => {
    const folder: Folder = {
      id: `f_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      name,
      color: color ?? FOLDER_COLORS[Math.floor(Math.random() * FOLDER_COLORS.length)],
      createdAt: Date.now(),
    };
    setFolders((prev) => [...prev, folder]);
    return folder;
  }, []);

  /**
   * Cria as pastas que ainda não existem (comparação case-insensitive), em
   * uma única atualização funcional de estado — necessário para ser seguro
   * sob o duplo-disparo de efeitos do StrictMode em dev, onde duas chamadas
   * separadas de createFolder() na mesma renderização veriam o mesmo
   * `folders` desatualizado e duplicariam a pasta.
   */
  const ensureFolders = useCallback((names: string[]) => {
    setFolders((prev) => {
      const existing = new Set(prev.map((f) => f.name.trim().toLowerCase()));
      const toAdd = names
        .filter((name) => !existing.has(name.trim().toLowerCase()))
        .map((name, i) => ({
          id: `f_${Date.now()}_${i}_${Math.random().toString(36).slice(2, 7)}`,
          name,
          color: FOLDER_COLORS[Math.floor(Math.random() * FOLDER_COLORS.length)],
          createdAt: Date.now(),
        }));
      return toAdd.length ? [...prev, ...toAdd] : prev;
    });
  }, []);

  const renameFolder = useCallback((folderId: string, name: string) => {
    setFolders((prev) => prev.map((f) => (f.id === folderId ? { ...f, name } : f)));
  }, []);

  const deleteFolder = useCallback((folderId: string) => {
    setFolders((prev) => prev.filter((f) => f.id !== folderId));
    setUserData((prev) => {
      const next: UserDataStore = {};
      for (const [pid, data] of Object.entries(prev)) {
        next[pid] = { ...data, folderIds: (data.folderIds ?? []).filter((id) => id !== folderId) };
      }
      return next;
    });
  }, []);

  const toggleProgramFolder = useCallback((programId: string, folderId: string) => {
    setUserData((prev) => {
      const current = prev[programId]?.folderIds ?? [];
      const next = current.includes(folderId)
        ? current.filter((id) => id !== folderId)
        : [...current, folderId];
      return {
        ...prev,
        [programId]: { ...prev[programId], folderIds: next, updatedAt: Date.now() },
      };
    });
  }, []);

  const resetProgramEdits = useCallback((programId: string) => {
    setUserData((prev) => {
      const { [programId]: _removed, ...rest } = prev;
      return rest;
    });
  }, []);

  const addCustomProgram = useCallback((fields: Record<string, CellValue>): RawProgram => {
    const program: RawProgram = { id: `custom_${Date.now()}_${Math.random().toString(36).slice(2, 7)}` };
    for (const col of COLUMNS) {
      const raw = fields[col.key];
      const trimmed = typeof raw === "string" ? raw.trim() : raw;
      const isBlank = trimmed === undefined || trimmed === null || trimmed === "";

      if (isBlank) {
        program[col.key] = "Verificar";
        continue;
      }

      if (NUMERIC_TYPES.includes(col.type) && typeof trimmed === "string" && trimmed.toLowerCase() !== "verificar") {
        const n = Number(trimmed.replace(/[^0-9.\-]/g, ""));
        program[col.key] = Number.isFinite(n) && trimmed.replace(/[^0-9.\-]/g, "") !== "" ? n : trimmed;
      } else {
        program[col.key] = trimmed;
      }
    }
    setCustomPrograms((prev) => [...prev, program]);
    return program;
  }, []);

  const deleteCustomProgram = useCallback((programId: string) => {
    setCustomPrograms((prev) => prev.filter((p) => p.id !== programId));
    setUserData((prev) => {
      const { [programId]: _removed, ...rest } = prev;
      return rest;
    });
  }, []);

  const restoreFromBackup = useCallback(
    (rows: BackupRow[]): RestoreSummary => {
      const norm = (s: string) => s.trim().toLowerCase();
      const currentPrograms = [
        ...rawPrograms.map((raw) => mergeProgram(raw, userData[raw.id], false)),
        ...customPrograms.map((raw) => mergeProgram(raw, userData[raw.id], true)),
      ];

      const folderByName = new Map(folders.map((f) => [norm(f.name), f]));
      const newFolders: Folder[] = [];
      const resolveFolderId = (name: string): string => {
        const key = norm(name);
        let existing = folderByName.get(key);
        if (!existing) {
          existing = {
            id: `f_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
            name: name.trim(),
            color: FOLDER_COLORS[Math.floor(Math.random() * FOLDER_COLORS.length)],
            createdAt: Date.now(),
          };
          folderByName.set(key, existing);
          newFolders.push(existing);
        }
        return existing.id;
      };

      const findMatch = (row: BackupRow) =>
        currentPrograms.find(
          (p) =>
            norm(String(p.raw.universidade ?? "")) === norm(row.universidade) &&
            norm(String(p.raw.estado ?? "")) === norm(row.estado) &&
            norm(String(p.raw.cidade ?? "")) === norm(row.cidade),
        );

      const nextUserData: UserDataStore = { ...userData };
      const newCustomPrograms: RawProgram[] = [];
      let matchedOriginal = 0;
      let matchedCustom = 0;
      let recreated = 0;

      for (const row of rows) {
        const folderIds = row.folderNames.map(resolveFolderId);
        const match = findMatch(row);

        if (match) {
          const edits: Record<string, CellValue> = {};
          for (const col of COLUMNS) {
            const backupVal = row.values[col.key];
            if (backupVal === undefined) continue;
            const originalVal = match.raw[col.key];
            if (String(backupVal) !== String(originalVal ?? "")) {
              edits[col.key] = backupVal;
            }
          }
          nextUserData[match.id] = {
            edits: Object.keys(edits).length ? edits : undefined,
            rating: row.rating || undefined,
            comments: row.comments || undefined,
            folderIds,
            updatedAt: Date.now(),
          };
          if (match.isCustom) matchedCustom++;
          else matchedOriginal++;
        } else {
          const id = `custom_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
          const program: RawProgram = { id };
          for (const col of COLUMNS) program[col.key] = row.values[col.key] ?? "Verificar";
          newCustomPrograms.push(program);
          nextUserData[id] = {
            rating: row.rating || undefined,
            comments: row.comments || undefined,
            folderIds,
            updatedAt: Date.now(),
          };
          recreated++;
        }
      }

      if (newFolders.length) setFolders((prev) => [...prev, ...newFolders]);
      if (newCustomPrograms.length) setCustomPrograms((prev) => [...prev, ...newCustomPrograms]);
      setUserData(nextUserData);

      return { total: rows.length, matchedOriginal, matchedCustom, recreated };
    },
    [rawPrograms, customPrograms, userData, folders],
  );

  const programs = useMemo<MergedProgram[]>(() => {
    const original = rawPrograms.map((raw) => mergeProgram(raw, userData[raw.id], false));
    const custom = customPrograms.map((raw) => mergeProgram(raw, userData[raw.id], true));
    return [...original, ...custom];
  }, [rawPrograms, customPrograms, userData]);

  const value: DataContextValue = {
    loading,
    error,
    rawPrograms,
    programs,
    folders,
    updateCell,
    setRating,
    setComments,
    createFolder,
    ensureFolders,
    renameFolder,
    deleteFolder,
    toggleProgramFolder,
    resetProgramEdits,
    addCustomProgram,
    deleteCustomProgram,
    restoreFromBackup,
  };

  return <DataContext.Provider value={value}>{children}</DataContext.Provider>;
}

export function useData(): DataContextValue {
  const ctx = useContext(DataContext);
  if (!ctx) throw new Error("useData deve ser usado dentro de um DataProvider");
  return ctx;
}
