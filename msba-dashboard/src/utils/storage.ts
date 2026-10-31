import { COLUMNS } from "../types/columns";
import type { Folder, RawProgram, UserDataStore } from "../types/program";
import { DEFAULT_STRATEGY_SETTINGS, type StrategyDataStore, type StrategySettings } from "../types/strategy";
import { migrateLegacyIds } from "./programId";

const USER_DATA_KEY = "msba-tracker:user-data:v1";
const FOLDERS_KEY = "msba-tracker:folders:v1";
const CUSTOM_PROGRAMS_KEY = "msba-tracker:custom-programs:v1";
const STRATEGY_DATA_KEY = "msba-tracker:strategy-data:v1";
const STRATEGY_SETTINGS_KEY = "msba-tracker:strategy-settings:v1";

function safeParse<T>(raw: string | null, fallback: T): T {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

export function loadUserData(): UserDataStore {
  const store = safeParse<UserDataStore>(localStorage.getItem(USER_DATA_KEY), {});
  const migrated = migrateLegacyIds(store, USER_DATA_KEY);
  if (!migrated) return store;
  saveUserData(migrated);
  return migrated;
}

export function saveUserData(data: UserDataStore): void {
  try {
    localStorage.setItem(USER_DATA_KEY, JSON.stringify(data));
  } catch (err) {
    console.warn("Não foi possível salvar seus dados no navegador (localStorage cheio ou indisponível).", err);
  }
}

export function loadCustomPrograms(): RawProgram[] {
  const programs = safeParse<RawProgram[]>(localStorage.getItem(CUSTOM_PROGRAMS_KEY), []);
  // Programas salvos antes de uma coluna nova existir no schema não têm a
  // chave dela; sem dado, o valor é "Verificar" (nunca vazio).
  return programs.map((program) => {
    const filled: RawProgram = { ...program };
    for (const col of COLUMNS) {
      if (filled[col.key] === undefined) filled[col.key] = "Verificar";
    }
    return filled;
  });
}

export function saveCustomPrograms(programs: RawProgram[]): void {
  try {
    localStorage.setItem(CUSTOM_PROGRAMS_KEY, JSON.stringify(programs));
  } catch (err) {
    console.warn("Não foi possível salvar as faculdades adicionadas manualmente.", err);
  }
}

export function loadFolders(): Folder[] {
  return safeParse<Folder[]>(localStorage.getItem(FOLDERS_KEY), []);
}

export function saveFolders(folders: Folder[]): void {
  try {
    localStorage.setItem(FOLDERS_KEY, JSON.stringify(folders));
  } catch (err) {
    console.warn("Não foi possível salvar suas pastas no navegador.", err);
  }
}

export function loadStrategyData(): StrategyDataStore {
  const store = safeParse<StrategyDataStore>(localStorage.getItem(STRATEGY_DATA_KEY), {});
  const migrated = migrateLegacyIds(store, STRATEGY_DATA_KEY);
  if (!migrated) return store;
  saveStrategyData(migrated);
  return migrated;
}

export function saveStrategyData(data: StrategyDataStore): void {
  try {
    localStorage.setItem(STRATEGY_DATA_KEY, JSON.stringify(data));
  } catch (err) {
    console.warn("Não foi possível salvar os dados de estratégia de aplicação.", err);
  }
}

export function loadStrategySettings(): StrategySettings {
  return safeParse<StrategySettings>(localStorage.getItem(STRATEGY_SETTINGS_KEY), DEFAULT_STRATEGY_SETTINGS);
}

export function saveStrategySettings(settings: StrategySettings): void {
  try {
    localStorage.setItem(STRATEGY_SETTINGS_KEY, JSON.stringify(settings));
  } catch (err) {
    console.warn("Não foi possível salvar as configurações de estratégia de aplicação.", err);
  }
}
