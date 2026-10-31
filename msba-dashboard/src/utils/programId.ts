// Id estável de um programa da planilha: nome + estado + cidade, normalizados.
// Antes o id era a posição da linha (p0, p1, ...) e remover uma linha deslocava
// as notas/comentários/pastas de todas as linhas seguintes para a faculdade de
// baixo. Agora remover ou reordenar linhas não afeta mais o overlay do usuário;
// só renomear nome/estado/cidade de uma linha muda o id dela.
import { LEGACY_ID_MAP } from "./legacyIdMap";

const norm = (v: unknown) =>
  String(v ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

export function programId(universidade: unknown, estado: unknown, cidade: unknown): string {
  return `prog:${norm(universidade)}|${norm(estado)}|${norm(cidade)}`;
}

const LEGACY_ID = /^p\d+$/;

/**
 * Converte um store salvo com ids antigos por posição para ids estáveis, usando
 * LEGACY_ID_MAP. Antes de converter, guarda o JSON original em
 * `${storageKey}:pre-id-migration` (só na primeira vez) para poder desfazer.
 * Ids antigos sem entrada no mapa são mantidos como estão, para não perder nada.
 * Devolve null quando não há nada a migrar.
 */
export function migrateLegacyIds<T>(store: Record<string, T>, storageKey: string): Record<string, T> | null {
  const keys = Object.keys(store);
  if (!keys.some((k) => LEGACY_ID.test(k))) return null;

  const backupKey = `${storageKey}:pre-id-migration`;
  try {
    if (localStorage.getItem(backupKey) === null) {
      localStorage.setItem(backupKey, JSON.stringify(store));
    }
  } catch (err) {
    console.warn("Não foi possível guardar a cópia de segurança antes da migração de ids.", err);
  }

  const migrated: Record<string, T> = {};
  for (const key of keys) {
    if (!LEGACY_ID.test(key)) migrated[key] = store[key];
  }
  for (const key of keys) {
    if (!LEGACY_ID.test(key)) continue;
    const target = LEGACY_ID_MAP[key];
    if (target === null) continue;
    const newKey = target ?? key;
    if (!(newKey in migrated)) migrated[newKey] = store[key];
  }
  return migrated;
}
