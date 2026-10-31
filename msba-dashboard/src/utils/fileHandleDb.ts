// IndexedDB guarda o FileSystemFileHandle entre sessões (localStorage não
// consegue serializar um handle). Um único registro: o arquivo de backup
// automático escolhido pelo usuário.

const DB_NAME = "msba-tracker-fs";
const STORE = "handles";
const KEY = "backup-file";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function storeBackupHandle(handle: FileSystemFileHandle): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(handle, KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

export async function loadBackupHandle(): Promise<FileSystemFileHandle | null> {
  const db = await openDb();
  const handle = await new Promise<FileSystemFileHandle | null>((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const req = tx.objectStore(STORE).get(KEY);
    req.onsuccess = () => resolve((req.result as FileSystemFileHandle | undefined) ?? null);
    req.onerror = () => reject(req.error);
  });
  db.close();
  return handle;
}

export async function clearBackupHandle(): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}
