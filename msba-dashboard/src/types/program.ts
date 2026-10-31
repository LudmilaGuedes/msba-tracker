export const VERIFICAR = "Verificar";

/** Valor bruto de uma célula: string, número, ou o marcador "Verificar" para dados faltantes. */
export type CellValue = string | number;

/** Um programa MSBA como veio da planilha original (55 colunas, chaves = ColumnDef.key). */
export type RawProgram = Record<string, CellValue> & {
  id: string;
};

export interface Folder {
  id: string;
  name: string;
  color: string;
  createdAt: number;
}

export interface ProgramUserData {
  /** Overrides de campos originais, indexado por chave da coluna. */
  edits?: Record<string, CellValue>;
  rating?: number; // 0-5
  comments?: string;
  folderIds?: string[];
  updatedAt?: number;
}

export type UserDataStore = Record<string, ProgramUserData>;

/** Programa após mesclar dados originais + edições do usuário + metadados (rating/comments/folders). */
export interface MergedProgram {
  id: string;
  raw: RawProgram;
  values: Record<string, CellValue>;
  editedKeys: Set<string>;
  rating: number;
  comments: string;
  folderIds: string[];
  /** true quando o programa foi criado manualmente no app, não vindo da planilha original. */
  isCustom: boolean;
}
