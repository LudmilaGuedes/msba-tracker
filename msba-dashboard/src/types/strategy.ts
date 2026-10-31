// Estado da aba "Estratégia de Aplicação": dados por-programa que não fazem
// parte do schema de 55 colunas (COLUMNS), somados às configurações globais
// de custo de envio de notas. Persistidos em localStorage, separados do
// overlay de userData (rating/comments/folders/edits) usado pelo resto do app.

export interface ProgramStrategyData {
  /** Override manual da Taxa de Aplicação (para zerar/reduzir por isenção). Undefined = usa o valor original da planilha. */
  applicationFee?: number;
  /** Anotações livres sobre regras de isenção da taxa (pré-preenchido com pesquisa quando disponível, editável). */
  feeNotes?: string;
  /** Usuário optou por enviar nota do GMAT/GRE, independente do status oficial (Obrigatório/Opcional) do programa. */
  gmatWillSend?: boolean;
  /** Data prevista do resultado/decisão daquele round (pré-preenchido com pesquisa quando disponível, editável). */
  resultDate?: string;
  updatedAt?: number;
}

export type StrategyDataStore = Record<string, ProgramStrategyData>;

export interface StrategySettings {
  /** Custo (USD) por envio de nota do GMAT/GRE a uma faculdade, além da cota gratuita (GMAT_FREE_SENDS). */
  gmatSendCost: number;
}

export const DEFAULT_STRATEGY_SETTINGS: StrategySettings = {
  gmatSendCost: 35,
};

/** Regra fixa do GMAT: os 5 primeiros envios de nota são gratuitos (cota do próprio exame); a partir do 6º, cada envio custa gmatSendCost. */
export const GMAT_FREE_SENDS = 5;

/** Nomes (case/acento-insensitive) das pastas usadas como filtro de elegibilidade para a aba de estratégia. */
export const STRATEGY_ELIGIBLE_FOLDER_NAMES = ["Aplicar com certeza", "Dúvida se vou aplicar"] as const;

/** Filtro de quais faculdades entram no cálculo do Resumo Geral (total final). */
export type FinalTotalFilter = "all" | "certeza" | "duvida";
