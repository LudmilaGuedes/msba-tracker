import { parseNumeric } from "./format";
import { GMAT_FREE_SENDS, type ProgramStrategyData, type StrategySettings } from "../types/strategy";
import type { StrategyEntry } from "./strategyGrouping";

export function effectiveApplicationFee(entry: StrategyEntry, strategyData: Record<string, ProgramStrategyData>): number {
  const override = strategyData[entry.program.id]?.applicationFee;
  if (override !== undefined) return override;
  return parseNumeric(entry.program.values.taxaAplicacao) ?? 0;
}

export function willSendGmat(entry: StrategyEntry, strategyData: Record<string, ProgramStrategyData>): boolean {
  return strategyData[entry.program.id]?.gmatWillSend ?? false;
}

export interface RoundCosts {
  feesTotal: number;
  gmatCount: number;
  gmatCost: number;
  roundTotalGross: number;
}

export function computeRoundCosts(
  entries: StrategyEntry[],
  strategyData: Record<string, ProgramStrategyData>,
  settings: StrategySettings,
): RoundCosts {
  const feesTotal = entries.reduce((sum, e) => sum + effectiveApplicationFee(e, strategyData), 0);
  const gmatCount = entries.filter((e) => willSendGmat(e, strategyData)).length;
  const gmatCost = gmatCount * settings.gmatSendCost;
  return { feesTotal, gmatCount, gmatCost, roundTotalGross: feesTotal + gmatCost };
}

export interface GlobalCosts {
  feesTotal: number;
  gmatCount: number;
  gmatGrossCost: number;
  gmatFreeSendsApplied: number;
  gmatDiscount: number;
  gmatNetCost: number;
  grandTotal: number;
}

/**
 * Calcula o total geral (todas as rodadas juntas, já respeitando o filtro de
 * elegibilidade escolhido pelo usuário para o Resumo Geral). Regra crítica:
 * os 5 primeiros envios de nota do GMAT são gratuitos (cota do próprio
 * exame) — o desconto é aplicado uma única vez aqui, sobre o total
 * cumulativo de envios, nunca round a round. O Duolingo (DET) não tem
 * custo: é gratuito enviar para quantas faculdades o usuário quiser.
 */
export function computeGlobalCosts(
  allEntries: StrategyEntry[],
  strategyData: Record<string, ProgramStrategyData>,
  settings: StrategySettings,
): GlobalCosts {
  const feesTotal = allEntries.reduce((sum, e) => sum + effectiveApplicationFee(e, strategyData), 0);
  const gmatCount = allEntries.filter((e) => willSendGmat(e, strategyData)).length;

  const gmatGrossCost = gmatCount * settings.gmatSendCost;
  const gmatFreeSendsApplied = Math.min(GMAT_FREE_SENDS, gmatCount);
  const gmatDiscount = gmatFreeSendsApplied * settings.gmatSendCost;
  const gmatNetCost = gmatGrossCost - gmatDiscount;

  return {
    feesTotal,
    gmatCount,
    gmatGrossCost,
    gmatFreeSendsApplied,
    gmatDiscount,
    gmatNetCost,
    grandTotal: feesTotal + gmatNetCost,
  };
}
