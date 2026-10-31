import type { CellValue, MergedProgram } from "../types/program";
import { parseNumeric } from "./format";

/** Custo total real do programa: tuition internacional + custo de vida total, quando ambos conhecidos. */
export function getRealTotalCost(program: MergedProgram): number | null {
  const tuition = parseNumeric(program.values.tuitionTotalIntl);
  const living = parseNumeric(program.values.custoVidaTotal);
  if (tuition === null && living === null) return null;
  return (tuition ?? 0) + (living ?? 0);
}

/** Componentes mensais de custo de vida que alimentam 'Custo Vida Total (programa)'. */
export const CUSTO_VIDA_MONTHLY_KEYS = [
  "aluguelMes",
  "mercadoMes",
  "seguroResidencialMes",
  "utilitiesMes",
  "celularMes",
  "wifiMes",
  "transporteMes",
  "outrosGastosMes",
] as const;

/** Além dos componentes mensais, mudar estes dois também deve recalcular o total. */
export const CUSTO_VIDA_TRIGGER_KEYS = [...CUSTO_VIDA_MONTHLY_KEYS, "seguroSaude", "duracaoMeses"] as const;

/**
 * Recalcula 'Custo Vida Total (programa) USD' a partir dos componentes
 * mensais de custo de vida × duração do programa, mais o seguro saúde
 * (valor já anual/de programa, não mensal — não é multiplicado pela
 * duração). Componentes desconhecidos ("Verificar"/vazio) contam como 0,
 * não bloqueiam o cálculo. Retorna null se a duração do programa também
 * for desconhecida — sem ela não há como extrapolar um total.
 */
export function computeCustoVidaTotal(values: Record<string, CellValue>): number | null {
  const duracao = parseNumeric(values.duracaoMeses);
  if (duracao === null) return null;

  let somaMensal = 0;
  for (const key of CUSTO_VIDA_MONTHLY_KEYS) {
    somaMensal += parseNumeric(values[key]) ?? 0;
  }
  const seguroSaude = parseNumeric(values.seguroSaude) ?? 0;

  return Math.round((somaMensal * duracao + seguroSaude) * 100) / 100;
}
