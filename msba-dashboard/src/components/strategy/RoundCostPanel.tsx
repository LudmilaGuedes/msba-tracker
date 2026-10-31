import { formatUSD } from "../../utils/format";
import type { RoundCosts } from "../../utils/strategyCosts";

export function RoundCostPanel({ costs }: { costs: RoundCosts }) {
  return (
    <div className="flex flex-wrap items-center gap-x-6 gap-y-2 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm">
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">Custos deste Grupo</span>
      <span className="text-slate-600">
        Taxas: <strong className="text-slate-900">{formatUSD(costs.feesTotal)}</strong>
      </span>
      <span className="text-slate-600">
        GMAT ({costs.gmatCount}×): <strong className="text-slate-900">{formatUSD(costs.gmatCost)}</strong>
      </span>
      <span className="ml-auto text-slate-600">
        Subtotal (sem desconto GMAT acumulado): <strong className="text-indigo-700">{formatUSD(costs.roundTotalGross)}</strong>
      </span>
    </div>
  );
}
