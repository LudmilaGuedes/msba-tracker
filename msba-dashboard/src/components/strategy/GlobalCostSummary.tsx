import { Calculator } from "lucide-react";
import { formatUSD } from "../../utils/format";
import type { GlobalCosts } from "../../utils/strategyCosts";
import { GMAT_FREE_SENDS, type FinalTotalFilter } from "../../types/strategy";

interface GlobalCostSummaryProps {
  costs: GlobalCosts;
  filter: FinalTotalFilter;
  onFilterChange: (filter: FinalTotalFilter) => void;
}

export function GlobalCostSummary({ costs, filter, onFilterChange }: GlobalCostSummaryProps) {
  return (
    <div className="rounded-xl border border-indigo-200 bg-indigo-50/60 p-5">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2 text-sm font-semibold text-indigo-900">
          <Calculator size={16} /> Resumo Geral (cumulativo)
        </div>
        <label className="flex items-center gap-2 text-xs text-indigo-700">
          Considerar no total:
          <select
            value={filter}
            onChange={(e) => onFilterChange(e.target.value as FinalTotalFilter)}
            className="rounded-md border border-indigo-200 bg-white px-2 py-1 text-xs text-indigo-900 outline-none focus:border-indigo-400"
          >
            <option value="all">Todas as faculdades</option>
            <option value="certeza">Só "Aplicar com certeza"</option>
            <option value="duvida">Só "Dúvida se vou aplicar"</option>
          </select>
        </label>
      </div>
      <dl className="grid grid-cols-2 gap-x-6 gap-y-2 text-sm sm:grid-cols-4">
        <div>
          <dt className="text-xs text-indigo-500">Taxas de Aplicação</dt>
          <dd className="font-semibold text-slate-900">{formatUSD(costs.feesTotal)}</dd>
        </div>
        <div>
          <dt className="text-xs text-indigo-500">GMAT bruto ({costs.gmatCount}× envios)</dt>
          <dd className="font-semibold text-slate-900">{formatUSD(costs.gmatGrossCost)}</dd>
        </div>
        <div>
          <dt className="text-xs text-indigo-500">
            Desconto GMAT ({costs.gmatFreeSendsApplied}/{GMAT_FREE_SENDS} grátis)
          </dt>
          <dd className="font-semibold text-emerald-600">− {formatUSD(costs.gmatDiscount)}</dd>
        </div>
        <div>
          <dt className="text-xs text-indigo-500">GMAT líquido</dt>
          <dd className="font-semibold text-slate-900">{formatUSD(costs.gmatNetCost)}</dd>
        </div>
      </dl>
      <div className="mt-4 flex items-center justify-end border-t border-indigo-200 pt-3">
        <span className="text-lg font-bold text-indigo-900">Total Geral: {formatUSD(costs.grandTotal)}</span>
      </div>
    </div>
  );
}
