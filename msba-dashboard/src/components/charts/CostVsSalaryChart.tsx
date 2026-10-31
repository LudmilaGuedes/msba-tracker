import {
  CartesianGrid,
  ResponsiveContainer,
  Scatter,
  ScatterChart,
  Tooltip,
  XAxis,
  YAxis,
  ZAxis,
} from "recharts";
import type { MergedProgram } from "../../types/program";
import { parseNumeric, formatUSD } from "../../utils/format";
import { getRealTotalCost } from "../../utils/derived";

interface CostVsSalaryChartProps {
  programs: MergedProgram[];
}

interface Point {
  name: string;
  cost: number;
  salary: number;
  rating: number;
}

export function CostVsSalaryChart({ programs }: CostVsSalaryChartProps) {
  const points: Point[] = programs
    .map((p) => {
      const cost = getRealTotalCost(p);
      const salary = parseNumeric(p.values.salarioMedioPos);
      if (cost === null || salary === null) return null;
      return { name: String(p.values.universidade), cost, salary, rating: p.rating || 1 };
    })
    .filter((p): p is Point => p !== null);

  if (points.length < 2) {
    return (
      <div className="flex h-64 items-center justify-center rounded-2xl border border-slate-200 bg-white text-sm text-slate-400">
        Dados insuficientes para o gráfico (custo e salário precisam estar preenchidos).
      </div>
    );
  }

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <h3 className="mb-1 text-sm font-semibold text-slate-800">Custo Real vs. Salário Médio Pós-Programa</h3>
      <p className="mb-3 text-xs text-slate-400">
        Quanto mais para baixo e à direita, melhor o custo-benefício. Tamanho do ponto = sua nota.
      </p>
      <ResponsiveContainer width="100%" height={320}>
        <ScatterChart margin={{ top: 10, right: 20, bottom: 10, left: 0 }}>
          <CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" />
          <XAxis
            type="number"
            dataKey="cost"
            name="Custo Total Real"
            tickFormatter={(v) => `$${Math.round(v / 1000)}k`}
            tick={{ fontSize: 11, fill: "#64748b" }}
          />
          <YAxis
            type="number"
            dataKey="salary"
            name="Salário Pós"
            tickFormatter={(v) => `$${Math.round(v / 1000)}k`}
            tick={{ fontSize: 11, fill: "#64748b" }}
          />
          <ZAxis type="number" dataKey="rating" range={[40, 220]} />
          <Tooltip
            cursor={{ strokeDasharray: "3 3" }}
            content={({ active, payload }) => {
              if (!active || !payload?.length) return null;
              const p = payload[0].payload as Point;
              return (
                <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs shadow-md">
                  <div className="mb-1 font-semibold text-slate-800">{p.name}</div>
                  <div className="text-slate-500">Custo total: {formatUSD(p.cost)}</div>
                  <div className="text-slate-500">Salário pós: {formatUSD(p.salary)}</div>
                </div>
              );
            }}
          />
          <Scatter data={points} fill="#6366f1" fillOpacity={0.75} />
        </ScatterChart>
      </ResponsiveContainer>
    </div>
  );
}
