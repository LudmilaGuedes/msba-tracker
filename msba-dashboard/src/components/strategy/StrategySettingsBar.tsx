import { Settings2 } from "lucide-react";
import { useStrategy } from "../../context/StrategyContext";
import { GMAT_FREE_SENDS } from "../../types/strategy";

export function StrategySettingsBar() {
  const { settings, updateSettings } = useStrategy();

  return (
    <div className="flex flex-wrap items-center gap-4 rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm shadow-sm">
      <span className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">
        <Settings2 size={14} /> Custo de Envio Padrão
      </span>
      <label className="flex items-center gap-2 text-slate-600">
        Custo Envio GMAT ($) após os 5 grátis
        <input
          type="number"
          min={0}
          step={1}
          value={settings.gmatSendCost}
          onChange={(e) => updateSettings({ gmatSendCost: Number(e.target.value) || 0 })}
          className="w-20 rounded-md border border-slate-200 px-2 py-1 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </label>
      <span className="text-xs text-slate-400">
        O GMAT/GRE oferece os {GMAT_FREE_SENDS} primeiros envios de nota grátis (aplicado no Resumo Geral); a partir do{" "}
        {GMAT_FREE_SENDS + 1}º, cada envio custa o valor acima. O Duolingo (DET) é gratuito para enviar a quantas
        faculdades você quiser.
      </span>
    </div>
  );
}
