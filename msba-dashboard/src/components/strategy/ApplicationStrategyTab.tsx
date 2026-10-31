import { useEffect, useMemo, useState } from "react";
import { useData } from "../../context/DataContext";
import { useStrategy } from "../../context/StrategyContext";
import { STRATEGY_ELIGIBLE_FOLDER_NAMES, type FinalTotalFilter } from "../../types/strategy";
import { LEAD_TIME_BUCKET_LABELS, ROUND_LABELS } from "../../utils/strategyParsing";
import { buildStrategyEntries, groupEntriesByBucket, hasFolderNamed, isStrategyEligible } from "../../utils/strategyGrouping";
import { computeGlobalCosts, computeRoundCosts } from "../../utils/strategyCosts";
import { StrategySettingsBar } from "./StrategySettingsBar";
import { RoundTable } from "./RoundTable";
import { RoundCostPanel } from "./RoundCostPanel";
import { GlobalCostSummary } from "./GlobalCostSummary";
import { FolderKanban } from "lucide-react";

const [CERTEZA_NAME, DUVIDA_NAME] = STRATEGY_ELIGIBLE_FOLDER_NAMES;

export function ApplicationStrategyTab() {
  const { programs, folders, ensureFolders } = useData();
  const { strategyData, settings } = useStrategy();
  const [finalFilter, setFinalFilter] = useState<FinalTotalFilter>("all");

  useEffect(() => {
    ensureFolders([...STRATEGY_ELIGIBLE_FOLDER_NAMES]);
  }, [ensureFolders]);

  const eligiblePrograms = useMemo(() => programs.filter((p) => isStrategyEligible(p, folders)), [programs, folders]);
  const allEntries = useMemo(() => buildStrategyEntries(eligiblePrograms), [eligiblePrograms]);
  const groups = useMemo(() => groupEntriesByBucket(allEntries), [allEntries]);

  const entriesForTotal = useMemo(() => {
    if (finalFilter === "all") return allEntries;
    const targetName = finalFilter === "certeza" ? CERTEZA_NAME : DUVIDA_NAME;
    return allEntries.filter((e) => hasFolderNamed(e.program, folders, targetName));
  }, [allEntries, finalFilter, folders]);

  const globalCosts = useMemo(
    () => computeGlobalCosts(entriesForTotal, strategyData, settings),
    [entriesForTotal, strategyData, settings],
  );

  return (
    <div className="space-y-5">
      <div className="rounded-xl border border-slate-200 bg-white p-4 text-sm text-slate-600 shadow-sm">
        <p>
          Mostrando faculdades marcadas com as pastas <strong>"Aplicar com certeza"</strong> ou{" "}
          <strong>"Dúvida se vou aplicar"</strong>. Use o gerenciador de pastas (
          <FolderKanban size={13} className="inline align-text-bottom" />) para classificar mais programas — eles aparecem
          aqui automaticamente. Cada faculdade é considerada pelo seu primeiro round/deadline (assumindo que você sempre
          aplica no primeiro round disponível) e agrupada pelo mês desse prazo: Round 1 = setembro a dezembro, Round 2 = janeiro a maio,
          Round 3 = junho a agosto — as mesmas janelas para todos os programas, qualquer que seja o mês de início do curso.
        </p>
      </div>

      <StrategySettingsBar />

      {eligiblePrograms.length === 0 ? (
        <div className="rounded-xl border border-dashed border-slate-300 bg-white p-10 text-center text-sm text-slate-400">
          Nenhuma faculdade classificada como "Aplicar com certeza" ou "Dúvida se vou aplicar" ainda.
        </div>
      ) : (
        <>
          {groups.map((group) => (
            <div key={group.bucket} className="space-y-3">
              <h2 className="text-base font-semibold text-slate-800">
                {ROUND_LABELS[group.bucket]} — {LEAD_TIME_BUCKET_LABELS[group.bucket]}
              </h2>
              <RoundTable entries={group.entries} folders={folders} />
              <RoundCostPanel costs={computeRoundCosts(group.entries, strategyData, settings)} />
            </div>
          ))}

          <GlobalCostSummary costs={globalCosts} filter={finalFilter} onFilterChange={setFinalFilter} />
        </>
      )}
    </div>
  );
}
