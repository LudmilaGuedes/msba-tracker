import { AlertTriangle, GraduationCap } from "lucide-react";
import type { Folder } from "../../types/program";
import { useStrategy } from "../../context/StrategyContext";
import { StarRating } from "../common/StarRating";
import { ToggleSwitch } from "../common/ToggleSwitch";
import { formatUSD, parseNumeric } from "../../utils/format";
import { parseDuolingo, parseGmatStatus, parseRecommendationLetters } from "../../utils/strategyParsing";
import { effectiveApplicationFee, willSendGmat } from "../../utils/strategyCosts";
import type { StrategyEntry } from "../../utils/strategyGrouping";

interface RoundTableProps {
  entries: StrategyEntry[];
  folders: Folder[];
}

export function RoundTable({ entries, folders }: RoundTableProps) {
  return (
    <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm">
      <table className="w-full min-w-[1400px] border-collapse text-sm">
        <thead>
          <tr className="border-b border-slate-200 bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
            <th className="px-3 py-2.5">Faculdade</th>
            <th className="px-3 py-2.5">Minhas Classificações</th>
            <th className="px-3 py-2.5">Data de Aplicação</th>
            <th className="px-3 py-2.5">Taxa de Aplicação ($)</th>
            <th className="px-3 py-2.5">OBS Taxa</th>
            <th className="px-3 py-2.5">Cartas Recomendação</th>
            <th className="px-3 py-2.5">Duolingo (DET)</th>
            <th className="px-3 py-2.5">GMAT</th>
            <th className="px-3 py-2.5">Dificuldade Entrada</th>
            <th className="px-3 py-2.5">Data Prevista Resultado</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {entries.map((entry) => (
            <StrategyRow key={entry.program.id} entry={entry} folders={folders} />
          ))}
        </tbody>
      </table>
    </div>
  );
}

function StrategyRow({ entry, folders }: { entry: StrategyEntry; folders: Folder[] }) {
  const { program, roundText } = entry;
  const { strategyData, updateStrategy } = useStrategy();
  const strategy = strategyData[program.id] ?? {};

  const fee = effectiveApplicationFee(entry, strategyData);
  const letters = parseRecommendationLetters(program.values.cartasRecomendacao, program.values.documentosNecessarios);
  const duolingo = parseDuolingo(program.values.testesIngles);
  const gmatStatus = parseGmatStatus(program.values.gmatGreObrigatorio);
  const gmatOn = willSendGmat(entry, strategyData);

  const programFolders = program.folderIds
    .map((id) => folders.find((f) => f.id === id))
    .filter((f): f is Folder => Boolean(f));

  return (
    <tr className="align-top hover:bg-slate-50/60">
      <td className="px-3 py-2.5">
        <div className="font-medium text-slate-800">{String(program.values.universidade ?? "")}</div>
        <div className="text-xs text-slate-400">{String(program.values.businessSchool ?? "")}</div>
        {program.rating > 0 && <StarRating value={program.rating} readOnly size={13} />}
      </td>

      <td className="px-3 py-2.5">
        <div className="flex flex-wrap gap-1">
          {programFolders.length === 0 && <span className="text-xs text-slate-300">—</span>}
          {programFolders.map((f) => (
            <span
              key={f.id}
              className="rounded-full px-2 py-0.5 text-[11px] font-medium text-white"
              style={{ backgroundColor: f.color }}
            >
              {f.name}
            </span>
          ))}
        </div>
      </td>

      <td className="px-3 py-2.5 text-slate-700">{roundText}</td>

      <td className="px-3 py-2.5">
        <input
          type="number"
          min={0}
          step={1}
          value={fee}
          onChange={(e) => updateStrategy(program.id, { applicationFee: Number(e.target.value) || 0 })}
          className="w-24 rounded-md border border-slate-200 px-2 py-1 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
        {strategy.applicationFee !== undefined && (
          <div className="mt-0.5 text-[11px] text-slate-400">
            original: {formatUSD(parseNumeric(program.values.taxaAplicacao))}
          </div>
        )}
      </td>

      <td className="px-3 py-2.5">
        <input
          type="text"
          placeholder="Ex.: Isento se assistir sessão"
          value={strategy.feeNotes ?? ""}
          onChange={(e) => updateStrategy(program.id, { feeNotes: e.target.value })}
          className="w-44 rounded-md border border-slate-200 px-2 py-1 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </td>

      <td className="px-3 py-2.5 text-slate-700">{letters === null ? "Verificar" : letters}</td>

      <td className="px-3 py-2.5">
        {duolingo.accepted === false ? (
          <div className="flex items-center gap-1.5 font-medium text-red-600">
            <AlertTriangle size={16} className="shrink-0" />
            Não aceito
          </div>
        ) : (
          <div className="text-slate-700">{duolingo.accepted === true ? `Mín. ${duolingo.minScore ?? "Verificar"}` : "Verificar"}</div>
        )}
      </td>

      <td className="px-3 py-2.5">
        <div className="space-y-1">
          <div className="flex items-center gap-1 text-slate-700">
            <GraduationCap size={14} className="text-slate-400" />
            {gmatStatus}
          </div>
          <ToggleSwitch checked={gmatOn} onChange={(v) => updateStrategy(program.id, { gmatWillSend: v })} label="Vou Enviar" />
        </div>
      </td>

      <td className="px-3 py-2.5 text-slate-700">{String(program.values.dificuldadeEntrada ?? "Verificar")}</td>

      <td className="px-3 py-2.5">
        <input
          type="text"
          placeholder="Ex.: 15/mar/2027"
          value={strategy.resultDate ?? ""}
          onChange={(e) => updateStrategy(program.id, { resultDate: e.target.value })}
          className="w-32 rounded-md border border-slate-200 px-2 py-1 text-sm outline-none focus:border-indigo-300 focus:ring-2 focus:ring-indigo-100"
        />
      </td>
    </tr>
  );
}
