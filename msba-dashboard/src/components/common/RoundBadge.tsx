import { LEAD_TIME_BUCKET_LABELS, ROUND_LABELS, type LeadTimeBucket } from "../../utils/strategyParsing";

const ROUND_CLASSES: Record<LeadTimeBucket, string> = {
  "ano-anterior": "bg-indigo-50 text-indigo-700",
  "mesmo-ano": "bg-amber-50 text-amber-700",
  "perto-inicio": "bg-rose-50 text-rose-700",
  "sem-data": "bg-slate-100 text-slate-500",
};

interface RoundBadgeProps {
  bucket: LeadTimeBucket;
  className?: string;
}

/** Selo curto "Round 1/2/3" (com a explicação completa da faixa de antecedência via tooltip nativo). */
export function RoundBadge({ bucket, className = "" }: RoundBadgeProps) {
  return (
    <span
      title={LEAD_TIME_BUCKET_LABELS[bucket]}
      className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold ${ROUND_CLASSES[bucket]} ${className}`}
    >
      {ROUND_LABELS[bucket]}
    </span>
  );
}
