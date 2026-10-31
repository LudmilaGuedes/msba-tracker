import { useState } from "react";
import { Star } from "lucide-react";

interface StarRatingProps {
  value: number;
  onChange?: (value: number) => void;
  size?: number;
  readOnly?: boolean;
}

export function StarRating({ value, onChange, size = 18, readOnly = false }: StarRatingProps) {
  const [hover, setHover] = useState<number | null>(null);
  const display = hover ?? value;

  return (
    <div
      className={`flex items-center gap-0.5 ${readOnly ? "" : "cursor-pointer"}`}
      onMouseLeave={() => setHover(null)}
      role={readOnly ? undefined : "radiogroup"}
      aria-label="Classificação por estrelas"
    >
      {[1, 2, 3, 4, 5].map((star) => (
        <button
          key={star}
          type="button"
          disabled={readOnly}
          className={`p-0 leading-none ${readOnly ? "cursor-default" : ""}`}
          onMouseEnter={() => !readOnly && setHover(star)}
          onClick={(e) => {
            e.stopPropagation();
            if (readOnly || !onChange) return;
            onChange(star === value ? 0 : star);
          }}
          aria-label={`${star} estrela${star > 1 ? "s" : ""}`}
        >
          <Star
            size={size}
            className={
              star <= display
                ? "fill-amber-400 text-amber-400 transition-colors"
                : "fill-transparent text-slate-300 transition-colors"
            }
          />
        </button>
      ))}
    </div>
  );
}
