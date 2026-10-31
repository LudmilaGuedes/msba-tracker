interface ToggleSwitchProps {
  checked: boolean;
  onChange: (checked: boolean) => void;
  label: string;
  disabled?: boolean;
}

export function ToggleSwitch({ checked, onChange, label, disabled = false }: ToggleSwitchProps) {
  return (
    <button
      type="button"
      role="switch"
      aria-checked={checked}
      disabled={disabled}
      onClick={(e) => {
        e.stopPropagation();
        onChange(!checked);
      }}
      className={`flex items-center gap-1.5 rounded-full border px-2 py-1 text-[11px] font-medium transition-colors ${
        disabled
          ? "cursor-not-allowed border-slate-100 bg-slate-50 text-slate-300"
          : checked
            ? "border-emerald-500 bg-emerald-50 text-emerald-700"
            : "border-slate-200 bg-white text-slate-500 hover:bg-slate-50"
      }`}
    >
      <span
        className={`h-3 w-3 rounded-full border transition-colors ${
          disabled
            ? "border-slate-200 bg-slate-100"
            : checked
              ? "border-emerald-600 bg-emerald-500"
              : "border-slate-300 bg-white"
        }`}
      />
      {label}
    </button>
  );
}
