import { Sparkles } from "lucide-react";
import type { MergedProgram } from "../../types/program";
import { generateResearchPrompt } from "../../utils/generateResearchPrompt";
import { useToast } from "../../context/ToastContext";

interface ResearchPromptButtonProps {
  program: MergedProgram;
  /** compact = ícone só (usado na tabela); default = ícone + texto (usado nos cards) */
  compact?: boolean;
}

export function ResearchPromptButton({ program, compact = false }: ResearchPromptButtonProps) {
  const { showToast } = useToast();

  const handleClick = async (e: React.MouseEvent) => {
    e.stopPropagation();
    const prompt = generateResearchPrompt(program);
    try {
      await navigator.clipboard.writeText(prompt);
      showToast("Prompt copiado! Cole no terminal do VS Code para iniciar a pesquisa.");
    } catch {
      showToast("Não foi possível copiar o prompt automaticamente. Tente novamente.");
    }
  };

  if (compact) {
    return (
      <button
        onClick={handleClick}
        title="Gerar Prompt de Pesquisa"
        aria-label="Gerar Prompt de Pesquisa"
        className="flex items-center justify-center rounded-lg p-1.5 text-violet-500 hover:bg-violet-50 hover:text-violet-600"
      >
        <Sparkles size={15} />
      </button>
    );
  }

  return (
    <button
      onClick={handleClick}
      title="Gerar Prompt de Pesquisa"
      className="flex items-center gap-1.5 rounded-lg border border-violet-200 bg-violet-50 px-2.5 py-1.5 text-xs font-medium text-violet-700 hover:bg-violet-100"
    >
      <Sparkles size={13} /> Gerar Prompt de Pesquisa
    </button>
  );
}
