import type { MergedProgram } from "../types/program";

/**
 * Gera o prompt de Deep Research / Data Enrichment para uma universidade específica,
 * pronto para ser colado no terminal do Claude Code (ou outro agente de engenharia de dados).
 */
export function generateResearchPrompt(program: MergedProgram): string {
  const nomeDaUniversidade = String(program.values.universidade ?? "");
  const cidade = String(program.values.cidade ?? "");
  const estado = String(program.values.estado ?? "");

  return `Atuando como um engenheiro de dados, faça um processo de Data Enrichment focado em uma única instituição.
Alvo: ${nomeDaUniversidade}, localizada em ${cidade}, ${estado}. Foco exclusivo no programa de Master of Science in Business Analytics (presencial, STEM).

Auditoria: Leia o msba_programs_database.xlsx. Localize a linha desta faculdade.

Deep Research: Busque por Tuition (valor por crédito e total, residente e internacional); Bolsas disponíveis — quais tipos existem (mérito, assistantship, fellowship, etc.) e quão realista é cada uma para um candidato internacional; Taxa de aplicação (valor e se há alguma forma de isenção); Requisitos completos de aplicação — GPA mínimo, GMAT/GRE, testes de inglês (incluindo se aceita Duolingo e a nota mínima), documentos exigidos e prazos; Taxa de empregabilidade e salário médio dos formados (e principais empregadores, se encontrado); Duração real do curso (em meses) e formato (presencial/online/híbrido).

Transformação: Respeite o schema de 55 colunas. Substitua dados incorretos e preencha lacunas. Mantenha 'Verificar' para o que não for encontrado.

Carga: Execute um script para atualizar APENAS a linha dessa universidade no msba_programs_database.xlsx sobrescrevendo o arquivo.

Resumo: Mostre no terminal o que foi alterado.`;
}
