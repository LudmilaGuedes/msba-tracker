// Heurísticas de parsing para extrair, a partir dos campos de texto livre da
// planilha original (preenchidos por pesquisa manual, formato inconsistente
// entre programas), a estrutura usada pela aba "Estratégia de Aplicação":
// rounds de deadline, status de GMAT, aceite/nota mínima de Duolingo e
// quantidade de cartas de recomendação. Nenhum desses campos existe como
// coluna estruturada em COLUMNS — são melhores esforços sobre texto livre,
// não substituem confirmação manual (por isso mantêm "Verificar" quando o
// texto não permite concluir algo com segurança).
import { isVerificar } from "./format";
import type { CellValue } from "../types/program";

export interface ParsedRound {
  label: string;
  dateText: string;
}

const ROUND_REGEX = /Round\s*(\d+)\s*[:\-]?\s*([^;]+)/gi;

/** Divide "Deadline Aplicação" em rounds. Sem padrão "Round N" reconhecível, trata como rodada única. */
export function parseRounds(deadlineAplicacao: CellValue | undefined): ParsedRound[] {
  if (deadlineAplicacao === undefined || isVerificar(deadlineAplicacao)) {
    return [{ label: "Rodada Única", dateText: "Verificar" }];
  }
  const text = String(deadlineAplicacao).trim();
  if (!text) return [{ label: "Rodada Única", dateText: "Verificar" }];

  const matches = [...text.matchAll(ROUND_REGEX)];
  if (matches.length === 0) {
    return [{ label: "Rodada Única", dateText: text }];
  }
  return matches.map((m) => ({ label: `Round ${m[1]}`, dateText: m[2].trim().replace(/[.,]$/, "") }));
}

// Mapa único PT+EN de meses: os textos de pesquisa misturam os dois
// idiomas de programa pra programa (ex.: "Round 1: 9 jan" vs "R1 Oct 10").
// Nomes completos são listados explicitamente e o match exige a palavra
// inteira: aceitar qualquer sufixo fazia "abrem" virar abril, "janela"
// virar janeiro e "decisão" virar dezembro.
const MONTHS: Record<string, number> = {
  january: 0, janeiro: 0, jan: 0,
  february: 1, fevereiro: 1, feb: 1, fev: 1,
  march: 2, março: 2, marco: 2, mar: 2,
  april: 3, abril: 3, apr: 3, abr: 3,
  may: 4, maio: 4, mai: 4,
  june: 5, junho: 5, jun: 5,
  july: 6, julho: 6, jul: 6,
  august: 7, agosto: 7, aug: 7, ago: 7,
  september: 8, setembro: 8, sept: 8, sep: 8, set: 8,
  october: 9, outubro: 9, oct: 9, out: 9,
  november: 10, novembro: 10, nov: 10,
  december: 11, dezembro: 11, dec: 11, dez: 11,
};
const MONTH_KEYS_BY_LENGTH = Object.keys(MONTHS).sort((a, b) => b.length - a.length);
// Fronteiras explícitas (em vez de \b) porque \b do JS não trata "ç" como letra.
const MONTH_PATTERN = `(?<![a-zà-ÿ])(${MONTH_KEYS_BY_LENGTH.join("|")})(?![a-zà-ÿ])`;

/** Índice (0-11) do primeiro nome/abreviação de mês (PT ou EN) encontrado no texto, ou null. */
function findFirstMonth(text: string): number | null {
  const re = new RegExp(MONTH_PATTERN, "i");
  const m = text.toLowerCase().match(re);
  return m ? MONTHS[m[1].toLowerCase()] : null;
}

/**
 * Tenta extrair uma data aproximada de um texto livre em PT/EN
 * (ex.: "9 jan", "Round 2: 6 mar", "~out/2026", "Oct 10") para permitir
 * ordenação cronológica. Retorna null quando não há padrão dia/mês
 * reconhecível — nesses casos o chamador deve ordenar o item por último.
 */
export function estimateDate(dateText: string, fallbackYear = new Date().getFullYear()): Date | null {
  const s = dateText.toLowerCase();
  const re = new RegExp(`(\\d{1,2})?\\s*[/ ]?\\s*${MONTH_PATTERN}\\s*[/ ]?\\s*(\\d{4})?`, "i");
  const m = s.match(re);
  if (!m) return null;
  const day = m[1] ? Number(m[1]) : 1;
  const month = MONTHS[m[2].toLowerCase()];
  const year = m[3] ? Number(m[3]) : fallbackYear;
  const d = new Date(year, month, day);
  return Number.isNaN(d.getTime()) ? null : d;
}

/**
 * Mês (0-11) do primeiro deadline mencionado em "Deadline Aplicação". Regra
 * da base (ver CLAUDE.md): o texto sempre começa pelo prazo MAIS CEDO da
 * entrada principal (Round 1, Early Action, prioridade de bolsa...), porque
 * aplicar na primeira data possível dá mais chance de admissão e bolsa.
 * Retorna null quando o texto não traz nenhum mês reconhecível (ex.:
 * "Verificar", "Rolling" sem data).
 */
export function firstDeadlineMonth(deadlineAplicacao: CellValue | undefined): number | null {
  if (deadlineAplicacao === undefined || isVerificar(deadlineAplicacao)) return null;
  return findFirstMonth(String(deadlineAplicacao));
}

export type LeadTimeBucket = "ano-anterior" | "mesmo-ano" | "perto-inicio" | "sem-data";

export const LEAD_TIME_BUCKET_ORDER: LeadTimeBucket[] = ["ano-anterior", "mesmo-ano", "perto-inicio", "sem-data"];

export const LEAD_TIME_BUCKET_LABELS: Record<LeadTimeBucket, string> = {
  "ano-anterior": "Primeiro prazo entre setembro e dezembro (ano anterior ao Fall)",
  "mesmo-ano": "Primeiro prazo entre janeiro e maio",
  "perto-inicio": "Primeiro prazo entre junho e agosto",
  "sem-data": "Data de aplicação não confirmada",
};

/** Rótulo curto "Round 1/2/3" por janela do calendário, para exibição rápida no modal de detalhe. */
export const ROUND_LABELS: Record<LeadTimeBucket, string> = {
  "ano-anterior": "Round 1",
  "mesmo-ano": "Round 2",
  "perto-inicio": "Round 3",
  "sem-data": "Verificar",
};

/**
 * Classifica o primeiro deadline de um programa em janelas FIXAS do
 * calendário, iguais para todos os programas, qualquer que seja o mês de
 * início do curso. Os rounds servem para planejar QUANDO aplicar, não para
 * medir a antecedência em relação ao início:
 *   Round 1 = setembro a dezembro
 *   Round 2 = janeiro a maio
 *   Round 3 = junho a agosto
 */
export function classifyLeadTimeBucket(deadlineAplicacao: CellValue | undefined): LeadTimeBucket {
  const deadlineMonth = firstDeadlineMonth(deadlineAplicacao);
  if (deadlineMonth === null) return "sem-data";
  if (deadlineMonth >= 8) return "ano-anterior";
  if (deadlineMonth <= 4) return "mesmo-ano";
  return "perto-inicio";
}

export type GmatStatus = "Obrigatório" | "Opcional" | "Não Aceito" | "Verificar";

/** Normaliza o texto livre de "GMAT/GRE Obrigatório" em um status curto para exibição. */
export function parseGmatStatus(gmatGreObrigatorio: CellValue | undefined): GmatStatus {
  if (gmatGreObrigatorio === undefined || isVerificar(gmatGreObrigatorio)) return "Verificar";
  const s = String(gmatGreObrigatorio).trim().toLowerCase();
  if (!s) return "Verificar";
  if (s.includes("não aceito") || s.includes("nao aceito") || s.includes("removido")) return "Não Aceito";
  if (s.startsWith("obrigat") || s.includes("obrigat")) return "Obrigatório";
  if (s.startsWith("opcional") || s.includes("opcional")) return "Opcional";
  return "Verificar";
}

/**
 * Quantidade de cartas de recomendação exigidas. Usa o número no início de
 * "Cartas de Recomendação" (ex.: "0 (opcionais)" -> 0, "2-3" -> 2); se a
 * coluna não tiver número, tenta extrair de "Documentos Necessários".
 */
export function parseRecommendationLetters(
  cartasRecomendacao: CellValue | undefined,
  documentosNecessarios: CellValue | undefined,
): number | null {
  if (typeof cartasRecomendacao === "number") return cartasRecomendacao;
  if (cartasRecomendacao !== undefined && !isVerificar(cartasRecomendacao)) {
    const lead = String(cartasRecomendacao).match(/^\s*(\d+)/);
    if (lead) return Number(lead[1]);
  }
  if (documentosNecessarios === undefined || isVerificar(documentosNecessarios)) return null;
  const s = String(documentosNecessarios).toLowerCase();
  const m = s.match(/(\d+)\s*cartas?\s*(de\s*)?recomenda/);
  if (m) return Number(m[1]);
  if (/sem\s+(gmat\/gre\s+nem\s+)?cartas?\s*(de\s*)?recomenda/.test(s)) return 0;
  return null;
}

export interface ParsedDuolingo {
  /** false = programa não aceita Duolingo (dispara alerta visual). null = dado não permite concluir (sem menção nem "Verificar" explícito). */
  accepted: boolean | null;
  minScore: number | null;
}

/** Extrai aceite/nota mínima de Duolingo (DET) a partir de "Testes Inglês Aceitos". */
export function parseDuolingo(testesIngles: CellValue | undefined): ParsedDuolingo {
  if (testesIngles === undefined || isVerificar(testesIngles)) return { accepted: null, minScore: null };
  const s = String(testesIngles).toLowerCase();
  const idx = s.indexOf("duolingo");
  if (idx === -1) return { accepted: false, minScore: null };

  const negAfter = /duolingo[^.;]{0,30}n[aã]o[^.;]{0,20}aceit/.test(s);
  const negBefore = /n[aã]o[^.;]{0,20}aceit[^.;]{0,30}duolingo/.test(s);
  if (negAfter || negBefore) return { accepted: false, minScore: null };

  const windowText = s.slice(idx, idx + 80);
  const scoreMatch = windowText.match(/(\d{2,3})/);
  return { accepted: true, minScore: scoreMatch ? Number(scoreMatch[1]) : null };
}
