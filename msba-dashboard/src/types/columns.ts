// Definição canônica das 55 colunas do banco de dados MSBA.
// Cada entrada mapeia o cabeçalho original (PT, planilha) para uma chave
// camelCase usada no app, junto do tipo/grupo que orientam parsing,
// renderização da tabela e os filtros.

import { US_STATES } from "./usStates";

export type ColumnType = "text" | "longtext" | "number" | "percent" | "currency" | "enum";

export interface ColumnDef {
  key: string;
  header: string;
  group: string;
  type: ColumnType;
  /** Para colunas enum: valores conhecidos (usado em filtros de select). */
  options?: string[];
  /** Largura sugerida para a tabela. */
  width?: number;
}

export const COLUMNS: ColumnDef[] = [
  { key: "universidade", header: "Nome da Universidade", group: "Universidade", type: "text", width: 220 },
  { key: "estado", header: "Estado", group: "Universidade", type: "enum", options: US_STATES, width: 90 },
  { key: "cidade", header: "Cidade", group: "Universidade", type: "text", width: 140 },
  { key: "businessSchool", header: "Business School", group: "Universidade", type: "text", width: 200 },
  { key: "tipo", header: "Tipo (Pub/Priv)", group: "Universidade", type: "enum", options: ["Pública", "Privada"], width: 110 },

  { key: "nomePrograma", header: "Nome Oficial do Programa", group: "Programa", type: "text", width: 260 },
  { key: "stemDesignated", header: "STEM Designated", group: "Programa", type: "enum", options: ["Sim", "Não"], width: 110 },
  { key: "formato", header: "Formato", group: "Programa", type: "text", width: 110 },
  { key: "duracaoMeses", header: "Duração (meses)", group: "Programa", type: "number", width: 110 },
  { key: "inicio", header: "Início", group: "Programa", type: "text", width: 100 },
  { key: "tamanhoTurma", header: "Tamanho Turma", group: "Programa", type: "number", width: 110 },
  { key: "destaquesCurriculo", header: "Destaques Currículo", group: "Programa", type: "longtext", width: 260 },

  { key: "totalCreditos", header: "Total Créditos", group: "Custos", type: "number", width: 110 },
  { key: "custoCreditoIntl", header: "Custo/Crédito (Intl) USD", group: "Custos", type: "currency", width: 140 },
  { key: "custoCreditoResidente", header: "Custo/Crédito (Residente) USD", group: "Custos", type: "currency", width: 150 },
  { key: "tuitionTotalIntl", header: "Tuition Total (Intl) USD", group: "Custos", type: "currency", width: 150 },
  { key: "tuitionTotalResidente", header: "Tuition Total (Residente) USD", group: "Custos", type: "currency", width: 160 },
  { key: "feesAdicionais", header: "Fees Adicionais USD", group: "Custos", type: "currency", width: 130 },
  { key: "aluguelMes", header: "Aluguel Privativo/mês USD", group: "Custos", type: "currency", width: 140 },
  { key: "mercadoMes", header: "Mercado/mês USD", group: "Custos", type: "currency", width: 120 },
  { key: "seguroSaude", header: "Seguro Saúde (programa) USD", group: "Custos", type: "currency", width: 150 },
  { key: "seguroResidencialMes", header: "Seguro Residencial/mês USD", group: "Custos", type: "currency", width: 150 },
  { key: "utilitiesMes", header: "Utilities/mês USD", group: "Custos", type: "currency", width: 120 },
  { key: "celularMes", header: "Celular/mês USD", group: "Custos", type: "currency", width: 120 },
  { key: "wifiMes", header: "WiFi/mês USD", group: "Custos", type: "currency", width: 110 },
  { key: "transporteMes", header: "Transporte/mês USD", group: "Custos", type: "currency", width: 130 },
  { key: "outrosGastosMes", header: "Outros Gastos/mês USD", group: "Custos", type: "currency", width: 140 },
  { key: "custoVidaTotal", header: "Custo Vida Total (programa) USD", group: "Custos", type: "currency", width: 160 },

  { key: "bolsasDisponiveis", header: "Bolsas Disponíveis", group: "Bolsas", type: "enum", options: ["Sim", "Não"], width: 120 },
  { key: "tiposBolsa", header: "Tipos de Bolsa", group: "Bolsas", type: "text", width: 180 },
  { key: "quaoRealistaIntl", header: "Quão Realista (Intl)", group: "Bolsas", type: "text", width: 140 },
  { key: "taRaDisponivel", header: "TA/RA Disponível", group: "Bolsas", type: "text", width: 150 },
  { key: "processoTaRa", header: "Processo TA/RA", group: "Bolsas", type: "text", width: 180 },
  { key: "caminhoTuitionResidente", header: "Caminho Tuition Residente", group: "Bolsas", type: "text", width: 160 },
  { key: "graduateAssistantships", header: "Graduate Assistantships", group: "Bolsas", type: "text", width: 160 },

  { key: "rankingUsNews", header: "Ranking US News", group: "Rankings", type: "number", width: 110 },
  { key: "rankingQs", header: "Ranking QS", group: "Rankings", type: "number", width: 100 },
  { key: "taxaEmpregabilidade", header: "Taxa Empregabilidade %", group: "Rankings", type: "percent", width: 140 },
  { key: "salarioMedioPos", header: "Salário Médio Pós USD", group: "Rankings", type: "currency", width: 140 },
  { key: "principaisEmpregadores", header: "Principais Empregadores", group: "Rankings", type: "longtext", width: 220 },

  { key: "gpaMinimo", header: "GPA Mínimo", group: "Admissão", type: "number", width: 100 },
  { key: "gpaMedioAdmitidos", header: "GPA Médio Admitidos", group: "Admissão", type: "number", width: 130 },
  { key: "gmatGreObrigatorio", header: "GMAT/GRE Obrigatório", group: "Admissão", type: "text", width: 140 },
  { key: "gmatMedio", header: "GMAT Médio", group: "Admissão", type: "number", width: 100 },
  { key: "taxaAceitacao", header: "Taxa Aceitação %", group: "Admissão", type: "percent", width: 120 },
  { key: "experienciaProfissional", header: "Experiência Profissional", group: "Admissão", type: "text", width: 170 },
  { key: "materiasPreRequisito", header: "Matérias Pré-Requisito", group: "Admissão", type: "longtext", width: 200 },
  { key: "testesIngles", header: "Testes Inglês Aceitos", group: "Admissão", type: "text", width: 160 },
  { key: "deadlineAplicacao", header: "Deadline Aplicação", group: "Admissão", type: "text", width: 140 },
  { key: "taxaAplicacao", header: "Taxa Aplicação USD", group: "Admissão", type: "currency", width: 130 },

  { key: "roundsAplicacao", header: "Rounds Aplicação", group: "Aplicação", type: "text", width: 130 },
  { key: "documentosNecessarios", header: "Documentos Necessários", group: "Aplicação", type: "longtext", width: 220 },
  { key: "cartasRecomendacao", header: "Cartas de Recomendação", group: "Aplicação", type: "text", width: 160 },
  { key: "entrevista", header: "Entrevista", group: "Aplicação", type: "text", width: 120 },
  { key: "dificuldadeEntrada", header: "Dificuldade Entrada", group: "Aplicação", type: "text", width: 140 },
];

export const COLUMN_GROUPS = Array.from(new Set(COLUMNS.map((c) => c.group)));

export const COLUMN_MAP: Record<string, ColumnDef> = Object.fromEntries(
  COLUMNS.map((c) => [c.key, c]),
);

export const NUMERIC_TYPES: ColumnType[] = ["number", "percent", "currency"];
