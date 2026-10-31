# Resumo de Contexto Técnico — Projeto MSBA Tracker

## Objetivo do Projeto
Criar uma ferramenta automatizada para simplificar a busca por programas de MSBA (Master of Science in Business Analytics) em universidades americanas para estudantes internacionais. O projeto substitui o processo manual de entrar site por site de universidades, consolidando todos os dados relevantes num banco de dados estruturado e num dashboard interativo filtrável.

## Fases do Projeto

### Fase 1 — Definição do Schema ✅ CONCLUÍDA
Definimos uma estrutura de dados com 54 campos (55 desde out/2026) organizados em 7 categorias para cada programa MSBA:

**1. Universidade (5 campos):** Nome da Universidade, Estado, Cidade, Business School, Tipo (Pública/Privada)

**2. Programa MSBA (7 campos):** Nome Oficial do Programa, STEM Designated (Sim/Não), Formato (presencial), Duração (meses), Início (semestres), Tamanho da Turma, Destaques do Currículo

**3. Custos (17 campos):**
- Total de Créditos para Formatura
- Custo por Crédito (Internacional)
- Custo por Crédito (Residente)
- Tuition Total do Programa (Internacional) — calculado como créditos × custo/crédito
- Tuition Total do Programa (Residente)
- Fees Adicionais
- Aluguel (quarto individual + banheiro privativo)/mês
- Mercado/Alimentação por mês
- Seguro Saúde (total do programa)
- Seguro Residencial (renter's insurance)/mês
- Água, Luz, Energia, Gás/mês
- Plano de Celular/mês
- Internet (Wi-Fi)/mês
- Transporte/mês
- Outros Gastos Necessários/mês
- Custo de Vida Total (duração do programa)

**4. Bolsas e Auxílios (7 campos):** Bolsas Disponíveis, Tipos de Bolsa, Quão Realista é Conseguir (Intl), TA/RA Disponível, Processo para TA/RA, Caminho para Tuition de Residente, Graduate Assistantships

**5. Rankings e Qualidade (5 campos):** Ranking US News, Ranking QS, Taxa de Empregabilidade %, Salário Médio Pós-Programa USD, Principais Empregadores

**6. Requisitos de Admissão (10 campos):** GPA Mínimo, GPA Médio Admitidos, GMAT/GRE Obrigatório, GMAT Médio, Taxa de Aceitação %, Experiência Profissional, Matérias Pré-Requisito, Testes de Inglês Aceitos (com scores mínimos), Deadline de Aplicação, Taxa de Aplicação

**7. Processo de Aplicação (5 campos):** Rounds de Aplicação, Documentos Necessários, Cartas de Recomendação (quantidade; campo adicionado em out/2026), Entrevista, Dificuldade de Entrada

**Regras de custo importantes:**
- Tuition deve ser SEMPRE o valor TOTAL do programa completo, nunca por semestre
- Quando a faculdade mostra preço por crédito, multiplicamos pelo total de créditos
- Custo de vida deve considerar: aluguel com quarto individual + banheiro privativo, mercado, seguro saúde, seguro residencial, utilities, celular, wifi, transporte, outros gastos
- Dados faltantes são marcados como "Verificar" (nunca em branco, nunca inventados)

### Fase 2 — Coleta e Consolidação de Dados ✅ CONCLUÍDA
Realizamos 6 rodadas de Deep Research cobrindo todos os 50 estados dos EUA:
- Rodada 1: Top 15-20 programas por ranking (MIT, UCLA, USC, UT Austin, UC Davis, UCSD, Minnesota, Purdue, Rochester, Notre Dame, UT Dallas, Texas A&M, Utah, Michigan State, Cincinnati, Rutgers)
- Rodada 2: Sul (FL, GA, NC, SC, VA, TN, AL, LA, MS, KY, AR)
- Rodada 3: Nordeste (NY, MA, CT, PA, NJ, MD, DE, DC)
- Rodada 4: Centro-Oeste (IL, OH, MI, IN, WI, IA, MO, KS, NE, ND, SD, MN)
- Rodada 5: Oeste (CA, WA, OR, CO, AZ, NV, NM, HI, MT, WY, ID)
- Rodada 6: Lacunas (TX restante, OK, programas Midwest não cobertos)

**Resultado: Planilha Excel com 82 programas qualificados** (presenciais, STEM-designated, abertos a F-1), incluindo dados de tuition, bolsas, requisitos, rankings e custo de vida. Campos sem dados confirmados estão marcados "Verificar" com fundo amarelo.

**Arquivo gerado:** `msba_programs_database.xlsx` — 2 abas (Programs + Resumo), 54 colunas, 82 linhas de dados, filtros automáticos, colunas congeladas.

### Fase 3 — Construção do Dashboard/App ⬅️ PRÓXIMO PASSO
Usar o Claude Code (terminal do VS Code) para construir um app web interativo (React) que permita:
- Filtrar programas por: custo total, estado, ranking, exigência de GMAT, disponibilidade de bolsa/TA, teste de inglês aceito, duração, dificuldade de entrada
- Ordenar por qualquer campo (custo total, salário pós, taxa de empregabilidade, etc.)
- Visualizar comparações lado a lado entre programas selecionados
- Calcular custo total real (tuition + custo de vida) para cada programa
- Destacar os "best value" (melhor relação custo-benefício considerando salário pós e empregabilidade)

**Dados de entrada:** O arquivo `msba_programs_database.xlsx` serve como banco de dados. O dashboard deve ler esse arquivo e renderizar as visualizações.

### Fase 4 — Skills e Automação (futuro)
Criar Skills personalizadas para que futuras pesquisas de programas sigam o mesmo schema automaticamente, permitindo atualização contínua do banco de dados.

## Perfil do Usuário-alvo
- Estudante internacional buscando MSBA nos EUA
- Precisa de visto F-1 (apenas programas presenciais qualificam)
- STEM designation é obrigatório (para 36 meses de OPT)
- Sensível a custos — quer ver custo total real (tuition + vida) e oportunidades de bolsa/TA
- Quer comparar programas por custo-benefício, não apenas ranking

## Stack Técnico Sugerido para Fase 3
- React (JSX) para o frontend
- Tailwind CSS para estilização
- Recharts ou similar para gráficos comparativos
- Dados carregados do XLSX via papaparse/SheetJS
- Filtros interativos, tabela ordenável, cards de comparação
