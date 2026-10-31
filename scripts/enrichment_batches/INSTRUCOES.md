# Instruções para os agentes de Data Enrichment (lotes)

Você é um engenheiro de dados auditando e completando programas de MSBA (presencial, STEM) para uma
candidato **internacional** (visto F-1, diploma de graduação estrangeiro) que vai aplicar para o próximo ciclo de admissão.

## Entrada e saída

- Entrada: `scripts/enrichment_batches/input_batch_NN.json`, com os valores ATUAIS das 54 colunas de cada linha.
- Saída: `scripts/enrichment_batches/output_batch_NN.json`. **NUNCA abra nem salve o `msba_programs_database.xlsx`.**
  Um script único aplica todos os lotes depois.
- Formato da saída (lista, um objeto por linha da entrada, na mesma ordem):

```json
[
  {
    "match": {"Nome da Universidade": "...", "Estado": "..", "Cidade": "..."},
    "updates": {"<cabeçalho exato da coluna>": <novo valor>, "...": "..."},
    "det_aceito": "sim" | "nao" | "nao_confirmado",
    "det_nota_minima": 120 | null,
    "fontes": ["https://...", "..."],
    "observacoes": "curto, em PT: o que foi corrigido, ressalvas, conflitos entre fontes"
  }
]
```

- `match` copia EXATAMENTE `Nome da Universidade`, `Estado` e `Cidade` da entrada (nunca altere essas três colunas).
- `updates` contém APENAS as colunas que mudam. Use os cabeçalhos exatamente como aparecem na entrada (com acentos).
- Salve o JSON com UTF-8. Grave o arquivo incrementalmente (após cada universidade) para não perder trabalho.

## Como pesquisar

- **Fontes oficiais primeiro** (site da universidade/business school, catálogo, páginas de tuition/bursar, cost of
  attendance, employment report oficial, U.S. News/QS para rankings). Agregadores (Yocket, Shiksha, Collegedunia,
  mastersportal, etc.) NÃO valem como fonte para gravar um dado; no máximo como pista para achar a página oficial.
- Muitos sites .edu bloqueiam o WebFetch (403). Nesse caso baixe com curl via Bash usando um User-Agent de navegador:
  `curl -sL -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36" URL -o arquivo.html`
  e extraia o texto com Python (remova tags). Salve arquivos temporários no seu scratchpad, nunca no repositório.
  PDFs: `pypdf` está instalado.
- Use os dados mais recentes publicados (ano acadêmico 2026-27 ou o ciclo de admissão mais recente, geralmente Fall 2027).
- Não invente. Se não encontrar em fonte confiável, deixe como está (ou `"Verificar"` se o valor atual estiver errado
  e não houver substituto confiável).
- Se o valor atual já está certo, não o inclua em `updates`.
- Tempo: seja eficiente, ~15-25 min por lote no máximo; priorize os campos abaixo.

## Campos prioritários (nesta ordem)

1. **Testes de inglês / Duolingo (DET)**: confirme se o programa (não só a universidade de graduação) aceita o DET e a nota
   mínima. Preencha `det_aceito` e `det_nota_minima`. Na coluna `Testes Inglês Aceitos`, liste mínimos de TOEFL/IELTS/DET/PTE.
   Se a página oficial do programa ou da pós-graduação lista só TOEFL/IELTS, `det_aceito = "nao"`.
2. **Tuition**: custo por crédito e total, residente e internacional (públicas: non-resident = internacional).
   Programas com preço fechado (lump sum): coloque o total e deixe o por crédito calculado ou "Verificar".
3. **Formato** (Presencial/Híbrido/Online) e **Duração (meses)** reais.
4. **Bolsas**: `Bolsas Disponíveis` (Sim/Não), `Tipos de Bolsa` (quais existem e valores/faixas publicados),
   `Quão Realista (Intl)` (começar com Baixa / Moderada / Alta + justificativa curta: internacionais são elegíveis?
   valor típico?), `TA/RA Disponível`, `Processo TA/RA`, `Graduate Assistantships`.
5. **Empregabilidade** (`Taxa Empregabilidade %`), **Salário Médio Pós USD**, `Principais Empregadores`
   (de relatórios oficiais do programa; diga no texto se o dado é da escola inteira e não do programa).
6. **Dificuldade Entrada** (use: Acessível / Moderada / Moderada-Alta / Alta / Muito Alta), `Taxa Aceitação %`,
   `GPA Médio Admitidos`, `GMAT Médio`, `GPA Mínimo`.
7. **Taxa Aplicação USD** (e isenções: coloque a isenção em `observacoes` e/ou no texto de `Rounds Aplicação`).
8. **Deadline Aplicação** e **Rounds Aplicação** (datas prioritárias).
9. **Requisitos**: `GMAT/GRE Obrigatório` (Sim / Não / Opcional / Waiver disponível + detalhe),
   `Documentos Necessários` (inclua **quantas cartas de recomendação**), `Experiência Profissional`,
   `Matérias Pré-Requisito`, `Entrevista`.
10. Se sobrar tempo: `Início`, `Tamanho Turma`, `Total Créditos`, `Fees Adicionais USD`, `Seguro Saúde (programa) USD`,
    aluguel/transporte/outros do cost of attendance oficial, `Ranking US News`, `Ranking QS`, `Destaques Currículo`.

## Regras do schema (obrigatórias)

- **Colunas numéricas recebem SÓ um número (int/float) ou a string "Verificar"**. Nada de texto com dígitos
  (o dashboard remove não-dígitos e "#107 (2025)" vira 1072025). Colunas numéricas:
  `Duração (meses)`, `Tamanho Turma`, `Total Créditos`, todas as colunas terminadas em `USD`,
  `Ranking US News`, `Ranking QS`, `Taxa Empregabilidade %`, `GPA Mínimo`, `GPA Médio Admitidos`, `GMAT Médio`,
  `Taxa Aceitação %`. Percentuais como número (ex.: 92, não "92%" nem 0.92). Rankings: só o número da posição
  do programa (de Business Analytics, se houver); contexto vai em `observacoes`.
- **NÃO altere `Mercado/mês USD`**: é um valor fixo de 300 por decisão da usuária.
- `Estado`: sigla (enum). `Tipo (Pub/Priv)`: "Pública" ou "Privada". `STEM Designated` e `Bolsas Disponíveis`: "Sim" ou "Não".
- Dado ausente = string "Verificar" (nunca vazio/null).
- Textos em português, concisos, separados por ";". Valores em dólar escritos como "US$1,200".
- **Deadline Aplicação**: comece pelo PRAZO MAIS CEDO do intake principal (Round 1 / Early / prioridade de bolsa;
  internacional antes de doméstico quando o internacional é mais cedo; Fall antes de Spring). Nada antes dele (nem mês de
  início, nem "aplicações abrem ..." — data de abertura vai no fim). Com rounds, use
  `Round 1: 1 nov 2026; Round 2: 15 jan 2027; ...` com algarismos arábicos e meses abreviados em PT
  (jan, fev, mar, abr, mai, jun, jul, ago, set, out, nov, dez).
- Não mude `Nome da Universidade`, `Estado`, `Cidade`. Pode corrigir `Nome Oficial do Programa`/`Business School` se estiver errado.
- Se descobrir que o programa não existe mais, é só online, ou não é STEM, NÃO remova — diga em `observacoes`
  (começando com "ALERTA:").

## Ao terminar

Responda com um resumo curto: por universidade, quantos campos mudaram, status do DET, e qualquer ALERTA.
