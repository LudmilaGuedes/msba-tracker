# MSBA Tracker

> **EN —** An interactive dashboard to compare **138 on-campus, STEM-designated Master of Science in Business Analytics (MSBA) programs in the US** by *real total cost* (tuition + living expenses), scholarships/assistantships and employability — not just by ranking. The 55-field dataset was researched from official university sources and the app was built end-to-end with **Claude Code** (AI-assisted "vibe coding"), with every data change scripted and auditable.
>
> **Demo:** **https://msba-tracker.vercel.app** · **Stack:** Vite · React · TypeScript

---

Escolher um mestrado nos EUA olhando só o ranking esconde o que mais pesa para um estudante internacional: **quanto o programa custa de verdade** (tuition + aluguel + seguro + custo de vida durante todo o curso), **qual a chance real de bolsa ou assistantship** e **qual o retorno em empregabilidade e salário**. Essas informações ficam espalhadas por dezenas de páginas de cada universidade. O MSBA Tracker junta tudo num só lugar, com os mesmos critérios para os 138 programas.

**Demo online:** **https://msba-tracker.vercel.app**

## Funcionalidades

**Comparação**
- **Custo Total Real** de cada programa: tuition total (internacional) + custo de vida pela duração do curso (aluguel de quarto individual, mercado, seguro saúde, seguro residencial, utilities, celular, internet, transporte e outros).
- **Gráfico Custo × Salário** (dispersão) para achar os programas com melhor custo-benefício.
- **55 campos por programa** em 7 grupos: Universidade, Programa, Custos, Bolsas e Auxílios, Rankings e Empregabilidade, Requisitos de Admissão e Processo de Aplicação.
- Duas visualizações: **cards** e **tabela** com ordenação por qualquer coluna, além de um modal de detalhes com todos os campos.

**Filtros**
- Busca por nome, estado (múltipla seleção) e cidade.
- Faixas de **duração (meses)**, **tuition total internacional** e **custo total real**.
- Apenas STEM, apenas favoritos (5 estrelas), apenas com anotações, por pasta.

**Organização pessoal**
- Nota de 0 a 5 estrelas, comentários e pastas coloridas por programa.
- Edição inline de células e cadastro de novos programas. As alterações ficam numa camada separada, e a planilha original nunca é modificada.
- Exportação para Excel da visão filtrada (com nota, comentários e pastas), backup automático e restauração.

**Estratégia de Aplicação**
- Agrupa os programas em **Round 1 (set–dez), Round 2 (jan–mai) e Round 3 (jun–ago)** pelo prazo mais cedo de cada um, para planejar *quando* aplicar.
- Por grupo: taxas de aplicação, envio de GMAT, cartas de recomendação, aceitação do Duolingo (DET), dificuldade de entrada e data prevista de resultado.
- Resumo global do custo do processo de aplicação (taxas + GMAT), filtrável por "aplicar com certeza" ou "em dúvida".

**Dados ausentes nunca viram zero:** campos não confirmados em fonte oficial aparecem como `Verificar` e são tratados como "desconhecido" em ordenações, filtros e somas.

## Como a IA foi usada

O projeto foi construído com o **[Claude Code](https://claude.com/claude-code)** no modo *vibe coding*: eu defini o problema, as regras e os critérios de qualidade, e a IA escreveu o código e os scripts de dados. O meu trabalho foi especificar, revisar e validar.

**1. Especificação antes do código.** Antes de qualquer linha de código, defini o schema de 55 campos em 7 categorias e as regras de negócio que a IA precisava seguir:
- tuition sempre pelo valor **total** do programa (créditos × custo por crédito), nunca por semestre;
- custo de vida padronizado (quarto individual + banheiro privativo, mais as mesmas categorias de gasto para todos os programas);
- dados não encontrados ficam como `Verificar`: **nunca em branco e nunca inventados**;
- só valem fontes oficiais (sites das universidades, catálogos, páginas de tuition, employment reports). Agregadores servem no máximo como pista.

Essas regras ficam em [`resumo_contexto_tecnico.md`](resumo_contexto_tecnico.md) e no [`CLAUDE.md`](CLAUDE.md), que funciona como especificação permanente: a cada sessão, a IA relê as convenções do projeto (schema como fonte única da verdade, regra dos prazos, política de backup) em vez de depender da memória da conversa.

**2. Coleta de dados iterativa.** Os dados foram levantados em rodadas de pesquisa: primeiro os programas mais bem ranqueados e depois cada região dos EUA, até cobrir o país. Seguiram-se rodadas de "lacunas" e um *data enrichment* em lote, com agentes em paralelo trabalhando sobre arquivos JSON e regras escritas ([`scripts/enrichment_batches/INSTRUCOES.md`](scripts/enrichment_batches/INSTRUCOES.md)). O filtro de entrada é **presencial + STEM-designated + aberto a visto F-1**.

**3. Toda mudança nos dados é um script auditável.** Nenhuma célula foi editada à mão. Cada correção virou um script Python em [`scripts/`](scripts/) (257 no total), que localiza a linha por *Universidade + Estado + Cidade* (nunca pelo número da linha), altera só os campos necessários e documenta as fontes no próprio arquivo. Assim, qualquer valor da planilha pode ser rastreado até a fonte e o motivo da mudança.

**4. Validação.** Eu não aceitei o código ou os dados gerados sem conferir:
- `npm run build` faz a checagem de tipos do TypeScript a cada mudança, e os testes funcionais foram feitos manualmente no navegador;
- fiz rodadas de auditoria contra as fontes oficiais (`etl_fix_verificacao_oficial_*`), começando pelos programas mais baratos e mais caros, onde um erro distorce mais a comparação;
- revisei programas um a um e removi os que não cumpriam os critérios (online, não-STEM ou sem MSBA de fato).

**Exemplos de problemas encontrados e corrigidos durante a iteração:**
- O ID de cada programa era a posição da linha, então remover duas faculdades da planilha deslocou os comentários do usuário para as faculdades erradas. O ID passou a ser estável (`nome|estado|cidade`), com uma migração única dos dados antigos.
- Textos gravados em colunas numéricas viravam números sem sentido no dashboard. A correção foi feita com um script de limpeza, e a regra foi adicionada às instruções.
- O `localStorage` é separado por porta. Se o servidor de desenvolvimento mudasse de porta, os dados do usuário pareceriam sumir, então a porta foi fixada.
- O pacote `xlsx` do npm tem vulnerabilidades sem correção. Ele é instalado direto do CDN oficial do SheetJS.

## Impacto

A ferramenta economizou aproximadamente **125 horas** de pesquisa manual.

**Cálculo:** levantar e conferir os 55 campos de um programa exige abrir de 4 a 6 páginas oficiais (tuition, custo de vida, admissão, bolsas, *employment report*), o que leva cerca de **45 minutos por programa**:

| Etapa | Tempo |
|---|---|
| 45 min × 138 programas | ~104 h |
| Comparação manual em planilha e revisão | ~20 h |
| **Total** | **~125 h** |

Além do tempo, a ferramenta passou a mostrar o que antes não aparecia: a diferença de custo total entre programas com ranking parecido, e quais deles oferecem um caminho realista de assistantship ou tuition reduzida para internacionais.

## Fonte dos dados

O dataset fica em [`data/msba_programs_database.xlsx`](data/msba_programs_database.xlsx) (aba `MSBA Programs`, 138 programas × 55 colunas). O dashboard lê uma cópia em `msba-dashboard/public/data/`.

- **Origem:** páginas oficiais de cada universidade e business school (catálogos, tuition e fees, *cost of attendance*, requisitos de admissão, *employment reports*), mais os rankings U.S. News e QS. Os scripts em `scripts/` registram as URLs consultadas.
- **Período de coleta:** agosto a outubro de 2026, com valores do ano acadêmico 2026–27 ou do ciclo de admissão mais recente publicado.
- **Escopo:** apenas programas **presenciais**, **STEM-designated** e abertos a estudantes com visto **F-1**.
- **Limitações:** tuition, prazos e requisitos mudam todo ano, e custo de vida e chance de bolsa são estimativas. Campos marcados como `Verificar` não foram confirmados em fonte oficial. **Confirme sempre no site da universidade antes de tomar decisões.**

## Stack

- **Frontend:** [Vite](https://vite.dev/) 6 · [React](https://react.dev/) 19 · [TypeScript](https://www.typescriptlang.org/) · [Tailwind CSS](https://tailwindcss.com/) 4
- **Gráficos e ícones:** [Recharts](https://recharts.org/) · [Lucide](https://lucide.dev/)
- **Leitura/escrita de Excel no navegador:** [SheetJS](https://sheetjs.com/) (a planilha é lida em tempo de execução, sem backend)
- **Pipeline de dados:** Python + [openpyxl](https://openpyxl.readthedocs.io/)
- **Persistência:** `localStorage` (notas, comentários, pastas e programas adicionados ficam no navegador de cada usuário)
- **Desenvolvimento:** [Claude Code](https://claude.com/claude-code)

## Como rodar localmente

Pré-requisito: [Node.js](https://nodejs.org/) 20 ou superior.

```bash
git clone https://github.com/LudmilaGuedes/msba-tracker.git
cd msba-tracker/msba-dashboard
npm install
npm run dev        # abre em http://localhost:5173
```

Outros comandos (dentro de `msba-dashboard/`):

```bash
npm run build      # checagem de tipos + build de produção em dist/
npm run preview    # serve o build de produção
npm run lint       # oxlint
```

Para atualizar os dados, rode os scripts Python a partir da raiz do repositório (`pip install openpyxl`). Depois, copie `data/msba_programs_database.xlsx` para `msba-dashboard/public/data/`.

## Estrutura do repositório

```
data/                  planilha-fonte (138 programas × 55 campos)
msba-dashboard/        app React (Vite + TypeScript)
  src/types/columns.ts   schema dos 55 campos (fonte única da verdade)
scripts/               scripts ETL de coleta, correção e auditoria dos dados
resumo_contexto_tecnico.md   schema, metodologia e histórico das fases
CLAUDE.md              especificação/convenções lidas pela IA a cada sessão
```

## Licença

Código sob a licença [MIT](LICENSE).
