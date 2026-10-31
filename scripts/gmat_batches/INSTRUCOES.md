# Pesquisa de GMAT médio dos admitidos (lotes)

Contexto: candidato internacional (visto F-1) que quer definir a meta de GMAT.

- Entrada: `scripts/gmat_batches/input_gmat_NN.json` (programas sem `GMAT Médio` numérico).
- Saída: `scripts/gmat_batches/output_gmat_NN.json`, lista na ordem da entrada, gravada após CADA programa
  (se o arquivo já existir, carregue e pule os que já estão nele). NUNCA abra/salve o xlsx.
- Arquivos temporários: SOMENTE na sua subpasta exclusiva do scratchpad indicada no prompt.

Formato de cada item:
```json
{
  "match": {"Nome da Universidade": "...", "Estado": "..", "Cidade": "..."},
  "gmat_medio": 640 | null,
  "gmat_escala": "classico" | "focus" | null,
  "gmat_faixa": "texto curto, ex.: 'middle 80%: 600-720'" | null,
  "gre_medio": "ex.: 'Q 162 / V 156' ou 'total 318'" | null,
  "ano_turma": "ex.: 'Class of 2026' / 'Fall 2025'" | null,
  "pct_com_teste": "ex.: '45% enviaram GMAT/GRE'" | null,
  "fontes": ["https://..."],
  "observacoes": "curto, em PT"
}
```

Regras:
- Fonte: SÓ página oficial do programa (class profile, admissions, FAQ, employment/admissions report, PDF oficial).
  Agregadores (Yocket, Shiksha, mastersportal, Poets&Quants etc.) não valem — no máximo como pista.
- `gmat_medio` = média OU mediana do PROGRAMA (diga qual em observacoes). Não use dado do MBA nem da escola inteira;
  se só houver dado da escola/de todos os MS, deixe `gmat_medio` null e descreva em observacoes.
- `gmat_escala`: "focus" se a escola diz GMAT Focus Edition ou o número só faz sentido nessa escala; senão "classico".
- Dado mais recente publicado. Não invente; null quando não achar.
- Sites que bloqueiam WebFetch: curl -sL -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36" URL -o arquivo, extraia texto com Python.
- No máximo ~3 min por programa; muitos não publicam GMAT — registre null e siga.

Ao terminar: resumo curto (quantos com GMAT achado, lista programa → valor/escala).
