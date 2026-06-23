# Design: Relatório de Processos — VBA Excel

**Data:** 2026-06-21  
**Projeto:** Vahzanatta  
**Escopo:** Novo módulo VBA (`RelatorioProcessos.bas`) que cruza um relatório de processos contra dois relatórios SAP (RESGATES e MOVIMENTO), identificando onde cada processo foi localizado e anotando divergências.

---

## 1. Estrutura das abas

### Aba `PROCESSOS` — Dados de entrada + saída

| Coluna | Índice | Campo |
|--------|--------|-------|
| A | 1 | PROCESSO LOCALIZADO |
| B | 2 | DATA DO CRÉDITO |
| C | 3 | VALOR RECUPERADO |
| D | 4 | CRÉDITO ÚNICO |
| E | 5 | AGÊNCIA E CONTA |
| F | 6 | *(nova)* Localizado em |
| G | 7 | *(nova)* Correspondência |
| H | 8 | *(nova)* Retorno |

- Linha 1: cabeçalho; dados a partir da linha 2
- Colunas F:H escritas pela macro, incluindo cabeçalho na linha 1

### Abas `RESGATES` e `MOVIMENTO` — Mesma estrutura da aba SAP (A:K)

| Coluna | Índice | Campo relevante |
|--------|--------|-----------------|
| D | 4 | Data de documento |
| H | 8 | Valor em moeda da empresa |
| K | 11 | Texto |

- Linha 1: cabeçalho; dados a partir da linha 2
- Apenas leitura — nenhuma coluna é escrita nessas abas

---

## 2. Lógica de busca

Para cada linha do PROCESSOS, a macro executa dois estágios em sequência, encerrando assim que uma correspondência é encontrada.

### Estágio 1 — Busca no RESGATES

**Passo 1a — Por número do processo** (somente se PROCESSO LOCALIZADO não estiver em branco):

- Percorre o array do RESGATES procurando `InStr(Texto_j, ProcessoLocalizado, vbTextCompare) > 0`
- O campo Texto pode conter outros dados além do número do processo — basta que o contenha
- Se encontrado:
  - **Localizado em:** `"Resgate"`
  - **Correspondência:** `"Número do processo"`
  - **Retorno:** `"Localizado"` se diferença de dias = 0, senão `"Localizado — divergência de N dias"`
  - Encerra busca para esta linha

**Passo 1b — Por valor** (somente se não localizado no Passo 1a):

- Percorre o RESGATES comparando Valor em moeda da empresa (numérico) com VALOR RECUPERADO e com CRÉDITO ÚNICO
- VALOR RECUPERADO tem prioridade: se bater, usa `"Valor recuperado"`; senão testa CRÉDITO ÚNICO: `"Crédito único"`
- Se encontrado:
  - **Localizado em:** `"Resgate"`
  - **Correspondência:** `"Valor recuperado"` ou `"Crédito único"`
  - **Retorno:** `"Localizado"` ou `"Localizado — divergência de N dias"`
  - Encerra busca para esta linha

### Estágio 2 — Busca no MOVIMENTO (somente para linhas não localizadas no Estágio 1)

- Percorre o MOVIMENTO comparando Valor em moeda da empresa com VALOR RECUPERADO e com CRÉDITO ÚNICO (mesma prioridade do Passo 1b)
- Se encontrado:
  - **Localizado em:** `"Movimento"`
  - **Correspondência:** `"Valor recuperado"` ou `"Crédito único"`
  - **Retorno:** `"Localizado — sem número de processo vinculado"` + `", divergência de N dias"` se houver diferença de data
- Se não encontrado em nenhum dos dois relatórios:
  - **Localizado em:** *(vazio)*
  - **Correspondência:** *(vazio)*
  - **Retorno:** `"Não localizado"`

### Divergência de dias

- `N = Abs(CDate(DataDocRelatorio) - CDate(DataCreditoProcesso))`
- Calculada em todos os casos de localização, independentemente da correspondência (processo ou valor)
- Se DATA DO CRÉDITO estiver em branco ou inválida: divergência não é calculada e o Retorno omite essa informação

---

## 3. Macros do módulo

| Macro | Visibilidade | Descrição |
|-------|-------------|-----------|
| `AnalisarProcessos` | Public | Macro principal — valida abas, carrega arrays, executa busca em dois estágios, escreve F:H em lote |
| `LimparAnalise` | Public | Limpa F1:H(últimaLinha) da aba PROCESSOS para permitir reprocessamento |

### Fluxo de `AnalisarProcessos`

1. **Validação:** verifica existência das abas PROCESSOS, RESGATES e MOVIMENTO; verifica dados a partir da linha 2 em cada uma; exibe erro e encerra se algo faltar
2. **Carregamento:** lê PROCESSOS (A:H), RESGATES (A:K) e MOVIMENTO (A:K) em arrays Variant
3. **Processamento:** para cada linha do PROCESSOS, executa Estágio 1 depois Estágio 2 conforme a seção 2
4. **Escrita em lote:** escreve cabeçalhos em F1:H1 e resultados em F2:H(últimaLinha) em uma única operação
5. **Mensagem final:** exibe total de linhas analisadas, localizados no RESGATES, localizados no MOVIMENTO e não localizados

### Mensagem final

```
Concluído!

Linhas analisadas:        N
Localizados em Resgate:   N
Localizados em Movimento: N
Não localizados:          N
```

---

## 4. Decisões de design

- **Módulo separado:** `RelatorioProcessos.bas` é um arquivo novo — nenhum módulo ou manual existente é alterado.
- **Array em memória:** mesmo padrão do `RelatorioPendencias` — zero leituras de célula durante o processamento.
- **Comparação de valor numérica:** usa `CDbl` para evitar falsos negativos por diferença de formato de string.
- **Prioridade de correspondência:** processo > valor recuperado > crédito único; RESGATES > MOVIMENTO.
- **InStr case-insensitive:** `vbTextCompare` em todas as comparações de texto.
- **Escrita em lote:** array de saída escrito nas colunas F:H em uma única atribuição `.Value`.
