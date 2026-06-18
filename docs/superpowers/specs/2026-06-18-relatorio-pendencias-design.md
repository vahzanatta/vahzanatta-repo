# Design: Relatório de Pendências — VBA Excel

**Data:** 2026-06-18  
**Projeto:** Vahzanatta  
**Escopo:** Macro VBA para processar partidas em aberto exportadas do SAP (FBL3H), identificar a área responsável por cada lançamento e adicionar colunas calculadas ao relatório.

---

## 1. Estrutura das abas

### Aba `SAP` — Dados do relatório FBL3H

| Coluna | Campo |
|--------|-------|
| A | Empresa |
| B | Conta do Razão |
| C | Data de documento |
| D | Data de lançamento |
| E | Nº documento |
| F | Tipo de documento |
| G | Valor em moeda da empresa |
| H | Chave de lançamento |
| I | Atribuição |
| J | Texto |
| K | *(nova)* Data Base |
| L | *(nova)* Dias em Aberto |
| M | *(nova)* Área Responsável |

- Linha 1: cabeçalho
- Dados a partir da linha 2
- Cabeçalhos das colunas K, L e M são escritos automaticamente pela macro

### Aba `BASE` — Tabela de regras de correspondência

| Coluna | Campo |
|--------|-------|
| A | Conta do Razão |
| B | Começa em |
| C | Contém |
| D | Área Responsável |

- Linha 1: cabeçalho
- Regras a partir da linha 2

---

## 2. Lógica de correspondência e pontuação

Para cada linha do SAP, a macro percorre todas as regras da BASE e calcula um **score** de 0 a 3:

| Condição | Pontos |
|----------|--------|
| Conta do Razão da BASE está em branco **ou** bate com a do SAP | +1 |
| Campo "Começa em" não está vazio **e** o Texto do SAP começa com ele | +1 |
| Campo "Contém" não está vazio **e** o Texto do SAP contém ele | +1 |

### Regras de eliminação (score zerado, regra descartada)

- Conta do Razão da BASE está preenchida **e** não bate com a do SAP → descarta
- "Começa em" está preenchido **e** o Texto do SAP **não** começa com ele → descarta
- "Contém" está preenchido **e** o Texto do SAP **não** contém ele → descarta

### Desempate

A regra com **maior score** vence. Em caso de empate, prevalece a **primeira na ordem da BASE**.

### Casos especiais

- Conta do Razão em branco na BASE → a regra vale para qualquer conta (match por texto apenas)
- "Começa em" e "Contém" ambos em branco na BASE → match apenas por Conta do Razão
- Nenhuma regra bate → Área Responsável recebe `"⚠ Não identificado"`

Todas as comparações de texto são **case-insensitive**.

---

## 3. Fluxo de processamento

A macro principal `ProcessarPendencias` executa em 4 etapas:

### Etapa 1 — Validação inicial
- Verifica se as abas "SAP" e "BASE" existem
- Verifica se há dados a partir da linha 2 em ambas
- Exibe mensagem de erro e encerra se algo estiver faltando

### Etapa 2 — Exclusão de linhas inválidas
- Percorre a aba SAP **de baixo para cima** para não deslocar índices
- Exclui todas as linhas onde a coluna F (Tipo de documento) estiver em branco

### Etapa 3 — Carregamento em array
- Carrega todo o intervalo de dados do SAP (colunas A:J) em um array Variant na memória
- Carrega toda a BASE (colunas A:D) em um array Variant na memória
- A partir deste ponto, nenhuma leitura de célula ocorre

### Etapa 4 — Processamento e escrita
- Para cada linha do array SAP, percorre todas as regras do array BASE calculando o score
- Armazena os resultados (Data Base, Dias em Aberto, Área Responsável) em um array de saída
- Ao final, escreve o array de saída inteiro nas colunas K:M em uma única operação

---

## 4. Colunas de saída

### Coluna K — Data Base
- Preenchida com a data atual (`Date`) em todas as linhas processadas
- Formatada como data curta (dd/mm/aaaa)

### Coluna L — Dias em Aberto
- Calculada como `Data Base − Data de lançamento (col D)`
- Resultado em número inteiro de dias
- Se Data de lançamento estiver em branco ou inválida: célula recebe `"—"`

### Coluna M — Área Responsável
- Valor da coluna D da BASE para a regra vencedora
- Se nenhuma regra bater: `"⚠ Não identificado"`
- Sem formatação de fundo ou cor

---

## 5. Macros do módulo

| Macro | Descrição |
|-------|-----------|
| `ProcessarPendencias` | Macro principal — executa as 4 etapas descritas acima |
| `LimparResultados` | Apaga as colunas K:M e os cabeçalhos para permitir reprocessamento |

---

## 6. Mensagem final

Ao concluir, `ProcessarPendencias` exibe uma caixa de mensagem com:

- Linhas analisadas
- Total atribuídos
- Total não identificados

---

## 7. Funções auxiliares

| Função | Assinatura | Descrição |
|--------|-----------|-----------|
| `CalcularScore` | `(contaSAP, textoSAP, contaBase, comecaEm, contem) As Integer` | Retorna o score (0–3) de uma regra para uma linha do SAP |

---

## 8. Decisões de design

- **Array em memória:** necessário dado o volume de 100k–250k linhas; leituras célula a célula seriam inviáveis em tempo.
- **Exclusão de baixo para cima:** padrão obrigatório ao deletar linhas em loop para evitar salto de índice.
- **Case-insensitive:** uso de `UCase` nas comparações para garantir consistência independentemente de como o dado foi digitado no SAP ou na BASE.
- **Sem formatação de saída:** colunas K:M recebem apenas valores, sem cores ou estilos.
