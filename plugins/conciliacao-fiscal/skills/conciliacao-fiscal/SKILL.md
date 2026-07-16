# Conciliação Fiscal de Armazenagem

## Resumo do Processo

Esta skill executa a conciliação fiscal de armazenagem, rastreando quantidades e ICMS de retornos contra remessas. Compara dados do ZSD (SAP) com XMLs de retorno para validar que o que foi devolvido bate com o que foi enviado.

**Resultado:** Planilha com análise de remessas vs retornos, mostrando saldos pendentes e divergências.

## Arquivos e Componentes

O script reside em `scripts/` dentro da skill:
- `concilia_fiscal.py` — executa conciliação fiscal com rastreamento de quantidades

A skill funciona com qualquer armazenagem que siga o layout de colunas esperado, exigindo sempre:
1. Arquivo Excel com abas `zsd` e `razao`
2. Pasta de XMLs de retorno

## Fluxo de Execução

**1. Identificar inputs:** pergunte qual arquivo Excel e qual pasta de XMLs usar (NUNCA assuma automaticamente)
**2. Executar conciliação:** rode `concilia_fiscal.py` com os parâmetros `--arquivo=` e `--pasta=`
**3. Validar resultado:** consulte a aba "Controle Fiscal" com saldos por remessa
**4. Informar arquivo gerado** à usuária

## Estrutura Esperada

### ZSD (aba com dados SAP)
- **Remessas** (CFOP 5xxx/6xxx)
- **Retornos** (CFOP 1xxx/2xxx/3xxx)
- **Coluna C**: DOC SAP
- **Coluna E**: Descrição (identifica tipo: "Nota Fiscal saída" vs "Nota Fiscal entrada")
- **Coluna P (col 15)**: Número da nota fiscal
- **Coluna V (col 21)**: Chave eletrônica 44 dígitos
- **Coluna X (col 23)**: Código do material
- **Coluna AB (col 27)**: CFOP
- **Coluna AG (col 32)**: Quantidade
- **Coluna AQ (col 42)**: ICMS por linha

### Razão (aba com dados contábeis)
- **Coluna G (col 6)**: Referência (número da nota ou DOC SAP)
- **Coluna P (col 15)**: Valor
- **Coluna Y (col 24)**: NF RETORNO (a ser preenchida)
- **Coluna Z (col 25)**: DATA RETORNO (a ser preenchida)

### XMLs de Retorno
- Cada item (`<det>`) tem `<qCom>` (quantidade comercial)
- Múltiplos items com mesma nota são **sumados**
- Referências via `<refNFe>` indicam qual remessa foi devolvida
- ICMS extraído de `<vICMS>` em `<ICMSTot>`

## Aba "Controle Fiscal" (Resultado)

Mostra por linha de remessa:

| Coluna | Remessa | Retorno | Saldo |
|--------|---------|---------|-------|
| A-F | Nº Nota, Chave, Valor, Qtde, ICMS, Data | | |
| G-L | | Nº Nota, Chave, Valor, Qtde, ICMS, Data | |
| M-O | | | Saldo Valor, Saldo Qtde, Saldo ICMS |

- **Primeira linha de retorno:** preenche remessa + primeiro retorno + saldo (considerando TODOS os retornos)
- **Linhas adicionais:** apenas retornos extras, sem repetição de remessa

---

## Referência: Critérios de Matching

- **C1 (refNFe)**: Chave de 44 dígitos da remessa no XML — relação direta
- **Múltiplos itens:** Se XML tem 2+ items, quantidades são **sumadas** por remessa
- **Múltiplas linhas ZSD:** Se mesma chave aparece N vezes no ZSD, XML é processado 1 única vez

## ICMS e Quantidades

- Cada linha do ZSD tem seu próprio ICMS (coluna AQ)
- Cada item do XML tem sua própria quantidade (tag `<qCom>`)
- **Não preencher** remessas/retornos com ICMS = R$ 0,00
