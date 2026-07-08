# Conciliação de Armazenagem — SLC Agrícola

## Resumo do Processo

Este skill executa a conciliação completa de armazenagem, verificando se remessas de grãos tiveram seus valores totalmente devolvidos via notas de retorno (NF-e). O resultado é codificado por cores na Planilha6:

- **Rosa**: Par completamente zerado (remessa + retorno(s) = R$0) com ICMS validado
- **Laranja**: Valor parcial ou incompleto em aberto
- **Azul**: Retorno sem remessa correspondente
- **Amarelo**: Fechamento matemático, mas divergência de ICMS no XML versus lançamento SAP

## Arquivos e Componentes

Os scripts residem em `scripts/` dentro da skill:
- `limpar_xmls.py` — limpa a pasta de XMLs antes da conciliação
- `concilia_pintar.py` — executa conciliação e pinta o Excel

A skill funciona com qualquer fazenda que siga o mesmo layout de colunas, exigindo sempre informação explícita do arquivo Excel e da pasta de XMLs.

## Fluxo de Execução

**1. Identificar inputs:** pergunte qual arquivo Excel e qual pasta de XMLs usar (NUNCA assuma automaticamente)
**2. Limpar XMLs:** execute `limpar_xmls.py` conforme regras (cancelamentos, operações não realizadas, etc.)
**3. Rodar conciliação:** execute `concilia_pintar.py` com os parâmetros `--arquivo=` e `--pasta=`
**4. Validar resultado:** consulte o resumo monetário e as abas "Resumo Conciliacao" e "Relatorio Laranja"
**5. Informar arquivo gerado** à usuária

## Critérios de Conciliação (prioridade decrescente)

- **C1 (refNFe):** chave de 44 dígitos da remessa no XML — relação direta, mais confiável
- **C2 (infAdProd + xPed):** números de NF em campos de produto/pedido
- **C3 (infCpl):** números extraídos de informações complementares, capturando listas completas

C2/C3 só consomem saldo quando zeram exato com o retorno, evitando contaminação de dados parciais.

## Verificação de Integridade

Antes de pintar rosa, compara `vICMS` do XML com o SAP (tolerância: ±R$0,10). Divergências maiores → amarelo obrigatório.

---

## Referência: Estrutura dos XMLs NF-e

### Namespace
`http://www.portalfiscal.inf.br/nfe`

### Campos relevantes

**Identificação da NF:**
- `<nNF>` — número da nota
- `<CNPJ>` (dentro de `<emit>`) — CNPJ do emitente

**Critério C1 — refNFe:**
Localiza-se em: `<det><prod><DI><adi><cAdicional><refNFe>` ou `<NFref><refNFe>`

**Critério C2 — infAdProd e xPed:**
Encontra-se em: `<det><prod>` com campos `<xPed>` e `<infAdProd>`

**Critério C3 — infCpl e xTexto:**
Localiza-se em: `<infAdic>` com `<infCpl>` e `<xTexto>`

### Estrutura do Excel principal (Planilha6)
- Col G (idx 6) — Referência SAP
- Col I (idx 8) — NOTA DE RETORNO
- Col R (idx 17) — Valor: positivo = remessa, negativo = retorno

### CNPJ da empresa (UNITRADING)
`21.425.093/0014-90` — retornos desta empresa são os principais candidatos para conciliação.

---

## Referência: Regras de Limpeza dos XMLs

**Convenção de nomes:** `{chave44}-nfe.xml` (NF principal) e `{chave44}-nfe-{codigo}.xml` (eventos)

**Regras por código de evento:**
- **Códigos 110111 e 210240** (Cancelamento/Operação Não Realizada): apagar TODOS os XMLs da chave (NF + todos os eventos)
- **Códigos 210210, 110110, 210200** (Ciência, CC-e, baixo impacto): remover apenas o arquivo do evento, preservar a NF principal

**Resultado esperado:** pasta deve conter apenas arquivos `-nfe.xml` válidos para conciliação.
