# Estrutura dos XMLs NF-e

## Namespace

```
http://www.portalfiscal.inf.br/nfe
```

Todos os elementos XML usam este namespace. Exemplo de busca em Python:
```python
NS = 'http://www.portalfiscal.inf.br/nfe'
root.find('.//{%s}nNF' % NS)
```

## Campos relevantes

### Identificação da NF
- `<nNF>` — número da NF (ex: `513`)
- `<CNPJ>` (dentro de `<emit>`) — CNPJ do emitente

### Chave de acesso (44 dígitos)
Formato: `CCYYYYMMCNPJNNNNNNNNMMMMDDDDDDDDDD`
- Posições [25:34] → número da NF (9 dígitos com zeros à esquerda)

### Critério C1 — refNFe
```xml
<det>
  <prod>
    <DI>
      <adi>
        <cAdicional>
          <refNFe>44digitkey</refNFe>
```
Ou diretamente em `<NFref>`:
```xml
<NFref>
  <refNFe>44digitkey</refNFe>
</NFref>
```

### Critério C2 — infAdProd e xPed
```xml
<det>
  <prod>
    <xPed>000012345</xPed>
    <infAdProd>NF 12345/001</infAdProd>
```

### Critério C3 — infCpl e xTexto
```xml
<infAdic>
  <infCpl>NOTAS FISCAIS: 12345, 67890</infCpl>
  <xTexto>NF 12345-001</xTexto>
```

## Estrutura do Excel principal

### Planilha6 (ledger completo)
- Col G (idx 6) — Referência: número SAP ou formato `000060203-001`
- Col I (idx 8) — NOTA DE RETORNO: preenchida pelo script com NF(s) do retorno
- Col R (idx 17) — Valor: positivo = remessa, negativo = retorno
- Linha 1 — Fórmula SUBTOTAL: `=SUBTOTAL(9,R3:R9030)`

### ZSD (tabela SAP → NF-e)
- Col E (idx 4) — Nº documento SAP
- Col AY (idx 50) — Número NF-e
- Col O (idx 14) — Chave de acesso 44 dígitos
- Col L (idx 11) — CNPJ

### razao
- Col G (idx 6) — Referência SAP
- Col T (idx 19) — Valor

## CNPJ da empresa (UNITRADING)
`21.425.093/0014-90`

Retornos desta empresa são os candidatos principais para conciliação.
