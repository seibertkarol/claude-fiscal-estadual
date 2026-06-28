# Regras de Limpeza de XMLs

## Padrão de nome de arquivo

```
{chave44}-nfe.xml           → NF principal (sempre manter, salvo cancelamento)
{chave44}-nfe-{codigo}.xml  → evento da NF
```

A chave tem 44 dígitos. O número da NF fica nas posições [25:34] da chave.

## Regras por código de evento

| Código | Evento | Ação |
|--------|--------|------|
| `110111` | Cancelamento | Apagar TODOS os XMLs desta chave (NF + todos os eventos) |
| `210240` | Operação não Realizada | Apagar TODOS os XMLs desta chave |
| `210210` | Ciência da Operação | Apagar só este evento; manter o `-nfe.xml` |
| `110110` | CC-e (Carta de Correção) | Apagar só este evento; manter o `-nfe.xml` |
| `210200` | (outros de baixo impacto) | Apagar só este evento; manter o `-nfe.xml` |

## Uso do script

```bash
# Modo normal (apaga de verdade)
py -3.14 limpar_xmls.py --pasta="INNFE_20260622154435"

# Modo simulação (só mostra o que faria, não apaga nada)
py -3.14 limpar_xmls.py --pasta="INNFE_20260622154435" --simular
```

## Resultado esperado

O script imprime:
- Total de XMLs encontrados
- Chaves únicas
- Ações tomadas por NF (REMOVE TUDO ou REMOVE EVENTO)
- Quantidade de arquivos apagados
- Quantidade de XMLs restantes (válidos para conciliação)

Após a limpeza, a pasta deve conter apenas arquivos `-nfe.xml` válidos.
