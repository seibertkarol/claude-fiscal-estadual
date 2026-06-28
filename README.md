# Marketplace de Plugins — Time Fiscal Estadual SLC Agrícola

Marketplace de plugins do Claude Code para o time fiscal estadual.

## Plugins disponíveis

### conciliacao-armazenagem

Concilia remessas e notas de retorno de NF-e automaticamente, pintando a
Planilha6 com base no status de cada par (zerado, parcial, sem remessa).

## Como instalar (colegas)

No Claude Code, execute:

```
/plugin marketplace add seibertkarol/Claude-Fiscal-Estadual-Armazenagem
/plugin install conciliacao-armazenagem@slc-fiscal-marketplace
```

Depois, para usar a skill, basta pedir: *"faça a conciliação da armazenagem"*.

## Como atualizar (após mudanças)

1. Edite os arquivos dentro de `plugins/conciliacao-armazenagem/`
2. Suba as mudanças para este repositório (`git push`)
3. Se a versão em `plugin.json` for incrementada, os colegas com
   auto-update habilitado recebem a atualização automaticamente
