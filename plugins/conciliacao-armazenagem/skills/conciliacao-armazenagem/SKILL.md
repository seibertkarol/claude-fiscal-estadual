---
name: conciliacao-armazenagem
description: >
  Executa o processo completo de conciliação de armazenagem da SLC Agrícola.
  Use esta skill sempre que a usuária disser "fazer a conciliação", "rodar a conciliação",
  "conciliar armazenagem", "processar os XMLs", "pintar a planilha", ou qualquer variação
  de querer executar o processo de conciliação de remessas e retornos de NF-e.
  Funciona para qualquer fazenda/planilha que siga o mesmo layout de colunas — sempre
  perguntar qual é a pasta de XMLs e qual é o arquivo Excel antes de rodar, nunca assumir.
  A skill limpa a pasta de XMLs, roda a conciliação e pinta a Planilha6 com as cores
  corretas (rosa = par zerado, laranja = parcial/incompleto, azul = sem remessa).
---

# Conciliação de Armazenagem — SLC Agrícola

## O que este processo faz

A conciliação de armazenagem verifica se cada **remessa** de grãos teve seu valor
completamente devolvido via **notas de retorno** (NF-e). O resultado é pintado na
Planilha6 do Excel principal:

- **ROSA** — par (ou grupo) completamente zerado: remessa + retorno(s) somam R$0
  E o ICMS do XML bate com o valor lançado no SAP (ver verificação de ICMS abaixo)
- **LARANJA** — parcial ou incompleto: algum valor ainda está em aberto
- **AZUL** — retorno sem remessa correspondente na planilha
- **AMARELO** — o retorno e suas remessas até fechariam matematicamente, mas o
  **ICMS declarado no XML não bate com o valor lançado no SAP** (erro de
  lançamento, geralmente corrigido depois por uma linha de ajuste separada).
  Nunca pinta de rosa, mesmo que a soma com o ajuste feche — precisa de revisão
  manual.

A garantia matemática: filtrar a Planilha6 pela cor ROSA deve resultar em soma = R$0.

Além da cor, o script preenche duas colunas auxiliares em toda linha pintada
(remessas e retornos):
- **Coluna I — NOTA DE RETORNO**: número(s) da(s) nota(s) de retorno que casaram
  com aquela linha (separados por vírgula quando há mais de um retorno parcial)
- **Coluna J — DATA DO RETORNO** (inserida pelo script, empurra as colunas
  originais uma posição para a direita): data de lançamento da nota de retorno

E duas abas novas são criadas/recriadas a cada rodagem:
- **"Resumo Conciliacao"**: contagem de linhas rosa/laranja/azul/amarelo
  (remessas e retornos separados) e o diagnóstico monetário — igual ao que o
  VBA antigo fazia.
- **"Relatorio Laranja"**: lista TODOS os casos que não são rosa (laranja,
  azul, amarelo, sem retorno no SAP), ordenados pela diferença monetária (do
  maior problema para o menor), com colunas extras `vICMS no XML` e
  `Diferença ICMS` para facilitar a revisão manual. Linhas amarelas fortes =
  divergência de ICMS; vermelho claro = outras diferenças grandes (>R$1.000).

### Verificação de integridade do ICMS

Antes de pintar rosa, o script compara o `vICMS` declarado no XML de cada
retorno com o valor lançado no SAP/Planilha6 para aquele retorno (tolerância:
**até R$0,10 de diferença é considerado igual**, por causa de arredondamento).
Se divergir além disso, o retorno (e as remessas que ele referencia) NUNCA
ficam rosa — ficam amarelos, mesmo que a soma com uma linha de ajuste feche
matematicamente. Isso existe porque já aconteceram casos de notas de retorno
lançadas no SAP com o valor errado, corrigidas depois por um ajuste manual —
e sem essa verificação, o script pintaria de rosa como se nada tivesse
acontecido, escondendo o erro de lançamento original.

## Funciona para qualquer fazenda

Este processo **não é exclusivo de uma planilha**. Várias fazendas/colegas usam o
mesmo `concilia_pintar.py`, desde que a planilha de cada uma siga o mesmo layout de
colunas (Planilha6, zsd, mesmas posições de Referência/Data de lançamento/Valor/etc).
**Nunca assuma qual arquivo ou pasta usar** — sempre pergunte os dois antes de rodar,
mesmo que pareça óbvio pelo contexto da conversa.

## Arquivos envolvidos

Os scripts ficam **dentro desta skill**, em `scripts/` — funcionam em qualquer
máquina, não dependem de nenhum caminho fixo de usuário. O Excel e a pasta de XML
variam a cada rodagem (são passados como parâmetro) e normalmente estão na pasta
de trabalho da usuária (ex: Downloads), não dentro da skill:

| Arquivo | Função |
|---------|--------|
| `scripts/limpar_xmls.py` | Limpa a pasta de XMLs antes da conciliação |
| `scripts/concilia_pintar.py` | Executa a conciliação e pinta o Excel (recebe `--arquivo` e `--pasta`) |
| `EXPORT_*.xlsx` (varia, fica na pasta da usuária) | Planilha principal da fazenda (Planilha6, zsd, razao) |
| `INNFE_*/` (varia, fica na pasta da usuária) | Pasta com XMLs das NF-e de retorno do lote a processar |

Ao rodar os comandos abaixo, use o caminho completo dos scripts dentro da pasta
da skill (algo como `.../skills/conciliacao-armazenagem/scripts/concilia_pintar.py`)
e rode a partir da pasta onde estão o Excel e a pasta de XML da usuária.

## Passo a passo do processo

### 1. Identificar os inputs

**Sempre pergunte, mesmo que ache que sabe:**
- Qual é o **arquivo Excel** desta fazenda? (ex: `EXPORT_20260622_091056.xlsx`)
- Qual é a **pasta de XMLs** deste lote? (ex: `INNFE_20260622154435`)

Não tente adivinhar pela data de modificação mais recente nem reaproveitar valores
de uma conversa antiga — cada rodagem pode ser de uma fazenda/lote diferente.

### 2. Limpar os XMLs

```bash
cd PASTA_DE_TRABALHO_DA_USUARIA
py -3.14 "CAMINHO_DA_SKILL/scripts/limpar_xmls.py" --pasta="NOME_DA_PASTA_XML"
```

**Regras de limpeza** (ver `references/regras_limpeza.md`):
- `110111` Cancelamento → apaga TODOS os XMLs da nota
- `210240` Operação não Realizada → apaga TODOS os XMLs da nota
- `210210` Ciência da Operação → apaga só o evento, mantém `-nfe.xml`
- `110110` CC-e → apaga só o evento
- `210200` → apaga só o evento

### 3. Rodar a conciliação

```bash
cd PASTA_DE_TRABALHO_DA_USUARIA
py -3.14 "CAMINHO_DA_SKILL/scripts/concilia_pintar.py" --arquivo="NOME_DO_EXCEL.xlsx" --pasta="NOME_DA_PASTA_XML"
```

O script vai:
1. Carregar Planilha6 e ZSD do Excel informado
2. Processar cada XML da pasta informada, em ordem crescente de NF
3. Tentar casar retornos com remessas por 3 critérios (C1 > C2 > C3)
4. Aplicar algoritmo de ponto fixo para determinar cores
5. Inserir a coluna J (DATA DO RETORNO) e preencher I/J em toda linha pintada
6. Criar a aba "Resumo Conciliacao" com a contagem da rodagem
7. Salvar automaticamente em um nome novo (`..._CONCILIADO.xlsx`, ou `_v2`, `_v3`...
   se já existir — nunca sobrescreve nem trava por arquivo aberto no Excel)

### 4. Verificar o resultado

Após rodar, leia o output do console e reporte à usuária:

```
RESUMO DOS RETORNOS:
  Zerou   (rosa):    X     ← retornos que bateram exato
  Parcial (laranja): X     ← retornos com diferença
  Sem remessa (azul): X    ← retornos sem match na planilha
  Sem retorno P6:    X     ← XMLs sem lançamento na planilha

DIAGNOSTICO MONETARIO:
  ROSA saldo liquido: R$X  ← deve ser ~R$0
```

Se o saldo rosa estiver longe de R$0, investigue (ver seção Troubleshooting).

### 5. Informar o arquivo de saída

Diga à usuária o nome exato do arquivo gerado (aparece no console como
"Arquivo salvo: ...") e que ela pode abrir, checar a aba "Resumo Conciliacao"
e filtrar a Planilha6 pela cor rosa para verificar.

---

## Critérios de conciliação (C1 > C2 > C3)

**C1 — refNFe** (mais confiável): o XML do retorno contém a chave de 44 dígitos
da remessa dentro da tag `<refNFe>`. Relação eletrônica direta.

**C2 — infAdProd + xPed**: números de NF extraídos dos campos de produto
(`infAdProd`) ou pedido (`xPed`). Texto estruturado por item.

**C3 — infCpl + xTexto**: números extraídos do campo de informações complementares
ou texto da NF. Pode referenciar VÁRIAS notas na mesma frase (ex: "NOTAS FISCAIS
66377, 66384 E 66389 | NOTAS FISCAIS 66725, 66726, 66727, 66728") — a extração
captura listas completas separadas por vírgula/"E", não só o primeiro número.

**Quando C2/C3 consome saldo (fica elegível a ROSA):** C1 sempre consome saldo,
mesmo em match PARCIAL. C2/C3 só consomem saldo (e só ficam elegíveis a ROSA)
quando o valor encontrado **zera exato** com o retorno — isso evita que um match
de texto parcial/ambíguo contamine o saldo de uma remessa indevidamente. Se a
soma das remessas via C2/C3 não fechar exato com o retorno, ele fica PARCIAL e
não consome saldo nenhum (apenas identifica quais remessas foram referenciadas,
para preencher a coluna I).

---

## Regra das cores (ponto fixo)

Um grupo de retornos + remessas fica **ROSA** quando:
- Todos os retornos do grupo são ZEROU (diferença ≤ R$0,05)
- Todas as remessas do grupo têm saldo = 0 (todo valor foi devolvido)
- O grupo é fechado: as remessas não são consumidas por retornos de fora do grupo

Exemplo correto: NF 722 e NF 756 compartilham a remessa 68582 (R$69.430).
Cada uma devolveu metade. Juntas zeram → todas as três ficam ROSA.

---

## Troubleshooting

**Rosa filter não soma R$0:**
Verifique o diagnóstico monetário no console. Se estiver desequilibrado, pode ser:
- Retornos C3 com falso positivo (diferença enorme → esperar é normal, são de períodos anteriores)
- Remessas referenciadas por um PARCIAL e um ZEROU (ficam laranja corretamente)

**SEM_RETORNO_P6:**
O XML existe mas a NF não está na Planilha6. Pode ser conta ou período diferente.
Liste as NFs e informe à usuária para verificação manual.

**"ERRO: informe o arquivo Excel..." ou "...a pasta de XMLs...":**
O script agora exige `--arquivo=` e `--pasta=` explicitamente (não tem mais valor
fixo no código). Pergunte à usuária os dois nomes e rode de novo com eles.

**Colunas parecem desalinhadas em uma planilha nova:**
O script pressupõe que a planilha desta fazenda segue o MESMO layout de colunas
da planilha original (mesma posição de Referência, Data de lançamento, Valor,
etc. na Planilha6 e na zsd). Se uma fazenda usar um layout diferente, isso
precisa ser ajustado no código antes de rodar — avise a usuária.

---

## Referências detalhadas

- `references/regras_limpeza.md` — regras completas de limpeza de XMLs
- `references/estrutura_xml.md` — estrutura das NF-e e campos relevantes
