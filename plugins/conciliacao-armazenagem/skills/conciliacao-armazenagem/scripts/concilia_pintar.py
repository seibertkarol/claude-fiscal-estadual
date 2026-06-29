import pandas as pd
import xml.etree.ElementTree as ET
import os, re, sys
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

NS = 'http://www.portalfiscal.inf.br/nfe'

# ---------------------------------------------------------------
# PARAMETROS — cada fazenda/colega informa o proprio arquivo e pasta de XML
# Uso: py -3.14 concilia_pintar.py --arquivo="EXPORT_XXXX.xlsx" --pasta="INNFE_XXXX"
# Pressupoe que TODAS as planilhas seguem o mesmo layout de colunas (Planilha6,
# zsd, mesma posicao das colunas Referencia/Data lancamento/Valor/etc).
# ---------------------------------------------------------------
ARQUIVO = None
XML_DIR = None
for arg in sys.argv[1:]:
    if arg.startswith('--arquivo='):
        ARQUIVO = arg.split('=', 1)[-1].strip().strip('"')
    elif arg.startswith('--pasta='):
        XML_DIR = arg.split('=', 1)[-1].strip().strip('"')

if not ARQUIVO:
    print("ERRO: informe o arquivo Excel com --arquivo=\"NOME.xlsx\"")
    sys.exit(1)
if not XML_DIR:
    print("ERRO: informe a pasta de XMLs com --pasta=\"NOME_DA_PASTA\"")
    sys.exit(1)

FILL_ROSA    = PatternFill('solid', fgColor='F2CEEF')   # ZEROU total
FILL_LARANJA = PatternFill('solid', fgColor='E97132')   # PARCIAL
FILL_AZUL    = PatternFill('solid', fgColor='467886')   # SEM REMESSA
FONT_BRANCO  = Font(color='FFFFFF')

# ---------------------------------------------------------------
# DETECCAO DE ESTADO: o arquivo pode ser "virgem" (primeira rodagem) ou ja
# ter sido processado antes (tem a linha TOTAL FILTRADO e a coluna DATA DO
# RETORNO inseridas por uma rodagem anterior nesta MESMA planilha evolutiva).
# Isso evita duplicar linha/coluna e desalinhar tudo a cada nova rodagem.
# ---------------------------------------------------------------
from openpyxl import load_workbook as _peek_wb
_wb_peek = _peek_wb(ARQUIVO, read_only=True, data_only=True)
_ws_peek = _wb_peek['Planilha6']
JA_PROCESSADO = (_ws_peek['A1'].value == 'TOTAL FILTRADO')
_wb_peek.close()

PANDAS_HEADER = 1 if JA_PROCESSADO else 0   # linha do cabecalho real (0-based)
COL_OFFSET    = 1 if JA_PROCESSADO else 0   # deslocamento das colunas J em diante

COL_REFERENCIA   = 6                  # nunca desloca (fica antes da coluna I)
COL_DATA_LANC    = 10 + COL_OFFSET    # K se ja processado, J na primeira vez
COL_VALOR        = 17 + COL_OFFSET    # S se ja processado, R na primeira vez

print(f"Arquivo {'JA PROCESSADO antes' if JA_PROCESSADO else 'NOVO (primeira rodagem)'} "
      f"— ajustando leitura automaticamente.")

print("Carregando dados...")
df_p6  = pd.read_excel(ARQUIVO, sheet_name='Planilha6', header=PANDAS_HEADER)
df_zsd = pd.read_excel(ARQUIVO, sheet_name='zsd',       header=0)

df_p6['_valor'] = pd.to_numeric(df_p6.iloc[:, COL_VALOR], errors='coerce')

COL_NOTA_RET = 8  # coluna I (0-based) — nunca desloca
print(f"Coluna NOTA DE RETORNO (I): '{df_p6.columns[COL_NOTA_RET]}'")

# ---------------------------------------------------------------
# MAPA 1: ZSD doc_sap -> nfe
# ---------------------------------------------------------------
zsd_by_doc = {}
for _, row in df_zsd.iterrows():
    doc = str(int(row.iloc[4])) if pd.notna(row.iloc[4]) else None
    if doc and doc not in zsd_by_doc:
        zsd_by_doc[doc] = {'nfe': row.iloc[50]}

# ---------------------------------------------------------------
# MAPA 2: P6 retorno nfe -> indices + valor
# ---------------------------------------------------------------
def ref_to_str(v):
    try:    return str(int(float(v)))
    except: return None

p6_retorno_by_nfe = {}
p6_retorno_valor  = {}
for idx, row in df_p6.iterrows():
    val = row['_valor']
    if pd.isna(val) or val >= 0: continue
    ref = ref_to_str(row.iloc[6])
    if not ref: continue
    info = zsd_by_doc.get(ref)
    if not info: continue
    try: nfe = int(info['nfe'])
    except: continue
    p6_retorno_by_nfe.setdefault(nfe, []).append(idx)
    p6_retorno_valor[nfe] = p6_retorno_valor.get(nfe, 0.0) + float(val)

print(f"Retornos identificados via ZSD: {len(p6_retorno_by_nfe)}")

# Data de lancamento de cada retorno (coluna K = idx 10), para preencher
# na coluna J (DATA DO RETORNO) de toda linha que ele referencia
nnf_data_lancamento = {}
for nfe, idxs in p6_retorno_by_nfe.items():
    for idx in idxs:
        data_lanc = df_p6.iloc[idx, COL_DATA_LANC]
        if pd.notna(data_lanc):
            nnf_data_lancamento[nfe] = data_lanc
            break

# ---------------------------------------------------------------
# MAPA 3: P6 remessa nf_num -> indices + valor original
# ---------------------------------------------------------------
def extrair_num(v):
    s = str(v).strip()
    if s.lower() in ('nan','none',''): return None
    if '-' in s: s = s.split('-')[0]
    s = s.lstrip('0')
    try:    return int(s)
    except: return None

p6_remessa_by_num = {}   # nf_num -> [pandas_idx]
p6_remessa_valor  = {}   # nf_num -> valor total original
for idx, row in df_p6.iterrows():
    val = row['_valor']
    if pd.isna(val) or val <= 0: continue
    n = extrair_num(str(row.iloc[6]))
    if n is not None:
        p6_remessa_by_num.setdefault(n, []).append(idx)
        p6_remessa_valor[n] = p6_remessa_valor.get(n, 0.0) + float(val)

# Saldo disponivel de cada remessa (vai sendo consumido conforme retornos processam)
saldo_remessa = dict(p6_remessa_valor)

# Rastreia quais retornos C1 consumiram saldo de cada remessa (por nf_num)
# Somente esses definem se a remessa fica rosa ou laranja
from collections import defaultdict
saldo_consumers = defaultdict(list)   # nf_num -> [nnf_retorno, ...]

# ---------------------------------------------------------------
# EXTRACAO DE NUMEROS DE TEXTO (C2/C3)
# ---------------------------------------------------------------
EXPORT_CTX = ['NFE DE EXPORTACAO','NF DE EXPORTACAO','GUIA DE SAIDA',
              'NAVIO','CONTAINER','BOOKING','CROSSDOCKING','DESTINO:','RES:']
PREFIXOS   = ['NOTAS FISCAIS','NOTA FISCAL','NOTAS','NOTA',
              'PESOS:','PESO:','NFE','NF','N:','N ']

def extrair_nums_texto(txt):
    """
    Extrai numeros de NF referenciados em texto livre (infCpl/xTexto).
    Captura LISTAS de numeros separados por virgula e/ou "E"
    (ex: "NOTAS FISCAIS 66377, 66384 E 66389" -> 66377, 66384, 66389),
    nao apenas o primeiro numero apos o prefixo.
    """
    if not txt: return set()
    nums, tu = set(), txt.upper()
    for m in re.finditer(r'(\d{5,9})-00[123]', txt):
        nums.add(int(m.group(1)))
    for pref in PREFIXOS:
        pos = 0
        while True:
            i2 = tu.find(pref, pos)
            if i2 == -1: break
            sub = txt[i2+len(pref):i2+len(pref)+150]
            # Casa uma sequencia: primeiro numero, depois repeticoes de
            # (virgula/"E"/"e") + numero, parando no primeiro token que nao bate
            m = re.match(
                r'\s*[:\-]?\s*(\d{5,9}(?:\s*(?:,|/|\bE\b)\s*\d{5,9})*)',
                sub, re.IGNORECASE
            )
            if m:
                ctx = tu[max(0, i2-60):i2]
                if not any(c in ctx for c in EXPORT_CTX):
                    for numstr in re.findall(r'\d{5,9}', m.group(1)):
                        nums.add(int(numstr))
            pos = i2 + len(pref)
    return nums

def buscar_remessas(refs_nums, val_retorno_abs):
    """
    Busca remessas com saldo disponivel.
    Aloca apenas o minimo entre saldo e o que o retorno ainda precisa.
    Retorna (indices_p6, soma_alocada, refs_encontradas).
    """
    idxs, soma, vistos, refs_found = [], 0.0, set(), set()
    restante = val_retorno_abs
    for n in refs_nums:
        saldo = saldo_remessa.get(n, 0.0)
        if saldo <= 0: continue
        for idx in p6_remessa_by_num.get(n, []):
            if idx in vistos: continue
            v = df_p6.at[idx, '_valor']
            if pd.notna(v) and v > 0:
                alocado = min(saldo, restante)
                idxs.append(idx)
                soma += alocado
                restante -= alocado
                vistos.add(idx)
                refs_found.add(n)
    return idxs, soma, refs_found

def consumir_saldo(refs_nums, val_retorno_abs, nnf_retorno):
    """Desconta o saldo e registra quem consumiu cada remessa."""
    restante = val_retorno_abs
    for n in refs_nums:
        saldo = saldo_remessa.get(n, 0.0)
        if saldo <= 0: continue
        usado = min(saldo, restante)
        saldo_remessa[n] = max(0.0, saldo - usado)
        if nnf_retorno not in saldo_consumers[n]:
            saldo_consumers[n].append(nnf_retorno)
        restante -= usado
        if restante <= 0: break

# ---------------------------------------------------------------
# PRE-PROCESSAMENTO: le todos os XMLs
# ---------------------------------------------------------------
refs_por_retorno = {}
xml_roots = {}
for fname in sorted(os.listdir(XML_DIR)):
    if not fname.endswith('.xml'): continue
    root = ET.parse(os.path.join(XML_DIR, fname)).getroot()
    nnf_el = root.find('.//{%s}nNF' % NS)
    if nnf_el is None: continue
    nnf = int(nnf_el.text)
    refs = set()
    for el in root.findall('.//{%s}refNFe' % NS):
        chave = (el.text or '').strip()
        if len(chave) == 44:
            try: refs.add(int(chave[25:34]))
            except: pass
    refs_por_retorno[nnf] = refs
    xml_roots[nnf] = root

# Detecta remessas compartilhadas
retornos_por_remessa = {}
for nnf, refs in refs_por_retorno.items():
    for r in refs:
        retornos_por_remessa.setdefault(r, set()).add(nnf)
compartilhadas = {r: nfs for r, nfs in retornos_por_remessa.items() if len(nfs) > 1}
if compartilhadas:
    print(f"Remessas compartilhadas (retorno parcial): {len(compartilhadas)}")

# ---------------------------------------------------------------
# PASSO 1: CONCILIACAO — processa em ordem crescente de NF
# determina status de cada retorno e quais remessas ele consumiu
# ---------------------------------------------------------------
print("\nPasso 1: Processando XMLs (calculando status)...")

xmls_ordenados = sorted(xml_roots.keys())

status_retorno   = {}  # nnf -> 'ZEROU' | 'PARCIAL' | 'SEM_REMESSA' | 'SEM_RAZAO'
ret_rem_idxs     = {}  # nnf -> [pandas_idx das remessas consumidas]
ret_refs_usadas  = {}  # nnf -> set de nf_nums de remessas consumidas
nota_ret_por_rem = {}  # pandas_idx_remessa -> [nnf_retorno, ...]
log_result = []
cnt = {'ZEROU': 0, 'PARCIAL': 0, 'SEM_REMESSA': 0, 'SEM_RAZAO': 0}

for nnf in xmls_ordenados:
    root = xml_roots[nnf]
    ret_idxs = p6_retorno_by_nfe.get(nnf, [])
    val_ret  = p6_retorno_valor.get(nnf, 0.0)
    val_ret_abs = abs(val_ret)

    if not ret_idxs:
        status_retorno[nnf] = 'SEM_RAZAO'
        cnt['SEM_RAZAO'] += 1
        log_result.append(f'[?] NF {nnf:>6} | SEM_RETORNO_P6')
        continue

    # C1: refNFe
    refs_c1  = refs_por_retorno.get(nnf, set())
    rem_idxs, val_rem, refs_found = [], 0.0, set()
    criterio = ''

    if refs_c1:
        rem_idxs, val_rem, refs_found = buscar_remessas(refs_c1, val_ret_abs)
        if rem_idxs: criterio = 'C1:refNFe'

    # C2: infAdProd + xPed
    if not rem_idxs:
        refs_c2 = set()
        for tag in ['infAdProd','xPed']:
            for el in root.findall('.//{%s}%s' % (NS, tag)):
                refs_c2 |= extrair_nums_texto(el.text or '')
        if refs_c2:
            rem_idxs, val_rem, refs_found = buscar_remessas(refs_c2, val_ret_abs)
            if rem_idxs: criterio = 'C2:infAdProd+xPed'

    # C3: infCpl + xTexto
    if not rem_idxs:
        refs_c3 = set()
        for tag in ['infCpl','xTexto']:
            for el in root.findall('.//{%s}%s' % (NS, tag)):
                refs_c3 |= extrair_nums_texto(el.text or '')
        if refs_c3:
            rem_idxs, val_rem, refs_found = buscar_remessas(refs_c3, val_ret_abs)
            if rem_idxs: criterio = 'C3:infCpl+xTexto'

    # Classifica status
    if not rem_idxs:
        status = 'SEM_REMESSA'; cnt['SEM_REMESSA'] += 1
    else:
        dif = val_ret + val_rem
        if abs(dif) <= 0.05:
            status = 'ZEROU';   cnt['ZEROU'] += 1
        else:
            status = 'PARCIAL'; cnt['PARCIAL'] += 1

    status_retorno[nnf] = status

    # Consome saldo (e registra como consumidor, elegivel a rosa) quando:
    #   - e C1 (chave eletronica, sempre confiavel, mesmo se PARCIAL)
    #   - OU e C2/C3 (texto) MAS zerou exato — so confiamos em C2/C3 quando
    #     o valor bate 100%, para nao contaminar com falso positivo parcial
    if rem_idxs and (criterio == 'C1:refNFe' or status == 'ZEROU'):
        consumir_saldo(refs_found, val_ret_abs, nnf)

    # Registra quais remessas este retorno consumiu
    ret_rem_idxs[nnf]   = rem_idxs
    ret_refs_usadas[nnf] = refs_found

    # Mapeia: remessa_idx -> quais retornos a usaram
    for idx in rem_idxs:
        nota_ret_por_rem.setdefault(idx, [])
        if nnf not in nota_ret_por_rem[idx]:
            nota_ret_por_rem[idx].append(nnf)

    dif_txt = f'dif=R${val_ret+val_rem:,.2f}' if rem_idxs else ''
    ic = '[OK]' if status=='ZEROU' else ('[~]' if status=='PARCIAL' else '[X]')
    log_result.append(
        f'{ic} NF {nnf:>6} | {status:<12} | {criterio:<22} '
        f'| ret={len(ret_idxs)} rem={len(rem_idxs)} '
        f'| R${val_ret:,.2f} + R${val_rem:,.2f} {dif_txt}'
    )

# ---------------------------------------------------------------
# PASSO 2+3: DETERMINA COR — ALGORITMO ITERATIVO (ponto fixo)
#
# Regra do ROSA:
#   Um grupo de retornos + remessas fica ROSA quando forma um ciclo
#   completamente fechado: todos os retornos do grupo sao ZEROU,
#   todas as remessas do grupo tem saldo = 0.
#   Ex: NF 722 + NF 756 compartilham remessa 68582. Juntos zeram -> todos ROSA.
#
# Garantia matematica: filtro rosa = soma R$0.
# O loop converge porque cores so podem ser perdidas (nunca ganhas).
# ---------------------------------------------------------------
print("Passo 2/3: Algoritmo iterativo de cores (ponto fixo)...")

FILL_MAP = {'rosa': FILL_ROSA, 'laranja': FILL_LARANJA, 'azul': FILL_AZUL}
pintura_p6 = {}   # pandas_idx -> (cor, nnf, criterio)

# Estado inicial de cor dos retornos (pelo status da conciliacao)
cor_retorno = {}   # nnf -> 'rosa'|'laranja'|'azul'
for nnf in xmls_ordenados:
    st = status_retorno.get(nnf)
    if st in ('SEM_RAZAO', None): continue
    cor_retorno[nnf] = {'ZEROU':'rosa', 'PARCIAL':'laranja', 'SEM_REMESSA':'azul'}.get(st, 'laranja')

# Registra retornos nas linhas de retorno da Planilha6
for nnf in xmls_ordenados:
    if nnf not in cor_retorno: continue
    for idx in p6_retorno_by_nfe.get(nnf, []):
        nota_ret_por_rem.setdefault(idx, [])
        if nnf not in nota_ret_por_rem[idx]:
            nota_ret_por_rem[idx].append(nnf)

MAX_ITER = 20
for iteracao in range(1, MAX_ITER + 1):
    mudancas = 0

    # --- Determina cor das REMESSAS ---
    cor_remessa = {}   # pandas_idx -> 'rosa'|'laranja'
    for idx, nfs_list in nota_ret_por_rem.items():
        val = df_p6.at[idx, '_valor']
        if pd.isna(val) or val <= 0:
            continue  # linha de retorno

        nf_num = extrair_num(str(df_p6.iloc[idx, 6]))
        retornos_que_usaram = [n for n in nfs_list if n in ret_rem_idxs and idx in ret_rem_idxs[n]]
        if not retornos_que_usaram:
            continue

        saldo_final = saldo_remessa.get(nf_num, p6_remessa_valor.get(nf_num, 1.0))
        saldo_zerado = saldo_final <= 0.05

        # Todos os consumidores C1 precisam ser atualmente rosa
        consumers_c1 = saldo_consumers.get(nf_num, [])
        todos_rosa = bool(consumers_c1) and all(cor_retorno.get(n) == 'rosa' for n in consumers_c1)

        cor_remessa[idx] = 'rosa' if (saldo_zerado and todos_rosa) else 'laranja'

    # --- Verifica e ajusta cor dos RETORNOS ZEROU ---
    for nnf in xmls_ordenados:
        if cor_retorno.get(nnf) != 'rosa': continue
        rem_idxs_nf = ret_rem_idxs.get(nnf, [])
        alguma_nao_rosa = any(
            cor_remessa.get(idx, 'laranja') != 'rosa'
            for idx in rem_idxs_nf
            if pd.notna(df_p6.at[idx, '_valor']) and df_p6.at[idx, '_valor'] > 0
        )
        if alguma_nao_rosa:
            cor_retorno[nnf] = 'laranja'
            mudancas += 1

    print(f"  iteracao {iteracao}: {mudancas} retornos ajustados")
    if mudancas == 0:
        print(f"  Ponto fixo atingido em {iteracao} iteracao(oes).")
        break

# Constroi pintura_p6 final
for nnf, cor in cor_retorno.items():
    for idx in p6_retorno_by_nfe.get(nnf, []):
        pintura_p6[idx] = (cor, nnf, '')

for idx, cor in cor_remessa.items():
    nfs_list = nota_ret_por_rem.get(idx, [])
    retornos_que_usaram = [n for n in nfs_list if n in ret_rem_idxs and idx in ret_rem_idxs[n]]
    ref = retornos_que_usaram[0] if retornos_que_usaram else 0
    pintura_p6[idx] = (cor, ref, '')

for idx, cor in cor_remessa.items():
    nfs_list = nota_ret_por_rem.get(idx, [])
    retornos_que_usaram = [n for n in nfs_list if n in ret_rem_idxs and idx in ret_rem_idxs[n]]
    ref = retornos_que_usaram[0] if retornos_que_usaram else 0
    pintura_p6[idx] = (cor, ref, '')

# ---------------------------------------------------------------
# APLICA PINTURA NA PLANILHA6
# ---------------------------------------------------------------
print(f"Aplicando cores em {len(pintura_p6)} linhas da Planilha6...")

wb = load_workbook(ARQUIVO)
ws = wb['Planilha6']

# Insere a coluna J (DATA DO RETORNO) SOMENTE na primeira rodagem. Se o
# arquivo ja foi processado antes, a coluna ja existe — nao inserir de novo
# (evita duplicar colunas a cada rodagem na mesma planilha evolutiva).
if not JA_PROCESSADO:
    ws.insert_cols(10)
    ws.cell(row=1, column=10, value='DATA DO RETORNO')

# O resultado final SEMPRE tem 1 linha TOTAL FILTRADO no topo (inserida agora
# ou ja existente de antes), entao a linha excel final = pandas_idx + 3, sempre.
EXCEL_ROW_BASE = 3

for pandas_idx, (cor, nnf, criterio) in pintura_p6.items():
    excel_row = pandas_idx + EXCEL_ROW_BASE
    fill = FILL_MAP[cor]
    for cell in ws[excel_row]:
        cell.fill = fill
        cell.font = FONT_BRANCO if cor == 'azul' else Font()

# Preenche coluna I com NF(s) de retorno (remessas e retornos)
COL_I_EXCEL = 9
print(f"Preenchendo NOTA DE RETORNO em {len(nota_ret_por_rem)} linhas...")
COL_J_EXCEL = 10

def fmt_data(d):
    try: return d.strftime('%d/%m/%Y')
    except: return str(d)

for pandas_idx, nfs_list in nota_ret_por_rem.items():
    if pandas_idx not in pintura_p6: continue
    excel_row = pandas_idx + EXCEL_ROW_BASE
    valor_col_i = ', '.join(str(n) for n in sorted(set(nfs_list)))
    ws.cell(row=excel_row, column=COL_I_EXCEL, value=valor_col_i)

    # Coluna J (DATA DO RETORNO): data de lancamento da(s) nota(s) de retorno
    datas_distintas = []
    for n in sorted(set(nfs_list)):
        d = nnf_data_lancamento.get(n)
        if d is not None and d not in datas_distintas:
            datas_distintas.append(d)

    if len(datas_distintas) == 1:
        # Uma so data -> escreve como data real (mantem formatacao de data no Excel)
        cell_j = ws.cell(row=excel_row, column=COL_J_EXCEL, value=datas_distintas[0])
        cell_j.number_format = 'DD/MM/YYYY'
    elif len(datas_distintas) > 1:
        # Datas diferentes (retornos parciais em datas distintas) -> texto separado por virgula
        ws.cell(row=excel_row, column=COL_J_EXCEL, value=', '.join(fmt_data(d) for d in datas_distintas))

# Linha de SUBTOTAL no topo — insere SOMENTE na primeira rodagem. Se ja
# existir (rodagem repetida na mesma planilha evolutiva), so atualiza a
# formula/estilo da linha 1 existente, sem inserir outra linha por cima.
if not JA_PROCESSADO:
    ws.insert_rows(1)
ws['A1'] = 'TOTAL FILTRADO'

from openpyxl.utils import get_column_letter
col_valor_letra = get_column_letter(COL_VALOR + 1 + (0 if JA_PROCESSADO else 1))
ultima_linha = ws.max_row
formula_cell = f'{col_valor_letra}1'
ws[formula_cell] = f'=SUBTOTAL(9,{col_valor_letra}3:{col_valor_letra}{ultima_linha})'
ws[formula_cell].number_format = '#,##0.00'
fill_total = PatternFill('solid', fgColor='1F4E79')
font_branco_bold = Font(bold=True, color='FFFFFF')
for col in range(1, ws.max_column + 1):
    ws.cell(row=1, column=col).fill = fill_total
    ws.cell(row=1, column=col).font = font_branco_bold
ws[formula_cell].number_format = '#,##0.00'

# Conta remessas/retornos por cor (precisa ser antes do save, para escrever na aba Resumo)
cnt_rem = {'rosa': 0, 'laranja': 0, 'azul': 0}
cnt_ret = {'rosa': 0, 'laranja': 0, 'azul': 0}
for pandas_idx, (cor, nnf, _) in pintura_p6.items():
    val = df_p6.at[pandas_idx, '_valor']
    if pd.notna(val):
        if val > 0: cnt_rem[cor] = cnt_rem.get(cor,0) + 1
        else:       cnt_ret[cor] = cnt_ret.get(cor,0) + 1

# Soma monetaria por cor (mesma logica do diagnostico mais abaixo)
soma_rosa_rem = soma_rosa_ret = soma_laranja_rem = soma_laranja_ret = 0.0
for pandas_idx, (cor, nnf, _) in pintura_p6.items():
    val = df_p6.at[pandas_idx, '_valor']
    if pd.isna(val): continue
    if cor == 'rosa':
        if val > 0: soma_rosa_rem += val
        else:       soma_rosa_ret += val
    elif cor == 'laranja':
        if val > 0: soma_laranja_rem += val
        else:       soma_laranja_ret += val

# ---------------------------------------------------------------
# ABA "RESUMO" — contagem desta rodagem (igual ao VBA antigo)
# ---------------------------------------------------------------
NOME_ABA_RESUMO = 'Resumo Conciliacao'
if NOME_ABA_RESUMO in wb.sheetnames:
    del wb[NOME_ABA_RESUMO]
ws_resumo = wb.create_sheet(NOME_ABA_RESUMO, 0)  # insere como primeira aba

fill_header = PatternFill('solid', fgColor='1F4E79')
font_header = Font(bold=True, color='FFFFFF', size=12)
fill_rosa   = PatternFill('solid', fgColor='F2CEEF')
fill_laranja= PatternFill('solid', fgColor='E97132')
fill_azul   = PatternFill('solid', fgColor='467886')

linhas_resumo = [
    ('RESUMO DA CONCILIAÇÃO', '', None),
    (f'Rodagem em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', '', None),
    ('', '', None),
    ('Categoria', 'Qtde. linhas', None),
    ('Remessas ROSA (100% zerado)', cnt_rem['rosa'], fill_rosa),
    ('Remessas LARANJA (parcial/incompleto)', cnt_rem['laranja'], fill_laranja),
    ('Retornos ROSA (ZEROU)', cnt_ret['rosa'], fill_rosa),
    ('Retornos LARANJA (PARCIAL)', cnt_ret['laranja'], fill_laranja),
    ('Retornos AZUL (sem remessa)', cnt_ret['azul'], fill_azul),
    ('', '', None),
    ('TOTAL DE LINHAS PINTADAS', len(pintura_p6), None),
    ('', '', None),
    ('Diagnóstico monetário', '', None),
    ('Saldo líquido ROSA (deve ser ~R$0)', round(soma_rosa_rem + soma_rosa_ret, 2), None),
    ('Saldo líquido LARANJA (em aberto)', round(soma_laranja_rem + soma_laranja_ret, 2), None),
]

for i, (label, valor, fill) in enumerate(linhas_resumo, start=1):
    c1 = ws_resumo.cell(row=i, column=1, value=label)
    c2 = ws_resumo.cell(row=i, column=2, value=valor if valor != '' else None)
    if fill:
        c1.fill = fill; c2.fill = fill
    if label in ('RESUMO DA CONCILIAÇÃO',):
        c1.font = Font(bold=True, size=16)
    if label == 'Categoria':
        c1.font = font_header; c2.font = font_header
        c1.fill = fill_header; c2.fill = fill_header
    if label == 'TOTAL DE LINHAS PINTADAS':
        c1.font = Font(bold=True); c2.font = Font(bold=True)
    if isinstance(valor, float):
        c2.number_format = '#,##0.00'

ws_resumo.column_dimensions['A'].width = 42
ws_resumo.column_dimensions['B'].width = 16

# Gera nome de saida automaticamente, sem sobrescrever rodagens anteriores
# e sem travar se uma versao anterior estiver aberta no Excel
base_output = ARQUIVO.replace('.xlsx', '_CONCILIADO')
n = 1
while True:
    candidato = f"{base_output}.xlsx" if n == 1 else f"{base_output}_v{n}.xlsx"
    if not os.path.exists(candidato):
        OUTPUT = candidato
        break
    n += 1

wb.save(OUTPUT)
print(f"Arquivo salvo: {OUTPUT}")
print(f"Aba '{NOME_ABA_RESUMO}' criada com o resumo desta rodagem.")

# ---------------------------------------------------------------
# RESUMO NO CONSOLE
# ---------------------------------------------------------------
print()
print('=' * 75)
print('RESULTADO DA CONCILIACAO')
print('=' * 75)
for linha in log_result:
    print(linha)

print()
print('RESUMO DOS RETORNOS (status da conciliacao):')
print(f'  Zerou   (rosa):      {cnt["ZEROU"]}')
print(f'  Parcial (laranja):   {cnt["PARCIAL"]}')
print(f'  Sem remessa (azul):  {cnt["SEM_REMESSA"]}')
print(f'  Sem retorno P6:      {cnt["SEM_RAZAO"]}')
print()
print('RESUMO DAS REMESSAS PINTADAS:')
print(f'  Rosa    (100% zerado por retornos ZEROU): {cnt_rem["rosa"]}')
print(f'  Laranja (parcial ou retorno incompleto):  {cnt_rem["laranja"]}')
print()
print(f'  Total linhas pintadas: {len(pintura_p6)}')
print()
print('Regra aplicada:')
print('  REMESSA fica ROSA apenas se:')
print('    1. Todo seu valor foi consumido (saldo = 0)')
print('    2. TODOS os retornos que a referenciam sao ZEROU')
print('  Caso contrario: LARANJA')
print()
print('Filtrar pela cor ROSA na Planilha6 deve resultar em soma = 0.')
print()

# Diagnostico monetario: soma por cor e tipo (ja calculado mais acima, reaproveitado aqui)
print('DIAGNOSTICO MONETARIO (valores das linhas pintadas):')
print(f'  ROSA   remessas (positivo): R$ {soma_rosa_rem:>15,.2f}')
print(f'  ROSA   retornos (negativo): R$ {soma_rosa_ret:>15,.2f}')
print(f'  ROSA   saldo liquido:        R$ {(soma_rosa_rem + soma_rosa_ret):>15,.2f}  <-- deve ser ~0')
print()
print(f'  LARANJA remessas:            R$ {soma_laranja_rem:>15,.2f}')
print(f'  LARANJA retornos:            R$ {soma_laranja_ret:>15,.2f}')
print(f'  LARANJA saldo liquido:       R$ {(soma_laranja_rem + soma_laranja_ret):>15,.2f}')
print()
saldo_total = soma_rosa_rem + soma_rosa_ret + soma_laranja_rem + soma_laranja_ret
print(f'  TOTAL pintado (rem+ret):     R$ {saldo_total:>15,.2f}')
if abs(soma_rosa_rem + soma_rosa_ret) < 1.0:
    print('  [OK] Rosa balanceado — filtro rosa soma ~R$0')
else:
    print(f'  [ATENCAO] Rosa fora de balanco por R$ {soma_rosa_rem + soma_rosa_ret:,.2f}')
