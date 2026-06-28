"""
limpar_xmls.py
--------------
Limpa a pasta de XMLs antes da conciliacao, aplicando as regras:

  -nfe-110111.xml  Cancelamento         -> apaga TODOS os XMLs daquela nota
  -nfe-210240.xml  Operacao nao Realiz. -> apaga TODOS os XMLs daquela nota
  -nfe-210210.xml  Ciencia da Operacao  -> apaga APENAS o arquivo de evento (mantem -nfe.xml)

Uso:
    py -3.14 limpar_xmls.py
    py -3.14 limpar_xmls.py --pasta="INNFE_20260622154435"
    py -3.14 limpar_xmls.py --simular
"""

import os, sys, re
from collections import defaultdict

PASTA_PADRAO = 'INNFE_20260622154435'

# Eventos que invalidam a nota inteira
EVENTOS_CANCELAR_TUDO  = {'110111', '210240'}
# Eventos que devem ser apagados mas a nota continua valida
EVENTOS_APAGAR_EVENTO  = {'210210', '210200', '110110'}  # CC-e e Ciencia — nota continua valida

# ---------------------------------------------------------------
pasta   = PASTA_PADRAO
simular = False
for arg in sys.argv[1:]:
    if arg == '--simular':
        simular = True
    elif arg.startswith('--pasta'):
        pasta = arg.split('=', 1)[-1].strip().strip('"')

if simular:
    print("*** MODO SIMULACAO — nenhum arquivo sera apagado ***")
print(f"Pasta: {pasta}")
print()

# ---------------------------------------------------------------
# LEITURA DOS ARQUIVOS
# Padrao: {chave44}-nfe.xml  ou  {chave44}-nfe-{codigo}.xml
# ---------------------------------------------------------------
RE_NOME = re.compile(r'^([0-9]{44})-nfe(?:-([0-9]+))?\.xml$', re.IGNORECASE)

try:
    arquivos = os.listdir(pasta)
except FileNotFoundError:
    print(f"ERRO: pasta '{pasta}' nao encontrada.")
    sys.exit(1)

por_chave = defaultdict(dict)   # chave -> {'nfe': fname, '110111': fname, ...}
ignorados = []

for fname in arquivos:
    if not fname.lower().endswith('.xml'):
        continue
    m = RE_NOME.match(fname)
    if m:
        chave  = m.group(1)
        codigo = m.group(2) or 'nfe'   # sem codigo = e a nota principal
        por_chave[chave][codigo] = fname
    else:
        ignorados.append(fname)

print(f"Total XMLs encontrados : {len([f for f in arquivos if f.endswith('.xml')])}")
print(f"Chaves unicas          : {len(por_chave)}")
if ignorados:
    print(f"Fora do padrao (ignorados): {ignorados[:5]}")
print()

# ---------------------------------------------------------------
# CONTA TIPOS
# ---------------------------------------------------------------
so_nfe      = sum(1 for v in por_chave.values() if set(v) == {'nfe'})
com_eventos = sum(1 for v in por_chave.values() if len(v) > 1)
print(f"Notas so com -nfe.xml  : {so_nfe}")
print(f"Notas com eventos      : {com_eventos}")
todos_eventos = set()
for v in por_chave.values():
    todos_eventos |= set(v) - {'nfe'}
print(f"Tipos de evento encontrados: {sorted(todos_eventos)}")
print()

# ---------------------------------------------------------------
# APLICA REGRAS
# ---------------------------------------------------------------
apagar  = []
log     = []
mantidas = 0

for chave, cod_map in sorted(por_chave.items()):
    nf_num = int(chave[25:34])
    codigos_presentes = set(cod_map.keys())
    eventos_invalida  = codigos_presentes & EVENTOS_CANCELAR_TUDO

    if eventos_invalida:
        # Apaga TUDO desta chave (nota + todos os eventos)
        for cod, fname in cod_map.items():
            apagar.append(os.path.join(pasta, fname))
        ev_str = ', '.join(sorted(eventos_invalida))
        log.append(f'[REMOVE TUDO ] NF {nf_num:>8} | eventos: {ev_str} | {len(cod_map)} arquivo(s)')
    else:
        # Apaga apenas eventos de baixo impacto
        apagou_evento = False
        for ev in EVENTOS_APAGAR_EVENTO:
            if ev in cod_map:
                apagar.append(os.path.join(pasta, cod_map[ev]))
                apagou_evento = True
        if apagou_evento:
            log.append(f'[REMOVE EVENTO] NF {nf_num:>8} | manteve -nfe.xml, removeu eventos')
        else:
            mantidas += 1

# ---------------------------------------------------------------
# RELATORIO
# ---------------------------------------------------------------
print("ACOES:")
print("-" * 65)
if log:
    for l in log: print(l)
else:
    print("Nenhuma acao — pasta ja esta limpa.")
print()
print(f"Arquivos a apagar    : {len(apagar)}")
print(f"Notas sem alteracao  : {mantidas}")
print()

# ---------------------------------------------------------------
# EXECUTA
# ---------------------------------------------------------------
if apagar and not simular:
    print("Apagando...")
    ok, erros = 0, 0
    for fpath in apagar:
        try:
            os.remove(fpath)
            ok += 1
        except Exception as e:
            print(f"  ERRO: {os.path.basename(fpath)} -> {e}")
            erros += 1
    restantes = len([f for f in os.listdir(pasta) if f.endswith('.xml')])
    print(f"Concluido: {ok} apagados, {erros} erros.")
    print(f"XMLs restantes: {restantes} (-nfe.xml validos para conciliacao)")
elif apagar and simular:
    print("Arquivos que seriam apagados:")
    for f in apagar:
        print(f"  {os.path.basename(f)}")
