"""Consolida as planilhas semanais do Projeto Centro numa base única (CSV).
Uso: python consolidar_boletins.py <pasta_com_xlsx_de_um_ano> <ano> <saida.csv>
Ex.:  python consolidar_boletins.py boletins/2024 2024 base_2024.csv
Gera base_2024.csv (para análise) e base_2024.js (o arquivo que o painel lê).
"""
import openpyxl, glob, re, sys, unicodedata, pandas as pd
src, ANO, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
MESES = {'JAN':1,'FEV':2,'MAR':3,'ABR':4,'MAI':5,'JUN':6,'JUL':7,'AGO':8,'SET':9,'OUT':10,'NOV':11,'DEZ':12}
KEEP = ['NUM_BO','ANO_BO','NOME_DELEGACIA','NOME_DELEGACIA_CIRC','DATA_OCORRENCIA_BO','HORA_OCORRENCIA_BO',
        'RUBRICA','LOGRADOURO','BAIRRO','LATITUDE','LONGITUDE','APRESENTACAO','PERIODO',
        'TIPO LOCAL','TIPO DE CRIME','MODUS OPERANDI','MEIO DE LOCOMOÇÃO INFRATOR','LOCAL DE SUBTRAÇÃO DO OBJETO',
        'OBJETO VISADO PELO CRIMINOSO','N_VITIMAS','NATUREZA ALTERADA PARA','CLASSIFICACAO FINAL']
SHIFTED = ['PERIODO','TIPO LOCAL','TIPO DE CRIME']
frames = []
for f in sorted(glob.glob(src + '/*.xlsx')):
    m = re.search(r'(\d{2})__SEMANA', f)
    sem = int(m.group(1))
    dm = re.search(r'(\d{2})([A-Z]{3})_A_(\d{2})([A-Z]{3})', f.split('/')[-1].upper())
    ini = pd.Timestamp(ANO, MESES[dm.group(2)], int(dm.group(1)))
    fim = pd.Timestamp(ANO + (1 if MESES[dm.group(4)] < MESES[dm.group(2)] else 0), MESES[dm.group(4)], int(dm.group(3)))
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    h = [str(x).strip() if x is not None else '' for x in rows[0]]
    idx = {c: h.index(c) for c in KEEP if c in h}
    # Detecta coluna de histórico inserida sem cabeçalho (desloca PERIODO/TIPO LOCAL/TIPO DE CRIME em +1)
    amostra = [r[idx['PERIODO']] for r in rows[1:30] if r and len(r) > idx['PERIODO']]
    if sum(bool(re.match(r'^(MADRUGADA|MANH|TARDE|NOITE)', str(v or ''))) for v in amostra) < len(amostra) / 2:
        for c in SHIFTED: idx[c] += 1
        print(f'  semana {sem:02d}: colunas deslocadas corrigidas')
    recs = []
    for r in rows[1:]:
        if r is None or all(v is None for v in r[:5]): continue
        recs.append({c: (r[i] if i < len(r) else None) for c, i in idx.items()})
    df = pd.DataFrame(recs); df['SEMANA'] = sem; df['SEMANA_INICIO'] = ini.date(); df['SEMANA_FIM'] = fim.date()
    frames.append(df)
d = pd.concat(frames, ignore_index=True)

def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return None
    s = re.sub(r'\s+', ' ', str(v)).strip().upper()
    return None if s in ('', 'NULL', 'NAN', 'NONE') else s
for c in KEEP:
    if c in d and c not in ('DATA_OCORRENCIA_BO', 'HORA_OCORRENCIA_BO', 'LATITUDE', 'LONGITUDE', 'N_VITIMAS'):
        d[c] = d[c].map(txt)

def dp(v):
    v = v if isinstance(v, str) else ''
    if re.match(r'^0?3', v): return '3º DP Campos Elíseos'
    if v.startswith('77'): return '77º DP Santa Cecília'
    return 'Outra DP'
d['DP'] = d['NOME_DELEGACIA_CIRC'].map(dp)
d['DATA'] = pd.to_datetime(d['DATA_OCORRENCIA_BO'].map(lambda x: x if isinstance(x, pd.Timestamp) else pd.to_datetime(str(x), dayfirst=True, errors='coerce'))).dt.date
d['HORA'] = d['HORA_OCORRENCIA_BO'].astype(str).str.extract(r'(\d{1,2}):')[0].astype(float)

NI = 'NÃO INFORMADO'
def norm_modus(v):
    v = v if isinstance(v, str) else NI
    v = v.replace('INFRATOR QUEBROU', 'QUEBROU')
    if v.startswith('GRAVE AMEAÇA SEM ARMA'): return 'GRAVE AMEAÇA SEM ARMA/SIMULACRO'
    if v.startswith('AMEAÇA COM ARMA DE FOGO'): return 'AMEAÇA COM ARMA DE FOGO/SIMULACRO'
    if v in ('MODUS OPERANDI NAO ESPECIFICADO', 'NAO ESPECIFICADO', 'BOLETIM SEM INFORMACOES', 'SEM INFORMACOES'): return NI
    return v
def norm_tipo(v):
    v = v if isinstance(v, str) else NI
    if v in ('SEM INFORMACAO',): return NI
    v = v.replace('ESTAÇAO', 'ESTAÇÃO').replace('ESTACAO', 'ESTAÇÃO').replace('SUBTRAÇAO', 'SUBTRAÇÃO')
    if v.startswith('FURTO EM INTERIOR DE ESTAÇÃO'): return 'FURTO EM INTERIOR DE ESTAÇÃO METRO/TREM/ÔNIBUS'
    return v
def norm_local(v):
    v = v if isinstance(v, str) else NI
    v2 = unicodedata.normalize('NFKD', v).encode('ascii', 'ignore').decode()
    if v2 in ('VIA PUBLICA',): return 'VIA PÚBLICA'
    if 'METROVIARIO' in v2 or v2.startswith('ESTACAO'): return 'METRÔ / TREM'
    if v2 in ('A PE', 'TRANSEUNTE'): return 'VIA PÚBLICA'
    if v2.startswith('ESTACIONAMENTO'): return 'ESTACIONAMENTO'
    if v2.startswith('APARTAMENTO') or v2 in ('CASA', 'AREA COMUM', 'CONDOMINIO RESIDENCIAL'): return 'RESIDÊNCIA / CONDOMÍNIO'
    if v2 in ('ESTABELECIMENTO COMERCIAL', 'COMERCIO', 'LOJAS', 'MERCADO', 'FARMACIA/DROGARIA', 'CONVENIENCIA', 'AUTOPECAS', 'DEPOSITO', 'SUPERMERCADO', 'SHOPPING CENTER'): return 'COMÉRCIO'
    if v2 in ('BAR/BOTEQUIM', 'CAFE/LANCHONETE', 'RESTAURANTE', 'HOTEL'): return 'BAR / RESTAURANTE / HOTEL'
    if v2 in ('FERROVIARIO',): return 'METRÔ / TREM'
    if 'VEICULO' in v2 or 'ONIBUS' in v2 or 'TRANSPORTE COLETIVO' in v2: return 'VEÍCULO / TRANSPORTE'
    if v2.startswith('AGENCIA'): return 'AGÊNCIA BANCÁRIA'
    if v2 in ('PRACA',): return 'PRAÇA'
    return v
d['CLASSIFICACAO'] = d['CLASSIFICACAO FINAL'].fillna(NI)
d['TIPO_CRIME'] = d['TIPO DE CRIME'].map(norm_tipo)
d['MODUS'] = d['MODUS OPERANDI'].map(norm_modus)
d['MEIO'] = d['MEIO DE LOCOMOÇÃO INFRATOR'].fillna(NI).replace({'NÃO ESPECIFICADO': NI})
d['LOCAL_SUBTRACAO'] = d['LOCAL DE SUBTRAÇÃO DO OBJETO'].fillna(NI).replace({'NÃO ESPECIFICADO': NI})
d['OBJETO'] = d['OBJETO VISADO PELO CRIMINOSO'].fillna(NI).replace({'NÃO ESPECIFICADO': NI})
d['TIPO_LOCAL'] = d['TIPO LOCAL'].map(norm_local)
d['APRESENTACAO'] = d['APRESENTACAO'].fillna(NI)
d['BAIRRO'] = d['BAIRRO'].fillna(NI).map(lambda s: unicodedata.normalize('NFKD', s).encode('ascii','ignore').decode())
d['ANO'] = ANO
for c in ('LATITUDE', 'LONGITUDE'):
    d[c] = pd.to_numeric(d[c].astype(str).str.replace(',', '.'), errors='coerce').round(5)
ok = d['LATITUDE'].between(-23.75, -23.35) & d['LONGITUDE'].between(-46.85, -46.4)
d.loc[~ok, ['LATITUDE', 'LONGITUDE']] = None
d['BAIRRO'] = d['BAIRRO'].replace({'SANTA EFIGENIA': 'SANTA IFIGENIA', 'SANTA HIFIGENIA': 'SANTA IFIGENIA',
                                   'CENTRO HISTORICO DE SAO PAULO': 'CENTRO', 'SE': 'CENTRO'})
d['OBJETO'] = d['OBJETO'].str.replace(r'^JOIAS(/RELÓGIO)?$', 'JOIAS/RELÓGIO', regex=True).str.replace(r'^JOIAS(/RELÓGIO)? E OUTROS$', 'JOIAS/RELÓGIO E OUTROS', regex=True)
d['TIPO_CRIME'] = d['TIPO_CRIME'].replace({'01 - SEM SELEÇÃO': NI})
raros = d['TIPO_LOCAL'].value_counts(); raros = raros[raros < 10].index
d.loc[d['TIPO_LOCAL'].isin(raros), 'TIPO_LOCAL'] = 'OUTROS'
d['ID'] = d['ANO'].astype(str) + '-' + d['SEMANA'].astype(str).str.zfill(2) + '-' + (d.groupby('SEMANA').cumcount() + 1).astype(str).str.zfill(4)
cols = ['ID','ANO','SEMANA','SEMANA_INICIO','SEMANA_FIM','NUM_BO','DP','DATA','HORA','CLASSIFICACAO','RUBRICA','TIPO_CRIME','MODUS','MEIO',
        'LOCAL_SUBTRACAO','OBJETO','TIPO_LOCAL','APRESENTACAO','BAIRRO','LOGRADOURO','LATITUDE','LONGITUDE']
d[cols].to_csv(out, index=False)
# Versão .js da mesma base: é a que o painel (index.html) carrega.
import json, os
js = os.path.splitext(out)[0] + '.js'
with open(js, 'w', encoding='utf-8') as fh:
    fh.write('registrarBase(' + json.dumps(d[cols].to_csv(index=False), ensure_ascii=False) + ');\n')
print(d.shape, '->', out, 'e', js)
