import requests
import re
import json
from lxml import etree
from collections import Counter

FILE_ID = '1DpKK-YplfAcW9dWkeElaJQY_xWasnCfq'
DOWNLOAD_URL = f'https://drive.google.com/uc?export=download&id={FILE_ID}'

print('Downloading Excel...')
r = requests.get(DOWNLOAD_URL, allow_redirects=True)
r.raise_for_status()
with open('pasantias.xls', 'wb') as f:
    f.write(r.content)
print(f'Downloaded {len(r.content)} bytes')

ns = 'urn:schemas-microsoft-com:office:spreadsheet'
parser = etree.XMLParser(recover=True)
tree = etree.parse('pasantias.xls', parser)
root = tree.getroot()
sheet = root.findall(f'.//{{{ns}}}Worksheet')[0]
rows = sheet.findall(f'.//{{{ns}}}Row')

header = None
records = []
for i, row in enumerate(rows):
    cells = row.findall(f'{{{ns}}}Cell')
    vals = [(c.find(f'{{{ns}}}Data').text if c.find(f'{{{ns}}}Data') is not None and c.find(f'{{{ns}}}Data').text else '') for c in cells]
    if i == 0:
        header = vals
    else:
        records.append(dict(zip(header, vals)))

print(f'Parsed {len(records)} records')

# Build ANIOS_DATA
anios_raw = {}
for r in records:
    f = r.get('Fecha de inicio', '')
    if not f or len(f) < 4 or not f[:4].isdigit():
        continue
    anio = f[:4]
    if anio not in anios_raw:
        anios_raw[anio] = {'total': 0, 'activa': 0, 'finalizada': 0, 'montos': []}
    anios_raw[anio]['total'] += 1
    estado = r.get('estado', '')
    if estado == 'activa':
        anios_raw[anio]['activa'] += 1
    elif estado == 'finalizada':
        anios_raw[anio]['finalizada'] += 1
    try:
        m = float(r.get('Monto estimulo', '0') or 0)
        if m > 0:
            anios_raw[anio]['montos'].append(m)
    except:
        pass

AD = {}
for anio in sorted(anios_raw.keys()):
    d = anios_raw[anio]
    avg = int(sum(d['montos']) / len(d['montos'])) if d['montos'] else 0
    AD[anio] = {'total': d['total'], 'activa': d['activa'], 'finalizada': d['finalizada'], 'montoAvg': avg}

# Build EMPRESAS
empresas = Counter(r.get('Razon social', '').strip() for r in records if r.get('Razon social', '').strip())
top15 = empresas.most_common(15)
otras_total = sum(v for k, v in empresas.items() if k not in dict(top15))
EA = [{'name': k, 'total': v} for k, v in top15]
EA.append({'name': 'Otras empresas', 'total': otras_total})

# Build DBY (duration by year)
dur_labels = [1, 2, 3, 4, 5, 6, 9, 12]
DBY = {}
for r in records:
    f = r.get('Fecha de inicio', '')
    if not f or len(f) < 4 or not f[:4].isdigit():
        continue
    anio = f[:4]
    try:
        dur = int(float(r.get('Duracion en meses', '0') or 0))
    except:
        dur = 0
    if anio not in DBY:
        DBY[anio] = [0] * 8
    if dur in dur_labels:
        DBY[anio][dur_labels.index(dur)] += 1

# Serialize
ad_str = json.dumps(AD, ensure_ascii=False, separators=(',', ':'))
ea_str = json.dumps(EA, ensure_ascii=False, separators=(',', ':'))
dby_parts = [f'"{k}":{json.dumps(v, separators=(",",":"))}' for k, v in sorted(DBY.items())]
dby_str = '{' + ','.join(dby_parts) + '}'

new_block = f'const AD={ad_str};\nconst EA={ea_str};\nconst DBY={dby_str};'

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

html = re.sub(r'const AD=\{[\s\S]*?\};[\s]*const EA=[\s\S]*?\};[\s]*const DBY=\{[\s\S]*?\};', new_block, html)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print('index.html updated successfully!')
