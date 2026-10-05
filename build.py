"""Pra-pemrosesan data BPS + build index.html.
Jalankan: python build.py  (butuh openpyxl, shapely; world-atlas via npm)
"""
import json, openpyxl
from shapely.geometry import shape, mapping

X = 'raw/vidsat_uas.xlsx'
wb = openpyxl.load_workbook(X, data_only=True)
num = lambda v: float(str(v).replace(',', '').replace(' ', '')) if v is not None else None
AKOM = 'Penyediaan Akomodasi dan Makan Minum'

# 1) Wisman menurut negara (25 negara, 2019-2025). Sel "  37, 043" dibersihkan -> 37043
flows = {}
for r in list(wb['negara_asal'].iter_rows(values_only=True))[1:]:
    if not r[0]: continue
    for i, y in enumerate(range(2019, 2026)):
        flows.setdefault(y, []).append({'k': r[0], 'c': r[1], 'v': int(num(r[2 + i]))})
entry = {}
for r in list(wb['Pintu_masuk'].iter_rows(values_only=True))[1:]:
    entry[r[0]] = {y: r[1 + i] for i, y in enumerate(range(2019, 2026))}

# 2) PDRB lapangan usaha (miliar Rp). Klungkung tercatat x1000 (12.003.728,92) -> dibagi 1000
pd = []
sh1 = {r[0]: r for r in list(wb['Sheet1'].iter_rows(values_only=True))[1:]}
for r in list(wb['hirarki_pdrb'].iter_rows(values_only=True))[1:]:
    if r[0] is None: continue
    val = r[4]
    if val is None and r[3].startswith('Penyediaan Akomodasi') and r[0] in sh1:  # sel kosong -> isi dari Sheet1
        val = sh1[r[0]][3]
    if val is None: continue
    v = float(val) / (1000 if r[0] == 'Klungkung' else 1)
    sec = AKOM if r[3].startswith('Penyediaan Akomodasi') else r[3]
    pd.append({'k': r[0], 'y': int(r[1]), 'g': r[2], 's': sec, 'v': round(v, 2)})
# Sheet1 (total & akomodasi 2025 per kab/kota, termasuk Tabanan); Klungkung /1000
s1 = []
for r in list(wb['Sheet1'].iter_rows(values_only=True))[1:]:
    f = 1000 if r[2] > 1e6 else 1
    s1.append({'k': r[0], 't': round(r[2] / f, 2), 'a': round(r[3] / f, 2)})
# validasi silang dengan Sheet1
for o in s1:
    tot = sum(p['v'] for p in pd if p['k'] == o['k'] and p['y'] == 2025)
    print('cek', o['k'], round(tot, 1), o['t'])

# 3) Geometri: sederhanakan + adjacency (queen contiguity)
gj = json.load(open('raw/Kabupaten-Kota__Provinsi_Bali_.geojson'))
geoms, feats = {}, []
for f in gj['features']:
    n = f['properties']['NAME_2']
    g = shape(f['geometry']); geoms[n] = g
    s = g.simplify(0.0015, preserve_topology=True)
    m = json.loads(json.dumps(mapping(s)))
    rnd = lambda c: [round(c[0], 4), round(c[1], 4)] if isinstance(c[0], float) else [rnd(x) for x in c]
    m['coordinates'] = rnd(m['coordinates'])
    feats.append({'type': 'Feature', 'properties': {'name': n, 'kode': f['properties']['CC_2'], 'tipe': f['properties']['TYPE_2']}, 'geometry': m})
names = list(geoms)
adj = [[a, b] for i, a in enumerate(names) for b in names[i + 1:] if geoms[a].buffer(0.002).intersects(geoms[b])]
print('adjacency', adj)

land = json.load(open('node_modules/world-atlas/land-110m.json'))
coords = {'Australia': [134, -25], 'India': [78, 22], 'China': [104, 35], 'Korea Selatan': [127.8, 36.5], 'Jepang': [138, 36],
          'Taiwan': [121, 23.7], 'Malaysia': [102, 4], 'Singapore': [103.8, 1.35], 'Philippines': [122, 12.5], 'Thailand': [101, 15],
          'Vietnam': [106, 16], 'Myanmar': [96, 21], 'Cambodia': [104.9, 12.5], 'Laos': [103.8, 18.2], 'Brunei Darussalam': [114.7, 4.5],
          'United Kingdom': [-2, 54], 'France': [2.5, 46.5], 'Germany': [10.4, 51], 'Russia': [60, 58], 'Netherlands': [5.3, 52.2],
          'Italy': [12.5, 42.5], 'Spain': [-3.7, 40.2], 'United States': [-98, 39], 'Canada': [-100, 58], 'Afrika Selatan': [24, -29]}
AKS = '4 Oktober 2026'
src = {
    'wisman': {'judul': 'Banyaknya Wisatawan Mancanegara yang Datang Langsung ke Bali menurut Kebangsaan', 'tahun': '2019-2025',
               'url': 'https://bali.bps.go.id/en/statistics-table?subject=561', 'akses': AKS},
    'pintu': {'judul': 'Banyaknya Wisatawan Mancanegara Bulanan ke Bali menurut Pintu Masuk', 'tahun': '2019-2025',
              'url': 'https://bali.bps.go.id/en/statistics-table/2/MTA2IzI=/banyaknya-wisatawan-mancanegara-bulanan-ke-bali-menurut-pintu-masuk.html', 'akses': AKS},
    'pdrb': {'judul': 'PDRB Kabupaten/Kota menurut Lapangan Usaha (Tinjauan Regional PDRB Kab/Kota, Buku 2 Jawa-Bali)', 'tahun': '2021-2025',
             'url': 'https://www.bps.go.id/id/publication/2025/09/15/a1daf2f268d53d17506677ac', 'akses': AKS},
    'batas': {'judul': 'Batas wilayah kab/kota Bali, GADM v4', 'tahun': '-', 'url': 'https://gadm.org/', 'akses': AKS}}
D = {'flows': flows, 'entry': entry, 'pdrb': pd, 's1': s1, 'geo': {'type': 'FeatureCollection', 'features': feats}, 'adj': adj,
     'land': land, 'coords': coords, 'bali': [115.19, -8.4], 'src': src}
json.dump(D, open('data/data.json', 'w'), separators=(',', ':'))
html = open('template.html', encoding='utf-8').read().replace('/*__DATA__*/null', json.dumps(D, separators=(',', ':')))
open('index.html', 'w', encoding='utf-8').write(html)
print('OK', len(html) // 1024, 'KB')
