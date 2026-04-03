import pandas as pd, base64, os, sys, openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── LOGO BASE64 (embebido) ──────────────────────────────────────────────────
LOGO = "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" 

script_dir = os.path.dirname(os.path.abspath(__file__))

# ── ENCONTRAR EXCEL ───────────────────────────────────────────────────────────
excel_file = None
for fname in os.listdir(script_dir):
    if fname.endswith('.xlsx') or fname.endswith('.xlsm'):
        try:
            xl = pd.ExcelFile(os.path.join(script_dir, fname))
            if 'BD' in xl.sheet_names and 'INCIDENCIA BARRIDO' in xl.sheet_names:
                excel_file = os.path.join(script_dir, fname)
                break
        except: pass

if not excel_file:
    input("ERROR: No se encontro el Excel. Presiona Enter para salir.")
    sys.exit(1)

print(f"Leyendo: {os.path.basename(excel_file)}")

# ════════════════════════════════════════════════════════════════════════
# PARTE 1 — LEER Y PROCESAR DATA
# ════════════════════════════════════════════════════════════════════════

# ── BD ────────────────────────────────────────────────────────────────────────
df = pd.read_excel(excel_file, sheet_name='BD', header=0)
df.columns = [str(c).strip() for c in df.columns]
df = df.drop(columns=[c for c in df.columns if 'Unnamed' in str(c)], errors='ignore')
col_zona   = next((c for c in df.columns if 'ZONA' in c.upper()), df.columns[0])
col_pasillo= next((c for c in df.columns if 'PASILLO' in c.upper()), df.columns[1])
col_nivel  = next((c for c in df.columns if 'NIVEL' in c.upper()), df.columns[3])
col_area   = next((c for c in df.columns if 'AREA' in c.upper() or 'CODIGO' in c.upper()), df.columns[6])
col_uni    = next((c for c in df.columns if 'UNIDAD' in c.upper()), df.columns[7])
col_status = next((c for c in df.columns if 'STATUS' in c.upper()), df.columns[8])
df['_ZONA']   = df[col_zona].astype(str).str.strip()
df['_PAS']    = df[col_pasillo].astype(str).str.strip()
df['_NIVEL']  = pd.to_numeric(df[col_nivel], errors='coerce')
df['_AREA']   = df[col_area].astype(str).str.strip().str.upper()
df['_UNI']    = pd.to_numeric(df[col_uni], errors='coerce').fillna(0)
df['_STATUS'] = df[col_status].astype(str).str.strip().str.upper()
for val in ['0','1','0.0','1.0']:
    df.loc[df['_STATUS']==val, '_STATUS'] = 'OCUPADA' if val in ['1','1.0'] else 'VACIA'

# ── INCIDENCIA BARRIDO ────────────────────────────────────────────────────────
df_bar = pd.read_excel(excel_file, sheet_name='INCIDENCIA BARRIDO', header=0)
df_bar.columns = [str(c).strip() for c in df_bar.columns]
col_ubi = next((c for c in df_bar.columns if 'UBIC' in c.upper()), df_bar.columns[0])
col_brd = next((c for c in df_bar.columns if 'BARRIDO' in c.upper()), df_bar.columns[1])
col_by  = next((c for c in df_bar.columns if c.upper()=='BY'), df_bar.columns[2])
col_fds = next((c for c in df_bar.columns if 'FUERA' in c.upper()), df_bar.columns[3])
df_bar['_BAR'] = pd.to_numeric(df_bar[col_brd], errors='coerce').fillna(0)
df_bar['_BY']  = pd.to_numeric(df_bar[col_by],  errors='coerce').fillna(0)
df_bar['_FDS'] = df_bar[col_fds].astype(str).str.strip().str.upper()
incid_bar = df_bar[
    (df_bar['_BAR']==1) & (df_bar['_BY']==0) & (df_bar['_FDS'].isin(['0','','NAN']))
][[ col_ubi, col_brd, col_by, col_fds ]].copy().reset_index(drop=True)
incid_bar.columns = ['UBICACION','BARRIDO','BY','FUERA_SISTEMA']

# ── INCIDENCIAS DB (BASE - PMZ PISO 1) ────────────────────────────────────────
try:
    df_base = pd.read_excel(excel_file, sheet_name='BASE', header=0)
    df_base.columns = [str(c).encode('utf-8','ignore').decode('utf-8').strip() for c in df_base.columns]
    col_map = {}
    for c in df_base.columns:
        cl = c.lower()
        if 'ubicaci' in cl: col_map[c]='UBICACION'
        elif 'numero de articulo' in cl: col_map[c]='CODIGO'
        elif 'descripci' in cl: col_map[c]='DESCRIPCION'
        elif 'id_area' in cl: col_map[c]='AREA'
        elif 'cantidad de unidades' in cl: col_map[c]='CANTIDAD'
        elif 'estado' in cl: col_map[c]='ESTADO'
    df_base = df_base.rename(columns=col_map)
    piso1 = df_base[df_base['AREA']=='AMZZPISO1'].copy()
    piso1['UBICACION'] = piso1['UBICACION'].astype(str).str.strip()
    piso1['CODIGO']    = piso1['CODIGO'].astype(str).str.strip()
    piso1['DESCRIPCION'] = piso1['DESCRIPCION'].astype(str).str.strip()
    incid_db = piso1.groupby('CODIGO').agg(
        DESCRIPCION=('DESCRIPCION','first'),
        N_UBICACIONES=('UBICACION','nunique'),
        UBICACIONES=('UBICACION', lambda x: ' | '.join(sorted(x.unique()))),
        CANTIDAD_TOTAL=('CANTIDAD','sum'),
        ESTADO=('ESTADO','first')
    ).reset_index()
    incid_db = incid_db[incid_db['N_UBICACIONES'] > 1].sort_values('N_UBICACIONES', ascending=False).reset_index(drop=True)
except Exception as e:
    print(f"  Aviso BASE: {e}")
    incid_db = pd.DataFrame(columns=['CODIGO','DESCRIPCION','N_UBICACIONES','UBICACIONES','CANTIDAD_TOTAL','ESTADO'])

print(f"  BD: {len(df):,} registros")
print(f"  Incidencias Barrido: {len(incid_bar)}")
print(f"  Incidencias DB PMZ P1: {len(incid_db)}")

# ════════════════════════════════════════════════════════════════════════
# PARTE 2 — CALCULAR KPIs PARA DASHBOARD HTML
# ════════════════════════════════════════════════════════════════════════
ocu = df[df['_STATUS']=='OCUPADA']
vac = df[df['_STATUS']=='VACIA']
fds = df[df['_STATUS']=='FUERA DE SISTEMA']
total_uni = int(ocu['_UNI'].sum())
grand = len(df)

def zona_stats(z):
    zdf = df[df['_ZONA']==z]
    zo = zdf[zdf['_STATUS']=='OCUPADA']
    return len(zo), len(zdf[zdf['_STATUS']=='VACIA']), len(zdf[zdf['_STATUS']=='FUERA DE SISTEMA']), int(zo['_UNI'].sum())

def area_stats(key):
    a = df[df['_AREA'].str.contains(key, na=False)]
    ao = a[a['_STATUS']=='OCUPADA']
    return len(ao), len(a[a['_STATUS']=='VACIA']), len(a[a['_STATUS']=='FUERA DE SISTEMA']), int(ao['_UNI'].sum())

prs_o,prs_v,prs_f,prs_u = zona_stats('PRS')
pmz_o,pmz_v,pmz_f,pmz_u = zona_stats('PMZ')
pgr_o,pgr_v,pgr_f,pgr_u = zona_stats('PGR')
rs_o,rs_v,rs_f,rs_u = area_stats('RACK SELECTIVO')
m1_o,m1_v,m1_f,m1_u = area_stats('PRIMER PISO')
m2_o,m2_v,m2_f,m2_u = area_stats('SEGUNDO PISO')
rg_o,rg_v,rg_f,rg_u = area_stats('GRILLA')

nivs = []
prs_df = df[df['_ZONA']=='PRS']
for n in range(1,10):
    ndf = prs_df[prs_df['_NIVEL']==n]
    nivs.append((n, len(ndf[ndf['_STATUS']=='OCUPADA']), len(ndf[ndf['_STATUS']=='FUERA DE SISTEMA'])))

hm = {'PRS':[], 'PMZ':[], 'PGR':[]}
for pas, grp in df.groupby('_PAS'):
    z = grp['_ZONA'].iloc[0]
    if z not in hm: continue
    go = grp[grp['_STATUS']=='OCUPADA']
    hm[z].append({'p':str(pas),'o':len(go),'v':len(grp[grp['_STATUS']=='VACIA']),'f':len(grp[grp['_STATUS']=='FUERA DE SISTEMA']),'u':int(go['_UNI'].sum())})

def hmjs(lst):
    parts = []
    for r in lst:
        parts.append('{p:' + "'" + r['p'] + "'" + ',o:' + str(r['o']) + ',v:' + str(r['v']) + ',f:' + str(r['f']) + ',u:' + str(r['u']) + '}')
    return '[' + ','.join(parts) + ']'

def bar_js(df_inc):
    parts = []
    for r in df_inc.itertuples(index=False):
        parts.append('{u:' + "'" + str(r.UBICACION) + "'" + ',b:' + str(r.BARRIDO) + ',by:' + str(r.BY) + ',f:' + "'" + str(r.FUERA_SISTEMA) + "'" + '}')
    return '[' + ','.join(parts) + ']'

def db_js(df_inc):
    parts = []
    for r in df_inc.itertuples(index=False):
        desc = str(r.DESCRIPCION).replace("'", "`")
        ubis = str(r.UBICACIONES).replace("'", "`")
        parts.append('{c:' + "'" + str(r.CODIGO) + "'" + ',d:' + "'" + desc + "'" + ',n:' + str(r.N_UBICACIONES) + ',u:' + "'" + ubis + "'" + ',q:' + str(int(r.CANTIDAD_TOTAL)) + '}')
    return '[' + ','.join(parts) + ']'

niv_js = ','.join('{n:' + "'N" + str(n) + "'" + ',o:' + str(o) + ',f:' + str(f) + '}' for n,o,f in nivs)

js = "const R={"
js += "global:{OCUPADA:" + str(len(ocu)) + ",VACIA:" + str(len(vac)) + ",FDS:" + str(len(fds)) + ",UNIDADES:" + str(total_uni) + "},"
js += "zonas:["
js += "{z:'PRS',o:" + str(prs_o) + ",v:" + str(prs_v) + ",f:" + str(prs_f) + ",u:" + str(prs_u) + "},"
js += "{z:'PMZ',o:" + str(pmz_o) + ",v:" + str(pmz_v) + ",f:" + str(pmz_f) + ",u:" + str(pmz_u) + "},"
js += "{z:'PGR',o:" + str(pgr_o) + ",v:" + str(pgr_v) + ",f:" + str(pgr_f) + ",u:" + str(pgr_u) + "}],"
js += "areas:["
js += "{a:'Rack Selectivo',z:'PRS',o:" + str(rs_o) + ",v:" + str(rs_v) + ",f:" + str(rs_f) + ",u:" + str(rs_u) + "},"
js += "{a:'Mezz. 1er Piso',z:'PMZ',o:" + str(m1_o) + ",v:" + str(m1_v) + ",f:" + str(m1_f) + ",u:" + str(m1_u) + "},"
js += "{a:'Mezz. 2do Piso',z:'PMZ',o:" + str(m2_o) + ",v:" + str(m2_v) + ",f:" + str(m2_f) + ",u:" + str(m2_u) + "},"
js += "{a:'Rack Grilla',z:'PGR',o:" + str(rg_o) + ",v:" + str(rg_v) + ",f:" + str(rg_f) + ",u:" + str(rg_u) + "}],"
js += "niv:[" + niv_js + "],"
js += "hm:{PRS:" + hmjs(hm['PRS']) + ",PMZ:" + hmjs(hm['PMZ']) + ",PGR:" + hmjs(hm['PGR']) + "},"
js += "incid_bar:" + str(len(incid_bar)) + ","
js += "incid_db:" + str(len(incid_db)) + ","
js += "incid_bar_data:" + bar_js(incid_bar) + ","
js += "incid_db_data:" + db_js(incid_db)
js += "};"

# ── LEER TEMPLATE Y GENERAR HTML ──────────────────────────────────────────────
tpl_path = os.path.join(script_dir, 'html_template.txt')
if not os.path.exists(tpl_path):
    input(f"ERROR: No se encontro html_template.txt. Presiona Enter.")
    sys.exit(1)
with open(tpl_path, 'r', encoding='utf-8') as f:
    html = f.read()
html = html.replace('DATA_PLACEHOLDER', js)
out_html = os.path.join(script_dir, 'index.html')
with open(out_html, 'w', encoding='utf-8') as f:
    f.write(html)

# ════════════════════════════════════════════════════════════════════════
# PARTE 3 — ACTUALIZAR HOJAS DE INCIDENCIAS EN EL EXCEL
# ════════════════════════════════════════════════════════════════════════
ORANGE="E8581A"; GRAY="4A4A4A"; WHITE="FFFFFF"; LGRAY="F2F2F2"; RED="C0392B"
def fill(c): return PatternFill("solid",fgColor=c)
def font(c=WHITE,sz=11,bold=False): return Font(name="Arial",size=sz,bold=bold,color=c)
def align(h="center",v="center"): return Alignment(horizontal=h,vertical=v)
def brd():
    s=Side(border_style="thin",color="DDDDDD")
    return Border(left=s,right=s,top=s,bottom=s)

wb = load_workbook(excel_file)

# Eliminar hojas viejas
for sname in ['INCID. BARRIDO','INCID. DB - PMZ P1']:
    if sname in wb.sheetnames:
        del wb[sname]

# ── HOJA 1: INCID. BARRIDO ────────────────────────────────────────────────────
ws1 = wb.create_sheet('INCID. BARRIDO')
ws1.sheet_properties.tabColor = RED
ws1.sheet_view.showGridLines = False
for ci,w in enumerate([3,28,12,12,22,3]):
    ws1.column_dimensions[get_column_letter(ci+1)].width = w
for r in range(1, len(incid_bar)+10):
    ws1.row_dimensions[r].height = 18

for row in ws1.iter_rows(min_row=1, max_row=len(incid_bar)+10, min_col=1, max_col=7):
    for cell in row: cell.fill = fill(LGRAY)

ws1.row_dimensions[2].height = 38
ws1.merge_cells("B2:F2")
c=ws1.cell(2,2,f"  ⚠  INCIDENCIAS BARRIDO — PALLETS MAL UBICADOS  ({datetime.now().strftime('%d/%m/%Y %H:%M')})")
c.font=font(WHITE,12,True); c.fill=fill(RED); c.alignment=align("left","center")

ws1.row_dimensions[4].height = 20
ws1.merge_cells("B4:F4")
c=ws1.cell(4,2,f"  Total: {len(incid_bar)} incidencias  |  Lógica: BARRIDO=1, BY=0, Sin fuera de sistema")
c.font=font(GRAY,9); c.fill=fill("FDECEA"); c.alignment=align("left","center")

ws1.row_dimensions[6].height = 22
for ci,h in enumerate(['UBICACION','BARRIDO','BY','FUERA DE SISTEMA']):
    c=ws1.cell(6,ci+2,h); c.font=font(WHITE,9,True)
    c.fill=fill(GRAY); c.alignment=align(); c.border=brd()

if len(incid_bar)==0:
    ws1.merge_cells("B7:F7")
    c=ws1.cell(7,2,"  ✓  Sin incidencias en los datos actuales")
    c.font=font("15803D",10,True); c.fill=fill("F0FFF4"); c.alignment=align("left","center")
else:
    for ri,row in enumerate(incid_bar.itertuples(index=False)):
        r=7+ri; ws1.row_dimensions[r].height=17
        bg=WHITE if ri%2==0 else "FEF2F2"
        for ci,val in enumerate([row.UBICACION,row.BARRIDO,row.BY,row.FUERA_SISTEMA]):
            c=ws1.cell(r,ci+2,val); c.font=font(GRAY,9)
            c.fill=fill(bg); c.alignment=align("left" if ci==0 else "center"); c.border=brd()

ws1.freeze_panes="B7"
ws1.auto_filter.ref="B6:F6"

# ── HOJA 2: INCID. DB - PMZ P1 ────────────────────────────────────────────────
ws2 = wb.create_sheet('INCID. DB - PMZ P1')
ws2.sheet_properties.tabColor = ORANGE
ws2.sheet_view.showGridLines = False
for ci,w in enumerate([3,18,42,16,16,55,16,3]):
    ws2.column_dimensions[get_column_letter(ci+1)].width = w
for r in range(1, len(incid_db)+10):
    ws2.row_dimensions[r].height = 18

for row in ws2.iter_rows(min_row=1, max_row=len(incid_db)+10, min_col=1, max_col=9):
    for cell in row: cell.fill = fill(LGRAY)

ws2.row_dimensions[2].height = 38
ws2.merge_cells("B2:G2")
c=ws2.cell(2,2,f"  ⚠  INCIDENCIAS D.B — CÓDIGOS EN MÁS DE 1 UBICACIÓN PMZ PISO 1  ({datetime.now().strftime('%d/%m/%Y %H:%M')})")
c.font=font(WHITE,12,True); c.fill=fill(ORANGE); c.alignment=align("left","center")

ws2.row_dimensions[4].height = 20
ws2.merge_cells("B4:G4")
c=ws2.cell(4,2,f"  Total: {len(incid_db)} códigos con más de 1 ubicación  |  Área: AMZZPISO1")
c.font=font(GRAY,9); c.fill=fill("FFF8F0"); c.alignment=align("left","center")

ws2.row_dimensions[6].height = 22
for ci,h in enumerate(['CÓDIGO','DESCRIPCIÓN','N° UBIC.','CANTIDAD','UBICACIONES','ESTADO']):
    c=ws2.cell(6,ci+2,h); c.font=font(WHITE,9,True)
    c.fill=fill(GRAY); c.alignment=align(); c.border=brd()

if len(incid_db)==0:
    ws2.merge_cells("B7:G7")
    c=ws2.cell(7,2,"  ✓  Sin códigos duplicados en PMZ Primer Piso en los datos actuales")
    c.font=font("15803D",10,True); c.fill=fill("F0FFF4"); c.alignment=align("left","center")
else:
    for ri,row in enumerate(incid_db.itertuples(index=False)):
        r=7+ri; ws2.row_dimensions[r].height=17
        bg=WHITE if ri%2==0 else "FFF5EC"
        vals=[row.CODIGO,row.DESCRIPCION,row.N_UBICACIONES,int(row.CANTIDAD_TOTAL),row.UBICACIONES,row.ESTADO]
        for ci,val in enumerate(vals):
            c=ws2.cell(r,ci+2,val); c.font=font(GRAY,9)
            c.fill=fill(bg); c.alignment=align("left" if ci in [1,4] else "center"); c.border=brd()
            if ci in [2,3]: c.number_format="#,##0"

ws2.freeze_panes="B7"
ws2.auto_filter.ref="B6:G6"

wb.save(excel_file)

# ════════════════════════════════════════════════════════════════════════
# RESULTADO FINAL
# ════════════════════════════════════════════════════════════════════════
pct = len(ocu)/grand*100 if grand else 0
print()
print("="*55)
print("  TODO GENERADO EXITOSAMENTE!")
print("="*55)
print(f"  Dashboard:          index.html")
print(f"  Total ubicaciones:  {grand:,}")
print(f"  Ocupacion:          {pct:.1f}%")
print(f"  Unidades stock:     {total_uni:,}")
print(f"  Incid. Barrido:     {len(incid_bar)}")
print(f"  Incid. DB PMZ P1:   {len(incid_db)}")
print(f"  Fecha:              {datetime.now().strftime('%d/%m/%Y %H:%M')}")
print()
print("  SIGUIENTE PASO:")
print("  Sube index.html a GitHub Pages")
print("="*55)
input("\nPresiona Enter para cerrar...")
