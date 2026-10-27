# Verificacion independiente: replica en pandas las reglas de BD_CLASIFICADA y compara con el libro recalculado.
# Uso: python3 -I verificar_clasificacion.py <base_original.xlsx> <libro_recalculado.xlsx>
import pandas as pd, numpy as np, json, openpyxl, datetime
import sys
SRC, OUT = sys.argv[1], sys.argv[2]
x=pd.read_excel(OUT,sheet_name='BD_CLASIFICADA',dtype=object)
o=pd.read_excel(SRC,sheet_name='BD FINAL',dtype=object).iloc[:,:31]
num=lambda s: pd.to_numeric(s,errors='coerce')
ing=num(o.INGRESO); fon=num(o.FONDO_COMPENSATORIO); ssf=num(o.SALDO_SSF); ant=num(o.ANTIGUEDAD_ANIOS); edad=num(o.EDAD)
inst=o.INSTITUCION.fillna('').astype(str).str.strip()
ok=ing>0
def cut(v,cuts,labels):
    return labels[np.searchsorted(cuts,v,side='right')]
exp={}
bd=pd.to_datetime(o.CLI_FECHA_NACIMIENTO); d=pd.Timestamp('2026-09-30')
er=d.year-bd.dt.year-((bd.dt.month>d.month)|((bd.dt.month==d.month)&(bd.dt.day>d.day))).astype(int)
exp['EDAD_RECALCULADA']=er
exp['NIVEL_ANTIGUEDAD']=[ 'Sin dato' if pd.isna(a) else 'Inconsistente - revisar' if (a<0 or a>e) else 'Alto' if a<=1 else 'Medio' if a<=4 else 'Bajo' for a,e in zip(ant,edad)]
exp['RANGO_INGRESO']=['Sin dato' if pd.isna(v) else 'Ingreso cero - revisar' if v<=0 else cut(v,[1500,3000,5000],['I - Ingreso bajo','II - Ingreso medio','III - Ingreso medio alto','IV - Ingreso alto']) for v in ing]
exp['RANGO_DEUDA']=[('Sin registro SSF' if i=='' else 'Dato faltante') if pd.isna(v) else 'Sin saldo (cero o a favor)' if v<=0 else cut(v,[1000,10000,50000],['D1 - Deuda baja','D2 - Deuda intermedia','D3 - Deuda elevada','D4 - Deuda muy elevada']) for v,i in zip(ssf,inst)]
dti=np.where(ok&ssf.notna(), ssf.clip(lower=0)/ing, np.nan)
exp['RANGO_DEUDA_INGRESO']=['No calculable - ingreso no válido' if not k else 'No calculable - sin dato SSF' if pd.isna(s) else 'Sin deuda' if r==0 else cut(r,[3,12,24],['R1 - Bajo','R2 - Medio','R3 - Alto','R4 - Muy alto']) for k,s,r in zip(ok,ssf,dti)]
exp['RANGO_FONDO']=['Sin fondo (cero)' if v<=0 else cut(v,[3000,15000,50000],['F1 - Fondo bajo','F2 - Fondo medio','F3 - Fondo alto','F4 - Fondo muy alto']) for v in fon]
foi=np.where(ok,fon/ing,np.nan)
exp['RANGO_FONDO_INGRESO']=['No calculable - ingreso no válido' if not k else 'Sin fondo' if f<=0 else cut(r,[1,6,12],['M1 - Cobertura baja','M2 - Cobertura media','M3 - Cobertura alta','M4 - Cobertura muy alta']) for k,f,r in zip(ok,fon,foi)]
cati=o.MAX_CAT_RIESGO.fillna('').str.strip(); cate=o.CALIFICACION_SSF.fillna('').str.strip()
exp['CAT_INTERNA']=[c if c else ('Dato faltante (con crédito vigente)' if cr=='SI' else 'Sin historial interno') for c,cr in zip(cati,o.TIENE_CREDITO_VIGENTE)]
exp['CAT_EXTERNA_SSF']=[c if c else ('Dato faltante (con saldo SSF)' if pd.notna(s) else 'Sin registro SSF') for c,s in zip(cate,ssf)]
crit=pd.DataFrame({
 'e':np.where(er.between(30,55),'Cumple','No cumple'),
 'i':np.where(ing.isna(),'Sin información',np.where(ing>0,'Cumple','No cumple')),
 'f':np.where(fon>0,'Cumple','No cumple'),
 'a':np.where(pd.Series(exp['NIVEL_ANTIGUEDAD']).isin(['Sin dato','Inconsistente - revisar']),'Sin información','Cumple'),
 'ci':np.where(cati.isin(['A1','A2']),'Cumple',np.where(cati=='','Sin información','No cumple')),
 'ce':np.where(cate.isin(['A1','A2']),'Cumple',np.where(cate=='','Sin información','No cumple'))})
exp['ELEGIBILIDAD_GENERAL']=np.where((crit=='No cumple').any(axis=1),'No cumple criterio verificable',np.where((crit=='Sin información').any(axis=1),'Información incompleta','Cumple criterios verificables'))
exp['ELEGIBILIDAD_TC']=np.where(o.TIENE_TARJETA=='SI','Excluido: ya tiene tarjeta',exp['ELEGIBILIDAD_GENERAL'])
def look(v,tab):
    los=[t[0] for t in tab]; return tab[np.searchsorted(los,v,side='right')-1][1]
TC=[(0,'NO APLICA'),(500,500),(750,1500),(1000,2000),(1500,2500),(1700,3000),(2000,3500),(2500,4500),(3000,5000),(3500,7000),(7500,8000),(8000,10000)]
VEH=[(0,'NO APLICA'),(3000,20000),(5000,25000),(10000,30000),(15000,40000)]
HIPO=[(0,'NO APLICA'),(3000,110000),(5000,150000),(9000,205000)]
for k,nm,tab in (('PRECA TC','TC',TC),('PRECA VEHICULO','VEHICULO',VEH),('PRECA HIPO','HIPO',HIPO)):
    p=num(o[k])
    exp['CONSIST_'+nm]=['No pre evaluado' if pd.isna(pv) else 'Pre evaluado sin ingreso válido' if not (iv>0) else ('Coincide con tabla' if look(iv,tab)==pv else 'No coincide con tabla') for pv,iv in zip(p,ing)]
bad=0
for c,v in exp.items():
    a=pd.Series(x[c]).astype(str).values; b=pd.Series(v).astype(str).values
    if c=='EDAD_RECALCULADA': a=num(x[c]).values; b=np.array(v)
    mism=(a!=b).sum(); bad+=mism
    print(f'{c:24s} mismatches={mism}', '' if mism==0 else list(zip(a[a!=b][:3],b[a!=b][:3])))
# numeric
for c,v in (('MULT_DEUDA_INGRESO',dti),('MESES_FONDO_INGRESO',foi)):
    a=num(x[c]).values; m=~(np.isnan(a)&np.isnan(v)); print(c,'max abs diff',np.nanmax(abs(a[m]-v[m])), 'nan mismatch',(np.isnan(a)!=np.isnan(v)).sum())
# originals preserved
oo=pd.read_excel(SRC,sheet_name='BD FINAL',dtype=object).iloc[:,:31]
xx=x.iloc[:,:31]
print('original cols identical:', oo.equals(xx), list(oo.columns)==list(xx.columns))
print('TOTAL MISMATCHES',bad)
for c in exp: print(c, pd.Series(exp[c]).value_counts().to_dict())
