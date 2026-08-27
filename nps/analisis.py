# -*- coding: utf-8 -*-
"""
NPS Banco Agrario — motivo de las malas notas y agentes con peor calificación.

NPS = %Promotores - %Detractores, sobre las encuestas con clasificación.
Se usa la columna NPS del archivo, NO la nota cruda: voz y chat manejan
escalas distintas (un 8 es promotor en voz y neutro en chat).
"""
import sys, glob, pandas as pd

def cargar(rutas):
    partes=[]
    for r in rutas:
        d=pd.read_csv(r)
        d['_origen']=r.split('/')[-1].replace('.csv','')
        partes.append(d)
    d=pd.concat(partes, ignore_index=True)
    d['NPS']=d['NPS'].astype(str).str.strip()
    d=d[d['NPS'].isin(['Promotores','Neutros','Detractores'])].copy()
    for c in ['macroproceso','proceso_new','subproceso_uno','AgentName','Nombre Limpio','SurveyName']:
        d[c]=d[c].astype(str).str.strip().replace({'nan':'(sin dato)','':'(sin dato)'})
    return d

def nps(g):
    """NPS y desglose de un grupo de encuestas."""
    tot=len(g)
    p=(g['NPS']=='Promotores').sum()
    n=(g['NPS']=='Neutros').sum()
    dt=(g['NPS']=='Detractores').sum()
    return pd.Series({'Encuestas':tot,'Promotores':p,'Neutros':n,'Detractores':dt,
                      '%Detractores':round(dt/tot*100,1) if tot else 0,
                      'NPS':round((p-dt)/tot*100,1) if tot else 0})

def por(d, col, minimo=0):
    t=d.groupby(col, group_keys=False).apply(nps, include_groups=False)
    t=t[t['Encuestas']>=minimo].sort_values('NPS')
    return t.reset_index()

def titulo(t):
    print('\n'+'═'*78); print(' '+t); print('═'*78)

def tabla(df, cols=None, n=None):
    v=df if n is None else df.head(n)
    if cols: v=v[cols]
    print(v.to_string(index=False))

if __name__=='__main__':
    rutas=sys.argv[1:] or sorted(glob.glob('nps/*.csv'))
    d=cargar(rutas)
    meses=', '.join(sorted(d['Mes'].dropna().astype(str).unique()))

    titulo(f'RESUMEN — {meses}  ({len(d):,} encuestas)'.replace(',','.'))
    g=nps(d)
    print(f"  NPS general: {g['NPS']}%   |  Promotores {g['Promotores']:,}  "
          f"Neutros {g['Neutros']:,}  Detractores {g['Detractores']:,} "
          f"({g['%Detractores']}%)".replace(',','.'))
    print('\n  Por canal:')
    tabla(por(d,'SurveyName'))
    if d['Mes'].nunique()>1:
        print('\n  Por mes:')
        tabla(por(d,'Mes'))

    # ── 1. EL MOTIVO: dónde nos ponen mala nota ──────────────────────────
    det=d[d['NPS']=='Detractores']

    titulo('1a. MOTIVO POR VOLUMEN — dónde se concentran los detractores')
    print('  Cuántos detractores aporta cada tipificación. Arreglar lo de arriba')
    print('  es lo que más mueve el indicador.\n')
    for col,nm in [('macroproceso','MACROPROCESO'),('proceso_new','PROCESO'),
                   ('subproceso_uno','SUBPROCESO')]:
        v=det[col].value_counts().head(10).rename_axis(nm).reset_index(name='Detractores')
        v['% del total']=(v['Detractores']/len(det)*100).round(1)
        base=d[col].value_counts()
        v['Encuestas']=v[nm].map(base).astype(int)
        v['%Detrac.']=(v['Detractores']/v['Encuestas']*100).round(1)
        print(f'  ── {nm} ──'); print(v.to_string(index=False)); print()

    titulo('1b. MOTIVO POR TASA — dónde peor nos califican')
    print('  NPS de cada tipificación, de peor a mejor. Se piden mínimo 30')
    print('  encuestas para que el dato no dependa de dos o tres respuestas.\n')
    for col,nm in [('proceso_new','PROCESO'),('subproceso_uno','SUBPROCESO')]:
        t=por(d,col,minimo=30).head(12)
        t=t.rename(columns={col:nm})
        print(f'  ── {nm} ──')
        print(t[[nm,'Encuestas','Detractores','%Detractores','NPS']]
              .astype({'Encuestas':int,'Detractores':int}).to_string(index=False))
        print()

    # ── 2. AGENTES CON PEOR NOTA ─────────────────────────────────────────
    da=d[~d['AgentName'].isin(['(sin dato)'])]
    MIN=int(sys.argv[0] and 20)          # mínimo de encuestas para entrar al ranking
    ag=por(da,'AgentName',minimo=MIN)

    titulo(f'2. AGENTES CON PEOR NPS  (mínimo {MIN} encuestas)')
    npsgen=nps(d)['NPS']
    peores=ag[ag['NPS']<npsgen]
    print(f'  NPS general: {npsgen}%. Están por debajo {len(peores)} de {len(ag)} agentes.')
    print(f'  Se listan los 10 peores y todos los demás que estén por debajo.\n')
    v=ag.head(max(10,len(peores))).copy()
    v.insert(0,'#',range(1,len(v)+1))
    v=v.rename(columns={'AgentName':'AGENTE'})
    print(v[['#','AGENTE','Encuestas','Promotores','Detractores','%Detractores','NPS']]
          .astype({'Encuestas':int,'Promotores':int,'Detractores':int}).to_string(index=False))

    excl=por(da,'AgentName').pipe(lambda t: t[t['Encuestas']<MIN])
    if len(excl):
        print(f'\n  ({len(excl)} agentes quedaron fuera por tener menos de {MIN} encuestas)')
