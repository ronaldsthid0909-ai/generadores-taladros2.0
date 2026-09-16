import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.graph_objects as go

st.set_page_config(page_title='Optimización de Generadores', page_icon='⚡', layout='wide')

LOW_DEFAULT = 20.0
HOURS_DEFAULT = 5.0
ACTIVE_DEFAULT = 1.0

st.markdown('''
<style>
.stApp{background:#071522;color:#F5F8FC}.block-container{padding-top:1.5rem!important;max-width:1500px!important}
section[data-testid="stSidebar"]{background:#081827;border-right:1px solid #17324D}
section[data-testid="stSidebar"] *{color:#EAF1F8!important}
section[data-testid="stSidebar"] input{color:#142536!important;background:#fff!important}
.dashboard-title{font-size:25px;font-weight:800;color:#fff;border-bottom:2px solid #00B7E5;padding:0 0 10px;margin-bottom:18px;line-height:1.2}
[data-testid="stFileUploader"]{background:#102B4A;border:1px solid #173E63;border-radius:12px;padding:8px}
[data-testid="stFileUploader"] *{color:#EAF1F8!important}
div[data-baseweb="select"]>div{background:#F7F9FC!important;border:1px solid #00B7E5!important;border-radius:8px!important}
div[data-baseweb="select"] *{color:#16283A!important}
.metric-card{background:#102B4A;border-radius:12px;min-height:88px;padding:13px 15px 11px 18px;position:relative;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.18)}
.metric-card:before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:#4FA3E3}.yellow:before{background:#FFC000}.green:before{background:#3A8E24}.red:before{background:#FF3B16}.cyan:before{background:#00B7E5}
.metric-label{color:#B8C7D6;font-size:11px;font-weight:700;margin-bottom:9px}.metric-value{color:#fff;font-size:27px;font-weight:800;line-height:1}.metric-sub{color:#9EB0C2;font-size:11px;margin-top:7px}
.section-title{color:#fff;font-size:20px;font-weight:800;margin:10px 0}.panel{background:#102B4A;border-radius:16px;padding:18px 20px}.panel-title{color:#fff;font-size:16px;font-weight:800;margin-bottom:14px}.criterion{color:#EAF1F8;font-size:13px;line-height:1.5;margin-bottom:13px}.criterion b{color:#fff}
[data-testid="stDataFrame"]{border-radius:8px;overflow:hidden}
</style>''', unsafe_allow_html=True)

def card(label, value, cls='blue', sub=''):
    return f'''<div class="metric-card {cls}"><div class="metric-label">{label}</div><div class="metric-value">{value}</div>{f'<div class="metric-sub">{sub}</div>' if sub else ''}</div>'''

def timestamp_col(df):
    cols=[]
    for c in df.columns:
        t=str(c).lower()
        if any(x in t for x in ['timestamp','datetime','date time','fecha','time']): cols.append(c)
    return cols[0] if cols else None

def generator_cols(df):
    out={}
    for c in df.columns:
        t=str(c).strip().lower()
        if not any(x in t for x in ['gen','generator','generador']): continue
        m=re.search(r'(?:gen(?:erator|erador)?)[\s_\-]*(\d+)',t,re.I)
        if not m: continue
        n=int(m.group(1))
        if any(x in t for x in ['percent power used','power used','load','carga','percent','%']) or n not in out: out[n]=c
    return dict(sorted(out.items()))

def kw_cols(df):
    out={}
    for c in df.columns:
        t=str(c).lower()
        if not any(x in t for x in ['gen','generator','generador']) or not any(x in t for x in ['kw','power']): continue
        m=re.search(r'(?:gen(?:erator|erador)?)[\s_\-]*(\d+)',t,re.I)
        if m: out[int(m.group(1))]=c
    return dict(sorted(out.items()))

def prepare(df, ts, gens):
    d=df.copy(); d.columns=[str(c).strip() for c in d.columns]
    ts2=next((c for c in d.columns if c==str(ts)),ts)
    d[ts2]=pd.to_datetime(d[ts2],errors='coerce'); d=d.dropna(subset=[ts2]).sort_values(ts2).reset_index(drop=True)
    for n,c in gens.items():
        d[c]=pd.to_numeric(d[c].astype(str).str.replace('%','',regex=False).str.replace(',','.',regex=False),errors='coerce')
    return d

def avg_no_zero(s):
    s=pd.to_numeric(s,errors='coerce'); s=s[(s.notna())&(s!=0)]
    return np.nan if s.empty else float(s.mean())

def detect(d,ts,gens,low,min_h,active):
    loads=pd.DataFrame({n:pd.to_numeric(d[c],errors='coerce') for n,c in gens.items()})
    act=loads>active; lowact=act&(loads<=low); cond=(act.sum(axis=1)>=2)&(lowact.sum(axis=1)>=2)
    times=d[ts]; diff=times.diff().dt.total_seconds()/3600
    valid=diff[(diff>0)&diff.notna()]; med=float(valid.median()) if not valid.empty else 1/60
    gaps=diff>max(med*3,1/6); groups=cond.ne(cond.shift()).cumsum()+gaps.fillna(False).cumsum(); rows=[]
    for _,b in d.groupby(groups):
        idx=b.index
        if len(idx)==0 or not bool(cond.loc[idx].all()): continue
        start,end=times.loc[idx[0]],times.loc[idx[-1]]; dur=(end-start).total_seconds()/3600+(min(med,1) if len(idx)>=2 else 0)
        if dur<=min_h: continue
        bl=loads.loc[idx]; involved=[n for n in gens if (bl[n]>active).any()]; lowgens=[n for n in gens if ((bl[n]>active)&(bl[n]<=low)).any()]
        if len(lowgens)<2: continue
        rows.append({'Inicio':start,'Fin':end,'Duración (h)':round(dur,2),'GEN involucrados':' + '.join(f'GEN {n}' for n in involved),'Carga promedio GEN involucrados (%)':round(float(bl[involved].mean().mean()),2),'Generadores baja carga':' + '.join(f'GEN {n}' for n in lowgens)})
    return pd.DataFrame(rows).sort_values('Duración (h)',ascending=False).reset_index(drop=True) if rows else pd.DataFrame()

def analyze(df, low, min_h, active):
    df=df.copy(); df.columns=[str(c).strip() for c in df.columns]; ts=timestamp_col(df); gens=generator_cols(df); kws=kw_cols(df)
    if ts is None:return None,'No se encontró columna de fecha/hora.'
    if len(gens)<2:return None,'No se encontraron al menos 2 generadores.'
    d=prepare(df,ts,gens)
    if d.empty:return None,'No hay datos válidos.'
    return {'data':d,'ts':ts,'gens':gens,'kws':kws,'events':detect(d,ts,gens,low,min_h,active)},None

def load_chart(rig,r,low):
    d,ts,gens,events=r['data'],r['ts'],r['gens'],r['events']; fig=go.Figure(); colors={1:'#39A7E8',2:'#FFC000',3:'#3C9B20',4:'#FF3B16'}
    for n,c in gens.items(): fig.add_trace(go.Scatter(x=d[ts],y=d[c],mode='lines',name=f'GEN {n}',line=dict(color=colors.get(n,'#B8C7D8'),width=1.8)))
    fig.add_trace(go.Scatter(x=[d[ts].min(),d[ts].max()],y=[low,low],mode='lines',name=f'Umbral {low:g}%',line=dict(color='white',width=2,dash='dot')))
    for _,e in events.iterrows(): fig.add_vrect(x0=e['Inicio'],x1=e['Fin'],fillcolor='#3C78A8',opacity=.22,line_width=0,layer='below')
    fig.add_annotation(xref='paper',x=.985,y=low,yref='y',text=f'UMBRAL {low:g}%',showarrow=False,xanchor='right',yanchor='bottom',font=dict(color='white',size=10),bgcolor='#102B4A')
    ymax=max(50,float(d[list(gens.values())].max().max())+3)
    fig.update_layout(title=dict(text=f'Carga de los 4 generadores — Rig {rig}',font=dict(size=16,color='white'),x=.01),paper_bgcolor='#102B4A',plot_bgcolor='#102B4A',font=dict(color='#EAF1F8'),height=400,margin=dict(l=50,r=25,t=50,b=45),legend=dict(x=.86,y=.97,bgcolor='rgba(0,0,0,0)',font=dict(color='white',size=10)),xaxis=dict(title='Tiempo',gridcolor='#24415E',tickfont=dict(color='#D7E2EC'),title_font=dict(color='#EAF1F8')),yaxis=dict(title='Carga (%)',range=[0,ymax],gridcolor='#24415E',tickfont=dict(color='#D7E2EC'),title_font=dict(color='#EAF1F8')),hovermode='x unified')
    return fig

def fleet_chart(summary):
    s=summary.copy(); fig=go.Figure(go.Bar(x=s['Taladro'].astype(str),y=s['Horas en eventos'],text=[f'{v:.1f} h' for v in s['Horas en eventos']],textposition='outside',marker_color='#39A7E8'))
    fig.update_layout(title=dict(text='Horas en eventos prolongados por taladro',x=.01,font=dict(size=16,color='white')),paper_bgcolor='#102B4A',plot_bgcolor='#102B4A',font=dict(color='#EAF1F8'),height=320,margin=dict(l=45,r=25,t=50,b=45),xaxis=dict(type='category',title='Taladro',gridcolor='#24415E',tickfont=dict(color='#EAF1F8')),yaxis=dict(title='Horas',gridcolor='#24415E',tickfont=dict(color='#EAF1F8')))
    return fig

with st.sidebar:
    st.markdown('## ⚙️ Configuración')
    low=st.number_input('Umbral de baja carga (%)',1.0,100.0,LOW_DEFAULT,1.0)
    min_h=st.number_input('Duración mínima para alerta (h)',.1,1000.0,HOURS_DEFAULT,.5)
    active=st.number_input('Umbral para GEN activo (%)',0.0,20.0,ACTIVE_DEFAULT,.5)
    st.markdown('---')
    st.markdown(f'<div style="color:#AFC0D1;font-size:12px;line-height:1.5"><b style="color:white">Regla:</b><br>≥2 GEN activos + ≥2 GEN en baja carga<br>+ duración continua &gt; {min_h:g} h.<br><br><b style="color:white">Promedios:</b><br>los valores 0% se excluyen.</div>',unsafe_allow_html=True)

st.markdown('<div class="dashboard-title">RIG — CARGA INDIVIDUAL DE LOS GENERADORES</div>',unsafe_allow_html=True)
up=st.file_uploader('📁 Carga el Excel consolidado',type=['xlsx','xls','csv'])
if up is None:
    st.info('Carga un Excel con una hoja por taladro: 992, M47, M48, X38, X40, X42, X43 y X45.')
    st.stop()
try:
    if up.name.lower().endswith('.csv'): sheets={'992':pd.read_csv(up)}
    else:
        xl=pd.ExcelFile(up); sheets={s:pd.read_excel(up,sheet_name=s) for s in xl.sheet_names}
except Exception as e:
    st.error(f'No fue posible leer el archivo: {e}'); st.stop()

analyses={}; errors={}
for rig,df in sheets.items():
    if df is None or df.empty: continue
    result,err=analyze(df,low,min_h,active)
    if result: analyses[str(rig)]=result
    else: errors[str(rig)]=err
if not analyses: st.error('Ninguna hoja pudo ser analizada.'); st.stop()

selected=st.selectbox('📋 Seleccione el taladro',['Todos los taladros']+list(analyses.keys()))

if selected=='Todos los taladros':
    rows=[]
    for rig,r in analyses.items():
        e=r['events']; hours=float(e['Duración (h)'].sum()) if not e.empty else 0; mx=float(e['Duración (h)'].max()) if not e.empty else 0
        maxgens=max((len(re.findall(r'GEN\s*\d+',str(x))) for x in e['GEN involucrados']),default=0)
        rows.append({'Taladro':str(rig),'Eventos >5 h':len(e),'Horas en eventos':round(hours,2),'Máx. duración (h)':round(mx,2),'Máx. GEN activos':maxgens})
    summary=pd.DataFrame(rows); total_events=int(summary['Eventos >5 h'].sum()); total_hours=float(summary['Horas en eventos'].sum()); max_duration=float(summary['Máx. duración (h)'].max()) if not summary.empty else 0
    c=st.columns(4)
    for col,label,val,cls in [(c[0],'TALADROS CARGADOS',len(summary),'blue'),(c[1],'EVENTOS >5 H',total_events,'yellow'),(c[2],'HORAS ACUMULADAS',f'{total_hours:.2f} h','green'),(c[3],'MÁX. DURACIÓN',f'{max_duration:.1f} h','red')]:
        with col: st.markdown(card(label,val,cls),unsafe_allow_html=True)
    left,right=st.columns([1.6,1])
    with left: st.plotly_chart(fleet_chart(summary.sort_values('Taladro',key=lambda s:s.astype(str))),use_container_width=True,config={'displayModeBar':False})
    with right: st.markdown(f'<div class="panel"><div class="panel-title">CRITERIO PARA DETECTAR EVENTOS</div><div class="criterion"><b>Condición:</b> 2 o más generadores activos con carga ≤ {low:g}%.</div><div class="criterion"><b>Alerta prolongada:</b> condición durante &gt; {min_h:g} h continuas.</div><div class="criterion"><b>Promedios:</b> los valores 0% no participan.</div></div>',unsafe_allow_html=True)
    st.markdown('<div class="section-title">RESUMEN POR TALADRO</div>',unsafe_allow_html=True); st.dataframe(summary.sort_values('Horas en eventos',ascending=False),use_container_width=True,hide_index=True)
    all_events=[]
    for rig,r in analyses.items():
        for _,e in r['events'].iterrows(): all_events.append({'Taladro':rig,**e.to_dict()})
    if all_events:
        ed=pd.DataFrame(all_events).sort_values('Duración (h)',ascending=False); st.markdown('<div class="section-title">EVENTOS DETECTADOS EN LA FLOTA</div>',unsafe_allow_html=True); st.dataframe(ed,use_container_width=True,hide_index=True); st.download_button('⬇️ Descargar eventos de toda la flota',ed.to_csv(index=False).encode('utf-8-sig'),'eventos_flota.csv','text/csv')
else:
    r=analyses[selected]; d,ts,gens,kws,e=r['data'],r['ts'],r['gens'],r['kws'],r['events']
    avgs={n:avg_no_zero(d[c]) for n,c in gens.items()}; vals=[v for v in avgs.values() if pd.notna(v)]; overall=float(np.mean(vals)) if vals else np.nan
    kwvals=[]
    for n,c in kws.items():
        s=pd.to_numeric(d[c],errors='coerce'); s=s[(s.notna())&(s!=0)]
        if not s.empty: kwvals.append(float(s.mean()))
    overallkw=float(np.mean(kwvals)) if kwvals else np.nan
    cards=st.columns(6); specs=[(1,'Promedio de Potencia de GEN 1','blue'),(2,'Promedio de Potencia de GEN 2','yellow'),(3,'Promedio de Potencia de GEN 3','green'),(4,'Promedio de Potencia de GEN 4','red')]
    for col,(n,label,cls) in zip(cards[:4],specs):
        v=avgs.get(n,np.nan); text='N/D' if pd.isna(v) else f'{v:.2f}%'; col.markdown(card(label,text,cls),unsafe_allow_html=True)
    cards[4].markdown(card('Promedio General de carga','N/D' if pd.isna(overall) else f'{overall:.2f}%','blue'),unsafe_allow_html=True)
    cards[5].markdown(card('Promedio General de Potencia','N/D' if pd.isna(overallkw) else f'{overallkw:,.2f} kW','cyan'),unsafe_allow_html=True)
    st.markdown('<br>',unsafe_allow_html=True)
    left,right=st.columns([2.15,1])
    with left: st.plotly_chart(load_chart(selected,r,low),use_container_width=True,config={'displayModeBar':False})
    hours=float(e['Duración (h)'].sum()) if not e.empty else 0; mx=float(e['Duración (h)'].max()) if not e.empty else 0; maxactive=max((len(re.findall(r'GEN\s*\d+',str(x))) for x in e['GEN involucrados']),default=0)
    with right:
        st.markdown(f'<div class="panel"><div class="panel-title">CRITERIO PARA DETECTAR EVENTOS</div><div class="criterion">- <b>Condición de revisión:</b> 2 o más generadores activos con carga ≤ {low:g}%.</div><div class="criterion">- <b>Alerta prolongada:</b> condición durante &gt; {min_h:g} h continuas.</div><div class="criterion">- <b>Exclusión de ceros:</b> 0% no se considera en promedios.</div></div><br>',unsafe_allow_html=True)
        a,b=st.columns(2); a.markdown(card('EVENTOS >5 H',len(e),'yellow',f'{hours:.2f} h acumuladas'),unsafe_allow_html=True); b.markdown(card('MÁX. DURACIÓN',f'{mx:.1f} h','red','evento más prolongado'),unsafe_allow_html=True)
        a,b=st.columns(2); a.markdown(card('MÁX. GEN ACTIVOS',maxactive,'green',f'{len(gens)} disponibles'),unsafe_allow_html=True); b.markdown(card('HORAS EN EVENTOS',f'{hours:.2f} h','cyan','total horas acumuladas'),unsafe_allow_html=True)
    st.markdown('<div class="section-title">EVENTOS DETECTADOS</div>',unsafe_allow_html=True)
    if e.empty: st.success(f'No se encontraron eventos con ≥2 generadores activos, ≥2 en ≤ {low:g}% y duración > {min_h:g} h.')
    else:
        de=e.copy(); de['Inicio']=de['Inicio'].dt.strftime('%Y-%m-%d %H:%M:%S'); de['Fin']=de['Fin'].dt.strftime('%Y-%m-%d %H:%M:%S'); st.dataframe(de,use_container_width=True,hide_index=True); st.download_button('⬇️ Descargar eventos de este taladro',e.to_csv(index=False).encode('utf-8-sig'),f'eventos_{selected}.csv','text/csv')

if errors:
    with st.expander('⚠️ Hojas que no pudieron analizarse'):
        for rig,err in errors.items(): st.write(f'**{rig}:** {err}')
