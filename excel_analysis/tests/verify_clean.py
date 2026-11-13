"""Independent check of the clean views. Usage: python verify_clean.py <source.xlsx> <output.xlsx>
Rules re-implemented from scratch: 4 brands (no Virseker); lines Traditional, Competition,
OQ PNV+PV (no E2E), Direct Marketing, Hippo (Branded+Unconverted+Unbranded); targets summed
over included lines; conversion/effectiveness weighted by contacts of lines carrying the rate."""
import sys, openpyxl
src_p,out_p=sys.argv[1],sys.argv[2]
SRC=openpyxl.load_workbook(src_p,read_only=True,data_only=True)
OUT=openpyxl.load_workbook(out_p,read_only=True,data_only=True)
num=lambda v:isinstance(v,(int,float)) and not isinstance(v,bool)
N=lambda v: v if num(v) else 0.0
G=lambda S,r,c: S[r][c] if r<len(S) and c is not None and c<len(S[r]) else None
fails=[]; checked=0
def close(e,g,tol=1e-6):
    global checked; checked+=1
    return (e is None and not num(g)) or (e is not None and num(g) and abs(e-g)<=tol*max(1,abs(e)))
BRANDS=['Budget','First For Women','Dialdirect','Auto & General']
LEAVES=['Generic - Traditional','E-Commerce - Traditional','Generic - Competition','E-Commerce - Competition','OQ - PNV','OQ - PV','Direct Marketing','Hippo - Branded','Hippo - Unconverted','Hippo - Unbranded']
ROWS=[('Traditional (Generic + E-Commerce)',[0,1]),('Generic - Traditional',[0]),('E-Commerce - Traditional',[1]),('Competition (Generic + E-Commerce)',[2,3]),
      ('Generic - Competition',[2]),('E-Commerce - Competition',[3]),('Online Quoting (PNV + PV, excl. End to End)',[4,5]),('OQ - PNV',[4]),('OQ - PV',[5]),
      ('Direct Marketing',[6]),('Hippo (Branded + Transfers)',[7,8,9]),('Hippo - Branded',[7]),('Hippo - Transfers (Unconverted + Unbranded)',[8,9]),('CLEAN TOTAL',list(range(10)))]
GROUPS=[('cxd','CXD'),('closing','CLOSE'),('gross effectiveness','GEFF'),('net effectiveness','NEFF'),('conversion','CONV'),('contacts','CON'),('leads','LEA'),('net sales','NS'),('gross sales','SAL')]
MEAS={'actual':'A','projecting':'P','target':'T','tar dif':'D','tar var':'V','w target':'W','w tar var':'WV'}
for sname in SRC.sheetnames:
    S=[list(r) for r in SRC[sname].iter_rows(values_only=True)]
    hr=lc=None
    for r in range(10):
        for c in range(8):
            if isinstance(G(S,r,c),str) and S[r][c].startswith('From Date'): hr,lc=r,c; break
        if hr is not None: break
    if hr is None: continue
    cm={}; cur=None
    for c in range(lc+1,len(S[hr])):
        t=(S[hr][c] or '').lower() if isinstance(S[hr][c],str) else ''
        g=next((v for k,v in GROUPS if k in t),None) if t else None
        cur=g or cur
        m=MEAS.get((S[hr+1][c] or '').strip().lower()) if isinstance(S[hr+1][c],str) else None
        if cur and m and (cur,m) not in cm: cm[(cur,m)]=c
    if ('CON','P') not in cm: cm[('CON','P')]=cm[('CON','A')]
    for g in ('LEA','SAL'):
        if (g,'P') not in cm: cm[(g,'P')]=cm[(g,'A')]
    p1,p2=S[hr][lc-1],S[hr+1][lc-1]
    cname=f"Clean - {p1}" if p1==p2 else f"Clean - YTD {p1}-{p2}"
    if cname not in OUT.sheetnames: fails.append('missing '+cname); continue
    O=[list(r) for r in OUT[cname].iter_rows(values_only=True)]
    # source leaves per brand (by label inside the brand page)
    def brand_rows(b):
        start=[i for i,r in enumerate(S) if G(S,i,lc)==f'{b} - Brand Total'][0]
        d={}
        for i in range(start,len(S)):
            l=G(S,i,lc)
            if i>start and isinstance(l,str) and l.startswith('From Date'): break
            if isinstance(l,str) and l not in d: d[l]=i
        return d
    leafvals={}
    for b in BRANDS:
        d=brand_rows(b)
        for k,l in enumerate(LEAVES):
            r=d.get(l)
            v={}
            for key in ('A','P','T'):
                v['CON'+key]=N(G(S,r,cm[('CON',key)])) if r is not None else 0
                v['LEA'+key]=N(G(S,r,cm[('LEA',key)])) if r is not None else 0
                v['SAL'+key]=N(G(S,r,cm[('SAL',key)])) if r is not None else 0
            def w(rate,vol):
                x=G(S,r,cm.get(rate)) if r is not None else None
                return (v[vol],x*v[vol]) if num(x) else (0,0)
            v['EC'],v['CN']=w(('CONV','A'),'CONA'); v['ECT'],v['CNT']=w(('CONV','T'),'CONT')
            v['EE'],v['EN']=w(('GEFF','A'),'CONA'); v['EET'],v['ENT']=w(('GEFF','T'),'CONT')
            leafvals[(b,k)]=v
    def expected(brands,lv):
        e={}
        for key in leafvals[(BRANDS[0],0)]:
            e[key]=sum(leafvals[(b,k)][key] for b in brands for k in lv)
        x={}
        for g in ('CON','LEA','SAL'):
            x[g+'A'],x[g+'P'],x[g+'T']=e[g+'A'],e[g+'P'],e[g+'T']
            x[g+'D']=e[g+'P']-e[g+'T']; x[g+'V']=(e[g+'P']/e[g+'T']-1) if e[g+'T'] else None
        x['CNVA']=e['CN']/e['EC'] if e['EC'] else None; x['CNVT']=e['CNT']/e['ECT'] if e['ECT'] else None
        x['CLOA']=e['SALA']/e['LEAA'] if e['LEAA'] else None; x['CLOT']=e['SALT']/e['LEAT'] if e['LEAT'] else None
        x['EFFA']=e['EN']/e['EE'] if e['EE'] else None; x['EFFT']=e['ENT']/e['EET'] if e['EET'] else None
        return x
    # locate tables in the output by their title, map columns from header rows
    titles={'4 BRANDS TOTAL':BRANDS}
    for b in BRANDS: titles[b.upper()]=[b]
    seen=set()
    for i,row in enumerate(O):
        t=row[0] if row else None
        if not isinstance(t,str): continue
        key=next((k for k in titles if (t==k or (k=='4 BRANDS TOTAL' and t.startswith(k)))),None)
        if not key or key in seen: continue
        seen.add(key)
        grp=None; col={}
        for c in range(1,len(O[i+1])):
            if O[i+1][c]: grp=str(O[i+1][c])
            m=O[i+2][c]
            if grp and m:
                gk={'Contacts':'CON','Leads':'LEA','Gross Sales':'SAL','Conversion':'CNV','Closing':'CLO','Effectiveness':'EFF'}.get(grp)
                mk={'Actual':'A','Projecting':'P','Target':'T','Var vs Tar':'D','Var %':'V','Var (pp)':'PP'}.get(m)
                if gk and mk: col[gk+mk]=c
        for j,(lab,lv) in enumerate(ROWS):
            r=i+3+j
            if (O[r][0] or '').strip()!=lab: fails.append(f'{cname} {key}: row {r+1} label {O[r][0]!r} != {lab!r}'); continue
            x=expected(titles[key],lv)
            for k2,c in col.items():
                if k2.endswith('PP'):
                    a,t=x.get(k2[:3]+'A'),x.get(k2[:3]+'T'); e=(a-t)*100 if a is not None and t is not None else None
                else: e=x.get(k2)
                if not close(e,O[r][c]): fails.append(f'{cname} {key} [{lab}] {k2}: got {O[r][c]} expected {e}')
    print(cname,'checked; fails so far',len(fails))
# ---------------- source views ----------------
SM=[list(r) for r in SRC['Summary'].iter_rows(values_only=True)]
per=[r for r in range(12) if isinstance(SM[r][2],str) and SM[r][2].startswith('Period')][0]
lab=lambda r: SM[r][2].strip() if r<len(SM) and isinstance(SM[r][2],str) else ''
INC=['Hippo Branded','Google','Social Media','Internet Organic','Internet Direct','Display Banner','DM - Internal','Online Quoting - PV','Online Quoting - PNV']
INC_T=[x for x in INC if not x.startswith('Online Quoting')]
blocks={}
for r in range(per+2,len(SM)-1):
    if lab(r) and lab(r+1) in ('Incoming Contacts','Answered Calls'):
        h=lab(r); kind=lab(r+1)
        if kind=='Answered Calls': srcn='JHB Inbound (Connex)'; br=h
        elif ' - ' in h and h.rsplit(' - ',1)[1] in BRANDS+['Virseker']: srcn,br=h.rsplit(' - ',1)
        else: continue
        rows={}
        for k in range(r+1,min(r+27,len(SM))):
            if k>r+1 and lab(k+1) in ('Incoming Contacts','Answered Calls'): break
            rows.setdefault(lab(k),k)
        m={'Incoming':rows.get('Incoming Contacts') or rows.get('Answered Calls'),'Processed':rows.get('Processed Contacts') or rows.get('N ACD Calls'),
           'Leads':rows.get('Leads'),'Gross Sales':rows.get('Gross Sales') or rows.get('Sales')}
        blocks.setdefault((srcn,br),m)
months=[c for c,v in enumerate(SM[per]) if num(v) and 200001<=v<=209912]
cur=max(months,key=lambda c:SM[per][c]); prv=sorted(months,key=lambda c:SM[per][c])[-2]
cap=per-1
d0=[c for c,v in enumerate(SM[cap]) if isinstance(v,str) and 'Daily' in v][0]; y0=[c for c,v in enumerate(SM[cap]) if isinstance(v,str) and 'Yesterday' in v][0]
days=[c for c in range(d0,y0) if SM[per+1][c] in ('Mon','Tue','Wed','Thu','Fri','Sat','Sun')]
def tot(srcs,brands,metric,cols):
    s=0
    for sn in srcs:
        for b in brands:
            m=blocks.get((sn,b))
            if m and m.get(metric): s+=sum(N(SM[m[metric]][c]) for c in cols)
    return s
X=[list(r) for r in OUT['Source x Brand'].iter_rows(values_only=True)]
sec=None
for row in X:
    if isinstance(row[0],str) and row[0].upper()==row[0] and row[1] is None and row[0].strip():
        sec={'4 BRANDS (EXCL VIRSEKER)':BRANDS}.get(row[0], [b for b in BRANDS if b.upper()==row[0]] or None)
        continue
    if sec and isinstance(row[0],str) and (row[0] in INC or row[0] in ('All dialler sources','JHB Inbound (Connex)')):
        srcs=INC_T if row[0]=='All dialler sources' else [row[0]]
        if row[0].startswith('Online Quoting'):
            for e,g in zip([tot(srcs,sec,'Incoming',[cur]),tot(srcs,sec,'Processed',[cur])],row[1:3]):
                if not close(e,g): fails.append(f'SxB OQ {sec} {row[0]}: got {g} expected {e}')
            if row[3]!='n/a': fails.append(f'SxB OQ leads should be n/a: {row[3]}')
            continue
        p,l,s=tot(srcs,sec,'Processed',[cur]),tot(srcs,sec,'Leads',[cur]),tot(srcs,sec,'Gross Sales',[cur])
        for e,g in zip([tot(srcs,sec,'Incoming',[cur]),p,l,s,l/p if p else None,s/l if l else None,s/p if p else None],row[1:8]):
            if not close(e,g): fails.append(f'SxB {sec} {row[0]}: got {g} expected {e}')
MV=[list(r) for r in OUT['Monday View'].iter_rows(values_only=True)]
sec=None
for row in MV:
    if isinstance(row[0],str) and row[0].upper()==row[0] and row[1] is None and row[0].strip():
        sec={'4 BRANDS (EXCL VIRSEKER)':BRANDS}.get(row[0], [b for b in BRANDS if b.upper()==row[0]] or None); continue
    if sec and isinstance(row[0],str) and (row[0] in INC or row[0] in ('All dialler sources','JHB Inbound (Connex)')):
        srcs=INC_T if row[0]=='All dialler sources' else [row[0]]
        if row[0].startswith('Online Quoting'): continue
        for gi,dn in enumerate([('Mon',),('Tue','Wed','Thu'),('Fri',),('Sat',)]):
            cols=[c for c in days if SM[per+1][c] in dn]
            p,l,s=tot(srcs,sec,'Processed',cols),tot(srcs,sec,'Leads',cols),tot(srcs,sec,'Gross Sales',cols)
            for e,g in zip([p/len(cols),l/p if p else None,s/l if l else None,s/p if p else None],row[1+gi*4:5+gi*4]):
                if not close(e,g): fails.append(f'Monday {sec} {row[0]} {dn[0]}: got {g} expected {e}')
EX=[list(r) for r in OUT['Source Explorer'].iter_rows(values_only=True)]
rowP=[i for i,r in enumerate(EX) if r[0]=='Processed'][0]
dcols_out=[c for c in range(1,len(EX[rowP-2])) if EX[rowP-2][c] in ('Mon','Tue','Wed','Thu','Fri','Sat','Sun')]
for j,c in enumerate(dcols_out):
    e=tot(INC_T,BRANDS,'Processed',[days[j]])
    if not close(e,EX[rowP][c]): fails.append(f'Explorer day {j}: got {EX[rowP][c]} expected {e}')
DB=[list(r) for r in OUT['Day by Day'].iter_rows(values_only=True)]
sec=None
for i,row in enumerate(DB):
    if isinstance(row[0],str) and row[0].upper()==row[0] and row[1] is None and row[0].strip():
        sec={'4 BRANDS (EXCL VIRSEKER)':BRANDS}.get(row[0], [b for b in BRANDS if b.upper()==row[0]] or None); continue
    if sec and row[0] in ('Processed','Leads','Gross Sales'):
        for j,c in enumerate(days):
            e=tot(INC_T,sec,row[0],[c])
            if not close(e,row[1+j]): fails.append(f'DayByDay {sec} {row[0]} day {j}: got {row[1+j]} expected {e}')
print('TOTAL values checked:',checked,'| FAILURES:',len(fails))
for f in fails[:30]: print('  ',f)
