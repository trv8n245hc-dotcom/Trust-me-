"""Independent verification of the UPdateReport output.
Usage: python verify.py <source.xlsx> <report_output.xlsx>
Everything is located by header text / labels (never fixed positions), then
the replica, Brand x Source and Monday View values are recomputed and compared."""
import sys, openpyxl, statistics as st
from openpyxl.utils import get_column_letter as L
src_p, out_p = sys.argv[1], sys.argv[2]
SRC = openpyxl.load_workbook(src_p, read_only=True, data_only=True)
OUT = openpyxl.load_workbook(out_p, read_only=True, data_only=True)
num = lambda v: isinstance(v,(int,float)) and not isinstance(v,bool)
def grid(wb,n): return [list(r) for r in wb[n].iter_rows(values_only=True)]
def N(v): return v if num(v) else 0.0
GROUPS = [('cxd','CXD'),('closing','CLOSE'),('gross effectiveness','GEFF'),('net effectiveness','NEFF'),('conversion','CONV'),
          ('contacts','CONTACTS'),('leads','LEADS'),('net sales','NSALES'),('gross sales','GSALES')]
MEAS = {'actual':'ACT','projecting':'PROJ','target':'TAR','tar dif':'DIF','tar var':'VAR','w target':'WT','w tar var':'WVAR'}
def gidx(t):
    t=(t or '').strip().lower() if isinstance(t,str) else ''
    for k,v in GROUPS:
        if k in t: return v
fails=[]; checked=0
# ---------------- Exec replicas ----------------
for sname in SRC.sheetnames:
    S=grid(SRC,sname)
    hdr=None
    for r in range(10):
        for c in range(8):
            if r<len(S) and c<len(S[r]) and isinstance(S[r][c],str) and S[r][c].startswith('From Date'): hdr=(r,c); break
        if hdr: break
    if not hdr: continue
    hr,lc=hdr
    p1,p2=S[hr][lc-1],S[hr+1][lc-1]
    rep = f"Replica - {p1}" if p1==p2 else f"Replica - YTD {p1}-{p2}"
    if rep not in OUT.sheetnames: fails.append(f'missing {rep}'); continue
    R=grid(OUT,rep)
    # source column map
    cmap={}; cur=None
    for c in range(lc+1,len(S[hr])):
        g=gidx(S[hr][c]); cur=g or cur
        m=MEAS.get((S[hr+1][c] or '').strip().lower()) if isinstance(S[hr+1][c],str) else None
        if cur and m and (cur,m) not in cmap: cmap[(cur,m)]=c
    # replica column map from replica header rows 4/5 (strip ' *')
    rmap={}; cur=None
    for c in range(3,len(R[3])):
        g=gidx(R[3][c]) if R[3][c] else None; cur=g or cur
        m=MEAS.get(str(R[4][c]).replace(' *','').strip().lower()) if R[4][c] else None
        if cur and m: rmap[(cur,m)]=c
    assert set(cmap)==set(rmap), (rep, set(cmap)^set(rmap))
    def sv(r,g,m): return S[r][cmap[(g,m)]] if (g,m) in cmap and cmap[(g,m)]<len(S[r]) else None
    def rv(r,g,m): return R[r+2][rmap[(g,m)]] if (g,m) in rmap else None
    nmis_calc=0
    for r in range(hr+2,len(S)):
        lab=S[r][lc] if lc<len(S[r]) else None
        if not isinstance(lab,str) or not lab.strip() or lab.startswith(('From Date','To Date')): continue
        # replica label in col C must equal source label
        if (R[r+2][2] or '').strip()!=lab.strip(): fails.append(f'{rep} row {r+3}: label {R[r+2][2]!r} != {lab!r}'); continue
        for (g,m) in cmap:
            s_=sv(r,g,m); x=rv(r,g,m); checked+=1
            # independent recomputation
            if m in ('ACT','PROJ','TAR','WT') and not ((g=='CLOSE' or g=='CXD') and m in ('ACT','TAR')):
                exp = s_ if num(s_) else None
            elif m=='DIF':
                p,t=sv(r,g,'PROJ'),sv(r,g,'TAR'); exp = p-t if num(p) and num(t) else None
            elif m=='VAR' and g in ('CONTACTS','LEADS','GSALES','NSALES'):
                p,t=sv(r,g,'PROJ'),sv(r,g,'TAR'); exp = (p-t)/t if num(p) and num(t) and t!=0 else None
            elif g=='CLOSE' and m in ('ACT','TAR'):
                a,b=sv(r,'GSALES',m),sv(r,'LEADS',m); exp = N(a)/N(b) if N(b)!=0 else None
            elif g=='CXD' and m in ('ACT','TAR'):
                a,b=sv(r,'NSALES',m),sv(r,'GSALES',m); exp = 1-N(a)/N(b) if N(b)!=0 else None
            elif m in ('VAR','WVAR'):
                # actual of recomputed group (closing/cxd recomputed, others linked)
                def act(mm):
                    if g=='CLOSE': a,b=sv(r,'GSALES',mm),sv(r,'LEADS',mm); return N(a)/N(b) if N(b)!=0 else None
                    if g=='CXD': a,b=sv(r,'NSALES',mm),sv(r,'GSALES',mm); return 1-N(a)/N(b) if N(b)!=0 else None
                    v=sv(r,g,mm); return v if num(v) else None
                a=act('ACT'); t=act('TAR') if m=='VAR' else (sv(r,g,'WT') if num(sv(r,g,'WT')) else None)
                exp = a/t-1 if a is not None and t not in (None,0) else None
            else:
                exp = s_ if num(s_) else None
            ok = (exp is None and not num(x)) or (exp is not None and num(x) and abs(x-exp) <= 1e-6*max(1,abs(exp)))
            if not ok: fails.append(f'{rep} {L(rmap[(g,m)]+1)}{r+3} [{lab}] {g}/{m}: replica {x!r} expected {exp!r}')
    print(f'{rep}: replica cells checked OK so far; fails={len(fails)}')
# ---------------- Summary-based sheets ----------------
SM=grid(SRC,'Summary')
per=[r for r in range(12) if isinstance(SM[r][2],str) and SM[r][2].startswith('Period')][0]
months=[(c,v) for c,v in enumerate(SM[per]) if num(v) and 200001<=v<=209912]
cur=max(months,key=lambda x:x[1])[0]; prv=sorted(months,key=lambda x:x[1])[-2][0]
cap=per-1
dailyStart=[c for c,v in enumerate(SM[cap]) if isinstance(v,str) and 'Daily' in v][0]
yest=[c for c,v in enumerate(SM[cap]) if isinstance(v,str) and 'Yesterday' in v][0]
dcols=[c for c in range(dailyStart,yest) if SM[per+1][c] in ('Mon','Tue','Wed','Thu','Fri','Sat','Sun')]
lab=lambda r: SM[r][2].strip() if r<len(SM) and isinstance(SM[r][2],str) else ''
blocks=[]
for r in range(per+2,len(SM)-1):
    if lab(r) and lab(r+1) in ('Incoming Contacts','Answered Calls'):
        kind='d' if lab(r+1)=='Incoming Contacts' else 'i'
        rows={}
        for k in range(r+1,min(r+27,len(SM))):
            if k>r+1 and lab(k+1) in ('Incoming Contacts','Answered Calls'): break
            rows.setdefault(lab(k),k)
        P=rows.get('Processed Contacts' if kind=='d' else 'N ACD Calls'); Ld=rows.get('Leads'); Sl=rows.get('Gross Sales') or rows.get('Sales')
        if P and Ld and Sl: blocks.append((r,P,Ld,Sl,rows.get('Attempts') if kind=='d' else None))
BX=grid(OUT,'Brand x Source'); MV=grid(OUT,'Monday View')
def close(a,b,tol=1e-6): return (a is None and not num(b)) or (a is not None and num(b) and abs(a-b)<=tol*max(1,abs(a)))
nb=0
for row in BX:
    if num(row[1]) and isinstance(row[0],str):
        h=int(row[1])-1; blk=[b for b in blocks if b[0]==h]
        if not blk: fails.append(f'BxS heading row {h+1} not a block'); continue
        _,P,Ld,Sl,_a=blk[0]
        p,l,s=N(SM[P][cur]),N(SM[Ld][cur]),N(SM[Sl][cur])
        exp=[p,l,s,(l/p if p else None),(s/l if l else None),(s/p if p else None)]
        got=row[3:9]
        for e,g in zip(exp,got):
            checked+=1
            if not close(e,g): fails.append(f'BxS row {h+1} {row[0]}: got {g} expected {e}')
        nb+=1
def dsum(r,days): return sum(N(SM[r][c]) for c in dcols if SM[per+1][c] in days)
def dcnt(days): return sum(1 for c in dcols if SM[per+1][c] in days)
nm=0
for row in MV:
    if num(row[1]) and isinstance(row[0],str):
        h=int(row[1])-1; _,P,Ld,Sl,At=[b for b in blocks if b[0]==h][0]
        for i,days in enumerate([('Mon',),('Tue','Wed','Thu'),('Fri',),('Sat',)]):
            p,l,s=dsum(P,days),dsum(Ld,days),dsum(Sl,days)
            exp=[p/dcnt(days), (l/p if p else None),(s/l if l else None),(s/p if p else None)]
            for e,g in zip(exp,row[2+4*i:6+4*i]):
                checked+=1
                if not close(e,g): fails.append(f'MonView row {h+1} {row[0]} {days[0]}: got {g} expected {e}')
        nm+=1
print('Brand x Source rows verified:',nb,'| Monday View rows verified:',nm)
print('TOTAL values checked:',checked,'| FAILURES:',len(fails))
for f in fails[:25]: print('  ',f)
