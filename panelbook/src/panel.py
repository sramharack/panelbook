"""One panel sheet: company-look schedule (print area) + input zones + hidden calc zone."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from common import *


def set_cell(ws, ref, value=None, f=None, al=None, fmt=None, fl=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    if f is not None:
        c.font = f
    if al is not None:
        c.alignment = al
    if fmt is not None:
        c.number_format = fmt
    if fl is not None:
        c.fill = fl
    if border is not None:
        c.border = border
    return c


def border_range(ws, rng, top=None, bottom=None, left=None, right=None, all_=None):
    """Apply sides to cells in rng (outer edges only for top/bottom/left/right)."""
    from openpyxl.utils.cell import range_boundaries
    c1, r1, c2, r2 = range_boundaries(rng)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            b = cell.border
            nb = Border(left=b.left, right=b.right, top=b.top, bottom=b.bottom)
            if all_ is not None:
                nb = Border(left=all_, right=all_, top=all_, bottom=all_)
            if top is not None and r == r1:
                nb.top = top
            if bottom is not None and r == r2:
                nb.bottom = bottom
            if left is not None and c == c1:
                nb.left = left
            if right is not None and c == c2:
                nb.right = right
            cell.border = nb


def merge(ws, rng):
    ws.merge_cells(rng)


def allow(i):
    """Allowable ampacity of conductor index i (one set): insulation-column ampacity x ambient x bundling
    factors (NEC 310.15(B)/(C)), capped at the lower of the termination and insulation columns (110.14(C))
    and at the 240.4(D) small-conductor limit."""
    return (f"MIN(INDEX(LK_Amp,{i},{pv('icol')})*{pv('fct')},INDEX(LK_Amp,{i},{pv('ccol')}),"
            f"IF(N(INDEX(LK_Lim,{i},{pv('lcol')}))>0,INDEX(LK_Lim,{i},{pv('lcol')}),1E+9))")


def idx_derated(amps):
    """Smallest index whose derated insulation-column ampacity carries `amps` (per set)."""
    return f'COUNTIF(INDEX(LK_Amp,0,{pv("icol")}),"<"&{amps}/MAX({pv("fct")},1E-6))+1'


def idx_cap(amps):
    """Smallest index whose cap-column (termination/insulation) ampacity carries `amps` (per set)."""
    return f'COUNTIF(INDEX(LK_Amp,0,{pv("ccol")}),"<"&{amps})+1'


def idx_ocpd(tc, t, ns):
    """Test (c): conductor protected by the breaker, plus 240.4(D). tc = trip, or the 240.4(B) threshold."""
    return (f'MAX({idx_derated(f"{tc}/{ns}")},{idx_cap(f"{tc}/{ns}")},'
            f'COUNTIF(INDEX(LK_Lim,0,{pv("lcol")}),"<"&{t}/{ns})+1)')


def pick_index(t, ns):
    """Smallest conductor index protected by a breaker of rating t over ns sets (used for the panel feeder)."""
    return f"MIN(ROWS(LK_WSize),MAX({pv('minw')},{idx_ocpd(t, t, ns)},IF({ns}>1,{pv('i10')},1)))"


def egc_index(t, wi):
    return (f"MIN({wi},IFERROR(MATCH(INDEX(LK_Egc,MIN(ROWS(LK_EgcA),COUNTIF(LK_EgcA,\"<\"&{t})+1),"
            f"{pv('lcol')}),LK_WSize,0),{wi}))")


def status_formula(Z, H, r):
    """Per-circuit status text (first matching rule wins)."""
    c = f"${Z['ckt']}{r}"
    has, cont, Pp, v = f"{H['has']}{r}", f"{H['cont']}{r}", f"{H['P']}{r}", f"{H['valid']}{r}"
    s, sp, va, own = f"{H['sub']}{r}", f"{H['sPol']}{r}", f"{H['VA']}{r}", f"{H['own']}{r}"
    amp, trip, cls = f"{H['amp']}{r}", f"${Z['trip']}{r}", f"${Z['cls']}{r}"
    xf, phd, rph = f"{H['xfm']}{r}", f"${Z['ph']}{r}", f"${ROWPH}{r}"
    oi, ns, wi = f"{H['oi']}{r}", f"{H['ns']}{r}", f"{H['wi']}{r}"
    fx = f"${Z['q1']}{r}:${Z['add']}{r}"
    slots, ph, wr, df, lim = pv("slots"), pv("ph"), pv("wires"), pv("def"), pv("lim")
    hl, vhl, vll, vln = pv("hl"), pv("vhl"), pv("vll"), pv("vln")
    rules = [
        (f'AND({cont}<>"",{has}=1)', f'"ERR: slot taken by CKT "&{cont}'),
        (f'{has}=0', f'IF({cont}<>"","(pole of CKT "&{cont}&")","")'),
        (f'{c}>{slots}', f'"ERR: beyond panel size ("&{slots}&" ckts)"'),
        (f'{c}+2*({Pp}-1)>{slots}', f'"ERR: "&{Pp}&"P breaker runs past last slot"'),
        (f'AND({Pp}=3,{ph}<>3)', '"ERR: 3P breaker on 1-phase panel"'),
        (f'AND({Pp}>=2,{ph}=1,{wr}<3)', '"ERR: multi-pole breaker on 2-wire panel"'),
        (f'AND({Pp}=1,{ph}=3,{wr}=3,{va}>0)', '"ERR: 1P load on 3-wire system (no neutral)"'),
        (f'AND({hl}<>"",{rph}={hl},{Pp}=1,{va}>0)',
         f'"ERR: 1P load on high leg "&{hl}&" ("&TEXT({vhl},"0")&" V to N)"'),
        (f'{v}=-1', '"ERR: a panel cannot feed itself"'),
        (f'{v}=-2', f'"ERR: no tab named \'"&{s}&"\'"'),
        (f'AND({v}=1,{xf}=0,{hl}<>"",{H["sHL"]}{r}="",{H["sN"]}{r}=1,ISNUMBER(FIND({hl},{phd})))',
         f'"ERR: "&{s}&" neutral loads would see "&TEXT({vhl},"0")&" V (feeder on high leg)"'),
        (f'AND({v}=1,{xf}=0,OR({H["sVLL"]}{r}<>{vll},{H["sVLN"]}{r}<>{vln}))',
         f'"WARN: "&{s}&" is "&{H["sVLN"]}{r}&"/"&{H["sVLL"]}{r}&" V: mark XFMR or fix"'),
        (f'AND({v}=1,{xf}=0,{Pp}<>{sp})', f'"WARN: "&{s}&" needs a "&{sp}&"P feeder"'),
        (f'AND({v}=1,COUNT({fx})>0)', '"WARN: fixture/added VA ignored (sub-panel)"'),
        (f'AND({xf}=1,{v}<>1)', '"WARN: XFMR only applies to a sub-panel feeder"'),
        (f'AND({va}>0,NOT(ISNUMBER({trip})))', '"WARN: no trip rating"'),
        (f'AND({va}>0,ISNUMBER({trip}),{H["Ia"]}{r}>N({trip})+0.001)',
         f'IF({H["is100"]}{r}=1,"WARN: load exceeds 100%-rated breaker",'
         f'"WARN: breaker below 125% of continuous load (210.20(A)/215.3)")'),
        (f'AND({va}>0,ISNUMBER({trip}),{amp}>IF({H["is100"]}{r}=1,1,{lim})*N({trip}))',
         f'"WARN: "&TEXT({amp}/{trip},"0%")&" of trip"'),
        (f'AND({hl}<>"",{rph}={hl},{Pp}=1,{has}=1)',
         '"WARN: 1P breaker on high leg - no 120 V loads here"'),
        (f'AND({hl}<>"",{Pp}=2,ISNUMBER(FIND({hl},{phd})),OR({va}>0,{v}=1),'
         f'NOT(ISNUMBER(SEARCH("STRAIGHT",${Z["brk"]}{r}&""))))',
         '"WARN: 2P on high leg: straight 240 V breaker, not 120/240 (NEC 240.85); 1-ph 240 V loads better on A-C"'),
        (f'N({oi})=-1', '"WARN: wire override not in Lookups size list"'),
        (f'AND(N({oi})>0,{va}>0,{oi}<MAX({H["xa"]}{r},{H["xb"]}{r},{H["xc"]}{r}))',
         f'"WARN: wire override fails "&IF({oi}<{H["xc"]}{r},"OCPD (240.4)",IF({oi}<{H["xa"]}{r},'
         f'"125% continuous (210.19/215.2)","derated load (310.15)"))'),
        (f'AND(N({oi})>0,N({oi})<{pv("minw")})', '"WARN: wire override below minimum "&LK_MinWire'),
        (f'AND({ns}>1,{wi}>0,{wi}<{pv("i10")})', '"WARN: parallel sets need 1/0 min (310.10(G))"'),
        (f'AND({own}>0,TRIM({cls})="")', f'"WARN: no load class (counted as "&{df}&")"'),
        (f'AND({v}=1,ISNUMBER(${Z["ovr"]}{r}))', '"WARN: VA override ignored on a sub-panel feeder (edit the sub-panel)"'),
        (f'{H["ov"]}{r}=1', f'"WARN: VA override (calculated "&TEXT({H["own0"]}{r},"#,##0")&" VA)"'),
        (f'AND({v}=1,ISNA(MATCH({s},LK_PanelNames,0)))', f'"WARN: add "&{s}&" to Panels sheet"'),
    ]
    out = '"OK"'
    for cond, res in reversed(rules):
        out = f"IF({cond},{res},{out})"
    return "=" + out


def helper_formulas(ws, Z, H, side, L):
    R0_, R1 = L.R0, L.R1
    kP = f"CHOOSE({{P}},1/{pv('vln')},2/{pv('vll')},SQRT(3)/{pv('vll')})"
    for r in range(R0_, R1 + 1):
        h = {k: f"{col}{r}" for k, col in H.items()}
        sub_in = f"${Z['sub']}{r}"
        ws[h["sub"]] = f"=TRIM({sub_in})"
        s, v = h["sub"], h["valid"]
        ws[v] = (f'=IF({s}="",0,IF({s}={pv("name")},-1,'
                 f'IF(ISERROR(INDIRECT("\'"&{s}&"\'!${PV}$1")),-2,1)))')
        for key, pk in [("sVA", "conn"), ("sA", "phA"), ("sB", "phB"), ("sC", "phC"), ("sPol", "hot"),
                        ("sDem", "dem"), ("sIA", "aA"), ("sIB", "aB"), ("sIC", "aC"), ("sVLL", "vll"),
                        ("sVLN", "vln"), ("sN", "neut")]:
            ws[h[key]] = f'=IF({v}=1,N(INDIRECT("\'"&{s}&"\'!${PV}${P[pk]}")),0)'
        ws[h["sHL"]] = f'=IF({v}=1,T(INDIRECT("\'"&{s}&"\'!${PV}${P["hl"]}")),"")'
        ws[h["xfm"]] = f'=IF(LEN(TRIM(${Z["xf"]}{r}&""))>0,1,0)'
        q1, v1, q2, v2, ad = (f"${Z[k]}{r}" for k in ["q1", "v1", "q2", "v2", "add"])
        ovr = f"${Z['ovr']}{r}"
        ws[h["own0"]] = f"=IF({v}=1,0,N({q1})*N({v1})+N({q2})*N({v2})+N({ad}))"
        ws[h["ov"]] = f"=IF(AND({v}<>1,ISNUMBER({ovr})),1,0)"
        ws[h["own"]] = f"=IF({h['ov']}=1,MAX(0,{ovr}),{h['own0']})"
        ws[h["VA"]] = f"=IF({v}=1,{h['sVA']},{h['own']})"
        pol = f"${Z['pol']}{r}"
        ws[h["P"]] = f"=IF(ISNUMBER({pol}),MAX(1,MIN(3,{pol})),IF({v}=1,MAX(1,{h['sPol']}),1))"
        Pp, VA, sp, xf = h["P"], h["VA"], h["sPol"], h["xfm"]
        match = f"AND({v}=1,{Pp}={sp},{xf}=0)"
        ws[h["sh1"]] = f"=IF({Pp}=1,{VA},IF({match},{h['sA']},{VA}/{Pp}))"
        ws[h["sh2"]] = f"=IF({Pp}<2,0,IF({match},{h['sB']},{VA}/{Pp}))"
        ws[h["sh3"]] = f"=IF({Pp}<3,0,IF({match},{h['sC']},{VA}/{Pp}))"
        k = kP.replace("{P}", Pp)
        # line current per pole: direct sub-panel = its own line currents; otherwise VA/pole x k(P)
        ws[h["c1"]] = f"=IF({match},{h['sIA']},{h['sh1']}*{k})"
        ws[h["c2"]] = f"=IF({Pp}<2,0,IF({match},{h['sIB']},{h['sh2']}*{k}))"
        ws[h["c3"]] = f"=IF({Pp}<3,0,IF({match},{h['sIC']},{h['sh3']}*{k}))"
        d, cl, tr = f"${Z['desc']}{r}", f"${Z['cls']}{r}", f"${Z['trip']}{r}"
        ws[h["has"]] = (f"=IF(OR(LEN({d})>0,LEN({cl})>0,COUNT({q1}:{ad})>0,LEN({sub_in})>0,"
                        f"ISNUMBER({pol}),ISNUMBER({tr}),ISNUMBER(${Z['ovr']}{r})),1,0)")
        ck = Z["ckt"]
        Hc = lambda key, rr: f"{H[key]}{rr}"
        if r == R0_:
            ws[h["cont"]] = '=""'
            ws[h["slot"]] = f"={h['sh1']}"
            ws[h["sla"]] = f"={h['c1']}"
        elif r == R0_ + 1:
            ws[h["cont"]] = f'=IF({Hc("P", r-1)}>=2,${ck}{r-1},"")'
            ws[h["slot"]] = f"={h['sh1']}+IF({Hc('P', r-1)}>=2,{Hc('sh2', r-1)},0)"
            ws[h["sla"]] = f"={h['c1']}+IF({Hc('P', r-1)}>=2,{Hc('c2', r-1)},0)"
        else:
            ws[h["cont"]] = (f'=IF({Hc("P", r-1)}>=2,${ck}{r-1},'
                             f'IF({Hc("P", r-2)}>=3,${ck}{r-2},""))')
            ws[h["slot"]] = (f"={h['sh1']}+IF({Hc('P', r-1)}>=2,{Hc('sh2', r-1)},0)"
                             f"+IF({Hc('P', r-2)}>=3,{Hc('sh3', r-2)},0)")
            ws[h["sla"]] = (f"={h['c1']}+IF({Hc('P', r-1)}>=2,{Hc('c2', r-1)},0)"
                            f"+IF({Hc('P', r-2)}>=3,{Hc('c3', r-2)},0)")
        ws[h["amp"]] = (f"=IF({VA}=0,0,MAX({h['c1']},{h['c2']},{h['c3']})"
                        f"*IF(AND({v}=1,{h['sVA']}>0),{h['sDem']}/{h['sVA']},1))")
        ws[h["key"]] = f'=IF({h["own"]}>0,IF(TRIM({cl})="",{pv("def")},TRIM({cl})),"")'
        if side == "odd":
            ws[h["orank"]] = f'=IF({h["ov"]}=1,COUNTIF({H["ov"]}${R0_}:{H["ov"]}{r},1),"")'
        else:
            ws[h["orank"]] = (f'=IF({h["ov"]}=1,COUNTIF({OH["ov"]}${R0_}:{OH["ov"]}${R1},1)'
                              f'+COUNTIF({H["ov"]}${R0_}:{H["ov"]}{r},1),"")')
        # schedule text row: 0 = this row's breaker, 1 = the breaker one row up (middle row of a 3P), -1 = none
        if r == R0_:
            ws[h["dsrc"]] = f'=IF(AND({h["cont"]}="",{h["P"]}<3),0,-1)'
        else:
            ws[h["dsrc"]] = (f'=IF(AND({h["cont"]}="",{h["P"]}<3),0,'
                             f'IF(AND({H["cont"]}{r-1}="",{H["P"]}{r-1}=3),1,-1))')
        if side == "odd":
            ws[h["frank"]] = f'=IF({v}=1,COUNTIF({H["valid"]}${R0_}:{H["valid"]}{r},1),"")'
        else:
            ws[h["frank"]] = (f'=IF({v}=1,COUNTIF({OH["valid"]}${R0_}:{OH["valid"]}${R1},1)'
                              f'+COUNTIF({H["valid"]}${R0_}:{H["valid"]}{r},1),"")')
        # ---- cable selection: smallest size passing all four NEC tests
        #   (a) cap-column ampacity >= 1.25 x continuous + noncontinuous      210.19(A)(1)(a) / 215.2(A)(1)(a)
        #   (b) derated ampacity >= load                                       210.19(A)(1)(b) / 215.2(A)(1)(b)
        #   (c) derated ampacity >= breaker trip (240.4(B) optional) + 240.4(D)
        t, oi, oa, ns, wi, ei = h["t"], h["oi"], h["oa"], h["ns"], h["wi"], h["ei"]
        cf, Ia, tc, xa, xb, xc = h["cf"], h["Ia"], h["tc"], h["xa"], h["xb"], h["xc"]
        Lr = h["amp"]
        wov = f"${Z['wov']}{r}"
        ws[t] = f"=IF(ISNUMBER({tr}),{tr},0)"
        ws[cf] = (f'=IF({v}=1,N(INDIRECT("\'"&{s}&"\'!${PV}${P["contf"]}")),'
                  f'IF({h["key"]}="",0,IF(IFERROR(INDEX(LK_Cont,MATCH({h["key"]},LK_ClassName,0)),"No")="Yes",1,0)))')
        ws[h["is100"]] = (f'=IF(OR({pv("brkmode")}=2,LEFT(UPPER(TRIM(${Z["brk"]}{r}&"")),4)="100%"),1,0)')
        # 210.19(A)(1)/215.2(A)(1) and 210.20(A)/215.3: 125% of continuous load, 100% if the breaker is 100% rated
        ws[Ia] = f"={Lr}*(IF({h['is100']}=1,1,1.25)*{cf}+(1-{cf}))"
        ws[tc] = (f'=IF(AND({pv("b2404")}=1,{t}>0,{t}<=800,{h["key"]}<>"Receptacle"),'
                  f'IF(COUNTIF(LK_Breakers,"<"&{t})>0,INDEX(LK_Breakers,COUNTIF(LK_Breakers,"<"&{t}))+0.001,{t}),{t})')
        mt = f'MATCH(TRIM({wov}&""),LK_WSize,0)'
        ws[oi] = f'=IF(TRIM({wov}&"")="","",IFERROR(IF({allow(mt)}>0,{mt},-1),-1))'
        ws[oa] = f"=IF(N({oi})>0,{allow(oi)},0)"
        ws[ns] = (f"=IF({t}=0,0,IF({oa}>0,IF({oi}<{pv('i10')},1,MAX(1,ROUNDUP({tc}/{oa},0))),"
                  f"MAX(1,ROUNDUP({tc}/{pv('maxamp')},0),ROUNDUP({Ia}/{pv('maxterm')},0),"
                  f"ROUNDUP({Lr}/{pv('maxamp')},0))))")
        ws[xa] = f"=IF({t}=0,0,{idx_cap(f'{Ia}/{ns}')})"
        ws[xb] = f"=IF({t}=0,0,{idx_derated(f'{Lr}/{ns}')})"
        ws[xc] = f"=IF({t}=0,0,{idx_ocpd(tc, t, ns)})"
        ws[wi] = (f"=IF({t}=0,0,IF({oa}>0,{oi},MIN(ROWS(LK_WSize),"
                  f"MAX({pv('minw')},{xa},{xb},{xc},IF({ns}>1,{pv('i10')},1)))))")
        ws[h["gov"]] = (f'=IF({wi}=0,"",IF({oa}>0,"override",IF({xc}={wi},"OCPD",IF({xa}={wi},"125%",'
                        f'IF({xb}={wi},"derated load",IF(AND({ns}>1,{wi}={pv("i10")}),"parallel 1/0","min size"))))))')
        ws[ei] = f"=IF({wi}=0,0,{egc_index(t, wi)})"
        ws[h["nN"]] = f"=IF({Pp}=1,1,IF(AND({v}=1,{xf}=0),{h['sN']},0))"
        show = f'AND({wi}>0,{h["cont"]}="",OR({VA}>0,{v}=1))'
        ws[h["wlab"]] = f'=IF({show},IF({ns}>1,{ns}&"# ","")&INDEX(LK_WLabel,{wi}),"")'
        ws[h["cab"]] = (f'=IF({show},IF({ns}>1,{ns}&" sets: ","")&({Pp}+{h["nN"]})&INDEX(LK_WShort,{wi})'
                        f'&" + "&INDEX(LK_WShort,{ei})&"G "&{pv("mat")}&"  ["&{h["gov"]}&"]","")')


def zone_visible_formulas(ws, Z, H, L):
    for r in range(L.R0, L.R1 + 1):
        ph = f"${ROWPH}{r}"
        n1 = f"${ROWPH}{r+1}" if r + 1 <= L.R1 else '""'
        n2 = f"${ROWPH}{r+2}" if r + 2 <= L.R1 else '""'
        ws[f"{Z['ph']}{r}"] = (f'=IF({H["cont"]}{r}<>"",{ph},{ph}&IF({H["P"]}{r}>=2,{n1},"")'
                               f'&IF({H["P"]}{r}>=3,{n2},""))')
        ws[f"{Z['va']}{r}"] = f'=IF({H["VA"]}{r}>0,{H["VA"]}{r},"")'
        ws[f"{Z['amp']}{r}"] = f'=IF({H["VA"]}{r}>0,{H["amp"]}{r},"")'
        ws[f"{Z['wire']}{r}"] = f'={H["cab"]}{r}'
        ws[f"{Z['stat']}{r}"] = status_formula(Z, H, r)


def build_panel(wb, sheet_name, hdr_vals, circuits, cap=42, tab_color="1F4E79", hide_unused=False):
    L = Layout(cap)
    R0_, R1 = L.R0, L.R1
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85

    # ---------------- column widths / row heights ----------------
    for Z in (ODD, EVEN):
        for k, col in Z.items():
            ws.column_dimensions[col].width = ZONE_WIDTHS[k]
    for k, col in S.items():
        ws.column_dimensions[col].width = SCH_WIDTHS[k]
    for g in GAPS:
        ws.column_dimensions[g].width = 1.5
    for ci in range(CALC0, CI(CALC_LAST) + 1):
        ws.column_dimensions[CL(ci)].width = 10
        ws.column_dimensions[CL(ci)].hidden = True
    ws.column_dimensions[PL].width = 30

    heights = {1: 30, 2: 28, 3: 24, 4: 34, 5: 22, 6: 22, 7: 18, 8: 18, 9: 34}
    for r in range(R0_, R1 + 1):
        heights[r] = 24
    for r in range(L.TOT, L.TOT + 3):
        heights[r] = 24
    heights[L.TOT + 2] = 26
    for r in range(L.BOT, L.LAST + 1):
        heights[r] = 21
    heights[L.BOT] = 30
    for r, h in heights.items():
        ws.row_dimensions[r].height = h

    # ================= HIDDEN CALC ZONE =================
    lab = font(8, False, "595959")
    for key, row in P.items():
        ws[f"{PL}{row}"] = P_LABELS[key]
        ws[f"{PL}{row}"].font = lab
    dm = f'MATCH({IN["dist"]},LK_DistName,0)'
    f = {
        "name": f'=TRIM({IN["name"]})',
        "title": f'=IF(TRIM({IN["title"]})="","Panelboard",TRIM({IN["title"]}))',
        "loc": f'=IF(LEN({IN["loc"]})=0,"",{IN["loc"]})',
        "manual": f'=IF(LEN({IN["manual"]})=0,"",{IN["manual"]})',
        "dist": f'=IF(LEN({IN["dist"]})=0,"",{IN["dist"]})',
        "ph": f'=IFERROR(INDEX(LK_DistPh,{dm}),3)',
        "wires": f'=IFERROR(INDEX(LK_DistW,{dm}),4)',
        "vln": f'=IFERROR(INDEX(LK_DistVLN,{dm}),120)',
        "vll": f'=IFERROR(INDEX(LK_DistVLL,{dm}),208)',
        "hot": f'=IF({pv("ph")}=3,3,IF({pv("wires")}>=3,2,1))',
        "neut": f'=IF({pv("wires")}>{pv("hot")},1,0)',
        "div": f'={pv("hot")}*{pv("vln")}',
        "cap": f"={cap}",
        "slots": f'=IF(ISNUMBER({IN["slots"]}),MAX(2,MIN({pv("cap")},{IN["slots"]})),{pv("cap")})',
        "mains": f'=IF(ISNUMBER({IN["mains"]}),{IN["mains"]},"")',
        "mcb": f'=IF(ISNUMBER({IN["mcb"]}),{IN["mcb"]},"")',
        "mtype": f'=IF(LEN({IN["mtype"]})=0,"",{IN["mtype"]})',
        "aic": f'=IF(LEN({IN["aic"]})=0,"",{IN["aic"]})',
        "conn": f'=SUM({CB["conn"]}{CLS_R0}:{CB["conn"]}{CLS_R0+NCLS-1})',
        "dem": f'=SUM({CB["dem"]}{CLS_R0}:{CB["dem"]}{CLS_R0+NCLS-1})',
        "amax": f'=MAX({pv("aA")},{pv("aB")},{pv("aC")})',
        "connA": f'={pv("amax")}',
        "demA": f'=IF({pv("conn")}>0,{pv("amax")}*{pv("dem")}/{pv("conn")},0)',
        "def": '=IF(LEN(LK_DefaultClass)=0,"Other",LK_DefaultClass)',
        "lim": '=IF(ISNUMBER(LK_BrkLimit),LK_BrkLimit,0.8)',
        "empty": '=T(LK_EmptyLabel)',
        "err": f'=COUNTIF(${ODD["stat"]}${R0_}:${ODD["stat"]}${R1},"ERR*")+COUNTIF(${EVEN["stat"]}${R0_}:${EVEN["stat"]}${R1},"ERR*")',
        "warn": f'=COUNTIF(${ODD["stat"]}${R0_}:${ODD["stat"]}${R1},"WARN*")+COUNTIF(${EVEN["stat"]}${R0_}:${EVEN["stat"]}${R1},"WARN*")',
        "hl1p": (f'=COUNTIF(${ODD["stat"]}${R0_}:${ODD["stat"]}${R1},"ERR: 1P load on high leg*")'
                 f'+COUNTIF(${EVEN["stat"]}${R0_}:${EVEN["stat"]}${R1},"ERR: 1P load on high leg*")'),
        "nfeed": f'=COUNTIF({OH["valid"]}{R0_}:{OH["valid"]}{R1},1)+COUNTIF({EH["valid"]}{R0_}:{EH["valid"]}{R1},1)',
        "avg": f'=IF({pv("hot")}>=2,({pv("phA")}+{pv("phB")}+IF({pv("hot")}=3,{pv("phC")},0))/{pv("hot")},{pv("phA")})',
        "fed": f'=IFERROR(INDEX({reg_rng(RV["fed"])},MATCH({pv("name")},LK_PanelNames,0))&"","")',
        "fedckt": f'=IFERROR(INDEX({reg_rng(RV["fckt"])},MATCH({pv("name")},LK_PanelNames,0))&"","")',
        "fedbrk": f'=IFERROR(INDEX({reg_rng(RV["fbrk"])},MATCH({pv("name")},LK_PanelNames,0))&"","")',
        "supply": f'=IF(AND({pv("fed")}<>"",{pv("fed")}<>"(multiple)"),{pv("fed")},{pv("manual")})',
        "rating": f'=IF(ISNUMBER({pv("mcb")}),{pv("mcb")},IF(ISNUMBER({pv("mains")}),{pv("mains")},""))',
        "loading": f'=IF(ISNUMBER({pv("rating")}),{pv("demA")}/{pv("rating")},"")',
        "ncls": f'=COUNT({CB["ord"]}{CLS_R0}:{CB["ord"]}{CLS_R0+NCLS-1})',
        "chk": f'={pv("phA")}+{pv("phB")}+{pv("phC")}-{pv("conn")}',
        # high leg
        "hlin": f'=UPPER(TRIM({IN["hlin"]}&""))',
        "hl": (f'=IF({pv("ph")}<>3,"",IFERROR(IF(T(INDEX(LK_DistHL,{dm}))="","",'
               f'IF(OR({pv("hlin")}="A",{pv("hlin")}="B",{pv("hlin")}="C"),{pv("hlin")},'
               f'UPPER(T(INDEX(LK_DistHL,{dm}))))),""))'),
        "vhl": f'=IF({pv("hl")}="",0,{pv("vll")}*SQRT(3)/2)',
        "sysbox": (f'=IF({pv("ph")}=3,IF({pv("hl")}<>"","Three Phase  Δ",IF({pv("wires")}=3,"Three Phase  Δ","Three Phase")),'
                   f'"Single Phase")'),
        "hlnote": (f'=IF({pv("hl")}="","","HIGH LEG: L-"&(CODE({pv("hl")})-64)&" ("&{pv("hl")}&") = "'
                   f'&TEXT({pv("vhl")},"0")&"V TO N.  NO 1-POLE LOADS.  MARK ORANGE (NEC 110.15).")'),
        "imb": (f'=IF(OR({pv("hot")}<2,{pv("avg")}=0),0,IF({pv("hl")}<>"",'
                f'IF(IF({pv("hl")}="A",{pv("phB")},{pv("phA")})+IF({pv("hl")}="C",{pv("phB")},{pv("phC")})=0,0,'
                f'ABS(IF({pv("hl")}="A",{pv("phB")},{pv("phA")})-IF({pv("hl")}="C",{pv("phB")},{pv("phC")}))'
                f'/(IF({pv("hl")}="A",{pv("phB")},{pv("phA")})+IF({pv("hl")}="C",{pv("phB")},{pv("phC")}))),'
                f'MAX(ABS({pv("phA")}-{pv("avg")}),ABS({pv("phB")}-{pv("avg")}),'
                f'IF({pv("hot")}=3,ABS({pv("phC")}-{pv("avg")}),0))/{pv("avg")}))'),
        # conductors
        "mat": (f'=IF(LEFT(UPPER(TRIM({IN["mat"]}&"")),2)="AL","Al",IF(LEFT(UPPER(TRIM({IN["mat"]}&"")),2)="CU","Cu",'
                f'IF(UPPER(TRIM(LK_DefMat&""))="AL","Al","Cu")))'),
        "ins": (f'=IF(ISNUMBER(SEARCH("90",{IN["mat"]}&"")),90,IF(ISNUMBER(SEARCH("75",{IN["mat"]}&"")),75,'
                f'IF(N(LK_DefIns)=90,90,75)))'),
        "b2404": '=IF(UPPER(LEFT(TRIM(LK_Use2404B&""),1))="Y",1,0)',
        "contf": (f'=IF({pv("conn")}<=0,0,SUMPRODUCT({CB["conn"]}{CLS_R0}:{CB["conn"]}{CLS_R0+NCLS-1},'
                  f'--({CB["cflag"]}{CLS_R0}:{CB["cflag"]}{CLS_R0+NCLS-1}="Yes"))/{pv("conn")})'),
        "basis": (f'="CONDUCTORS: "&IF({pv("mat")}="Al","ALUMINUM","COPPER")&", "&{pv("ins")}&" °C MIN. ("'
                  f'&IF({pv("ins")}=90,"THHN/THWN-2, XHHW-2","THWN, THHN/THWN-2, XHHW-2")&"), SIZED TO NEC 310.16"'
                  f'&IF({pv("fct")}<0.995,", DERATED x"&TEXT({pv("fct")},"0.00")&" ("&{pv("amb")}&" °C, "&{pv("ccc")}'
                  f'&" CURRENT-CARRYING CONDUCTORS)","")&"."'),
        "term": (f'=IF(OR(N({IN["term"]})=60,N({IN["term"]})=75,N({IN["term"]})=90),N({IN["term"]}),'
                 f'IF(OR(N(LK_DefTerm)=60,N(LK_DefTerm)=90),N(LK_DefTerm),75))'),
        "amb": f'=IF(ISNUMBER({IN["amb"]}),{IN["amb"]},IF(ISNUMBER(LK_DefAmb),LK_DefAmb,30))',
        "ambf": (f'=INDEX(LK_AmbF,MIN(ROWS(LK_AmbT),COUNTIF(LK_AmbT,"<"&{pv("amb")})+1),'
                 f'({pv("ins")}-45)/15)'),
        "ccc": f'=IF(ISNUMBER({IN["ccc"]}),MAX(1,{IN["ccc"]}),IF(ISNUMBER(LK_DefCCC),LK_DefCCC,3))',
        "cccf": f'=INDEX(LK_CccF,MIN(ROWS(LK_CccF),COUNTIF(LK_CccN,"<"&{pv("ccc")})+1))',
        "fct": f'={pv("ambf")}*{pv("cccf")}',
        "ccol": f'=IF({pv("mat")}="Al",3,0)+(MIN({pv("term")},{pv("ins")})-45)/15',
        "icol": f'=IF({pv("mat")}="Al",3,0)+({pv("ins")}-45)/15',
        "lcol": f'=IF({pv("mat")}="Al",2,1)',
        "minw": '=IFERROR(MATCH(TRIM(LK_MinWire&""),LK_WSize,0),1)',
        "maxw": '=IFERROR(MATCH(TRIM(LK_MaxWire&""),LK_WSize,0),ROWS(LK_WSize))',
        "maxamp": "=" + allow(pv("maxw")),
        "maxterm": f'=INDEX(LK_Amp,{pv("maxw")},{pv("ccol")})',
        "nov": f'=COUNTIF({OH["ov"]}{R0_}:{OH["ov"]}{R1},1)+COUNTIF({EH["ov"]}{R0_}:{EH["ov"]}{R1},1)',
        "ntyped": "=" + "+".join(
            f'SUMPRODUCT(--NOT(_xlfn.ISFORMULA({S[a_]}{R0_}:{S[b_]}{R1})))'
            for a_, b_ in (("ckto", "wireo"), ("L1o", "L3e"), ("wiree", "ckte"))),
        "brkmode": (f'=IF(ISNUMBER(SEARCH("ALL",{IN["brk100"]}&"")),2,'
                    f'IF(ISNUMBER(SEARCH("MCB",{IN["brk100"]}&"")),1,0))'),
        "n100": (f'=COUNTIF(${ODD["brk"]}${R0_}:${ODD["brk"]}${R1},"100%*")'
                 f'+COUNTIF(${EVEN["brk"]}${R0_}:${EVEN["brk"]}${R1},"100%*")'),
        "mcbreq": (f'={pv("demA")}*(IF({pv("brkmode")}>=1,1,1.25)*{pv("contf")}+(1-{pv("contf")}))'),
        "brknote": (f'=IF({pv("brkmode")}=2,"ALL BREAKERS 100% RATED; ASSEMBLY LISTED FOR 100% CONTINUOUS OPERATION.",'
                    f'IF(AND({pv("brkmode")}=1,ISNUMBER({pv("mcb")})),"MAIN BREAKER 100% RATED (ASSEMBLY LISTED FOR 100% '
                    f'OPERATION). ALL OTHER BREAKERS 80% RATED (STANDARD)","ALL BREAKERS 80% RATED (STANDARD)")'
                    f'&IF({pv("n100")}>0," UNLESS NOTED 100% RATED","")&".")'),
        "i10": '=MATCH("1/0",LK_WSize,0)',
        "fns": f'=IF(ISNUMBER({pv("rating")}),MAX(1,ROUNDUP({pv("rating")}/{pv("maxamp")},0)),0)',
        "fwi": f'=IF({pv("fns")}=0,0,{pick_index(pv("rating"), pv("fns"))})',
        "fei": f'=IF({pv("fwi")}=0,0,{egc_index(pv("rating"), pv("fwi"))})',
        "fcab": (f'=IF({pv("fwi")}=0,"",IF({pv("fns")}>1,{pv("fns")}&"# ","")&INDEX(LK_WLabel,{pv("fwi")})'
                 f'&" x "&({pv("hot")}+{pv("neut")})&"C + "&INDEX(LK_WLabel,{pv("fei")})&" (G) "&{pv("mat")}'
                 f'&" "&{pv("ins")}&"°C")'),
        "mca": (f'=IF({pv("demA")}>0,INDEX(LK_Breakers,MIN(ROWS(LK_Breakers),'
                f'COUNTIF(LK_Breakers,"<"&{pv("demA")})+1)),"")'),
    }
    for ph_key, letter in [("phA", "A"), ("phB", "B"), ("phC", "C")]:
        f[ph_key] = (f'=SUMIF(${ROWPH}${R0_}:${ROWPH}${R1},"{letter}",{OH["slot"]}${R0_}:{OH["slot"]}${R1})'
                     f'+SUMIF(${ROWPH}${R0_}:${ROWPH}${R1},"{letter}",{EH["slot"]}${R0_}:{EH["slot"]}${R1})')
    for a_key, letter in [("aA", "A"), ("aB", "B"), ("aC", "C")]:
        f[a_key] = (f'=SUMIF(${ROWPH}${R0_}:${ROWPH}${R1},"{letter}",{OH["sla"]}${R0_}:{OH["sla"]}${R1})'
                    f'+SUMIF(${ROWPH}${R0_}:${ROWPH}${R1},"{letter}",{EH["sla"]}${R0_}:{EH["sla"]}${R1})')
    assert set(f) == set(P), set(P) ^ set(f)
    for key, formula in f.items():
        ws[f"{PV}{P[key]}"] = formula

    # row phase
    for r in range(R0_, R1 + 1):
        i = f"(${ODD['ckt']}{r}-1)/2"
        ws[f"{ROWPH}{r}"] = (f'=IF({pv("ph")}=3,CHOOSE(MOD({i},3)+1,"A","B","C"),'
                             f'IF({pv("wires")}>=3,CHOOSE(MOD({i},2)+1,"A","B"),"A"))')
    helper_formulas(ws, ODD, OH, "odd", L)
    helper_formulas(ws, EVEN, EH, "even", L)

    ws[f"{ROWPH}9"] = "row phase"
    for H, tag in [(OH, "odd"), (EH, "even")]:
        for k, col in H.items():
            ws[f"{col}9"] = f"{tag}:{k}"
    for k, col in CB.items():
        ws[f"{col}9"] = f"class:{k}"
    for k, col in FB.items():
        ws[f"{col}9"] = f"feeder:{k}"
    for j in range(NFEED):
        ws[f"{GC[j]}9"] = f"subconn{j+1}"; ws[f"{GN[j]}9"] = f"subcnt{j+1}"; ws[f"{GM[j]}9"] = f"submax{j+1}"

    # class block
    okey, oown = f"{OH['key']}${R0_}:{OH['key']}${R1}", f"{OH['own']}${R0_}:{OH['own']}${R1}"
    ekey, eown = f"{EH['key']}${R0_}:{EH['key']}${R1}", f"{EH['own']}${R0_}:{EH['own']}${R1}"
    for k in range(1, NCLS + 1):
        r = CLS_R0 + k - 1
        C = {kk: f"{col}{r}" for kk, col in CB.items()}
        ws[C["cname"]] = f"=T(INDEX(LK_ClassName,{k}))"
        nm = C["cname"]
        ws[C["loc"]] = f'=IF({nm}="",0,SUMIF({okey},{nm},{oown})+SUMIF({ekey},{nm},{eown}))'
        ws[C["lcnt"]] = f'=IF({nm}="",0,COUNTIF({okey},{nm})+COUNTIF({ekey},{nm}))'
        ws[C["lmax"]] = (f'=IF({nm}="",0,MAX(SUMPRODUCT(MAX(({okey}={nm})*{oown})),'
                         f'SUMPRODUCT(MAX(({ekey}={nm})*{eown}))))')
        ws[C["subc"]] = f"=SUM({GC[0]}{r}:{GC[-1]}{r})"
        ws[C["subn"]] = f"=SUM({GN[0]}{r}:{GN[-1]}{r})"
        ws[C["subm"]] = f"=MAX({GM[0]}{r}:{GM[-1]}{r})"
        ws[C["conn"]] = f"={C['loc']}+{C['subc']}"
        ws[C["cnt"]] = f"={C['lcnt']}+{C['subn']}"
        ws[C["mx"]] = f"=MAX({C['lmax']},{C['subm']})"
        ws[C["meth"]] = f'=IF({nm}="","",IF(T(INDEX(LK_Method,{k}))="","Constant",T(INDEX(LK_Method,{k}))))'
        ws[C["f1"]] = f"=IF(ISNUMBER(INDEX(LK_F1,{k})),INDEX(LK_F1,{k}),1)"
        ws[C["l1"]] = f"=IF(ISNUMBER(INDEX(LK_L1,{k})),INDEX(LK_L1,{k}),1E+15)"
        ws[C["f2"]] = f"=IF(ISNUMBER(INDEX(LK_F2,{k})),INDEX(LK_F2,{k}),{C['f1']})"
        ws[C["l2"]] = f"=IF(ISNUMBER(INDEX(LK_L2,{k})),INDEX(LK_L2,{k}),1E+15)"
        ws[C["f3"]] = f"=IF(ISNUMBER(INDEX(LK_F3,{k})),INDEX(LK_F3,{k}),{C['f2']})"
        cn, m = C["conn"], C["meth"]
        tier = (f"MIN({cn},{C['l1']})*{C['f1']}+MAX(0,MIN({cn},{C['l2']})-{C['l1']})*{C['f2']}"
                f"+MAX(0,{cn}-{C['l2']})*{C['f3']}")
        ws[C["dem"]] = (f'=IF({cn}<=0,0,IF({m}="Quantity",{cn}*INDEX(LK_QtyFactor,MAX(1,MIN({C["cnt"]},ROWS(LK_QtyFactor)))),'
                        f'IF(LEFT({m},7)="Largest",{cn}*{C["f1"]}+0.25*{C["mx"]},{tier})))')
        ws[C["df"]] = f'=IF({cn}>0,{C["dem"]}/{cn},"")'
        ws[C["ord"]] = f'=IF({cn}>0,COUNTIF({CB["conn"]}${CLS_R0}:{cn},">0"),"")'
        ws[C["cflag"]] = f'=IF({nm}="","",IF(T(INDEX(LK_Cont,{k}))="Yes","Yes","No"))'
        for j in range(NFEED):
            fn = f"${FB['fname']}${FEED_R0+j}"
            for grid, src in [(GC, CB["conn"]), (GN, CB["cnt"]), (GM, CB["mx"])]:
                ws[f"{grid[j]}{r}"] = f'=IF({fn}="",0,N(INDIRECT("\'"&{fn}&"\'!{src}{r}")))'

    # feeder list
    def pick(oc, ec, j):
        return (f"IFERROR(INDEX({oc}${R0_}:{oc}${R1},MATCH({j},{OH['frank']}${R0_}:{OH['frank']}${R1},0)),"
                f"IFERROR(INDEX({ec}${R0_}:{ec}${R1},MATCH({j},{EH['frank']}${R0_}:{EH['frank']}${R1},0)),\"\"))")
    for j in range(1, NFEED + 1):
        r = FEED_R0 + j - 1
        F = {kk: f"{col}{r}" for kk, col in FB.items()}
        ws[F["fname"]] = "=" + pick(OH["sub"], EH["sub"], j)
        ws[F["fcktn"]] = "=" + pick("$" + ODD["ckt"], "$" + EVEN["ckt"], j)
        ws[F["fpol"]] = "=" + pick(OH["P"], EH["P"], j)
        ws[F["ftrip"]] = "=" + pick("$" + ODD["trip"], "$" + EVEN["trip"], j)
        ws[F["fxf"]] = "=" + pick(OH["xfm"], EH["xfm"], j)
        fn, n, p = F["fname"], F["fcktn"], F["fpol"]
        ws[F["fckt"]] = f'=IF({fn}="","",{n}&IF({p}>=2,","&({n}+2),"")&IF({p}>=3,","&({n}+4),""))'
        ws[F["fbrk"]] = (f'=IF({fn}="","",IF(N({F["ftrip"]})>0,{F["ftrip"]}&" A / ","")&{p}&"P"'
                         f'&IF(N({F["fxf"]})=1," · XFMR",""))')
        ws[F["fva"]] = f'=IF({fn}="","",N(INDIRECT("\'"&{fn}&"\'!${PV}${P["conn"]}")))'
        ws[F["fdem"]] = f'=IF({fn}="","",N(INDIRECT("\'"&{fn}&"\'!${PV}${P["dem"]}")))'

    for row in ws.iter_rows(min_row=1, max_row=max(R1, 75), min_col=CALC0, max_col=CI(CALC_LAST)):
        for c in row:
            if c.value is not None:
                c.font = font(8, False, "595959")

    # ================= INPUT ZONES =================
    for Z, title in [(ODD, "PANEL DATA   ·   ODD CIRCUITS  (left side of schedule)"),
                     (EVEN, "PANEL CHECKS   ·   EVEN CIRCUITS  (right side of schedule)")]:
        first, last = Z["ckt"], Z["stat"]
        merge(ws, f"{first}1:{last}1")
        set_cell(ws, f"{first}1", title, font(10, True, "FFFFFF"), AL_L, fl=F_ZTITLE)
        heads = {"ckt": "CKT", "ph": "Phase", "desc": "Circuit Description", "cls": "Load Classification",
                 "q1": "Fixture 1\nQty", "v1": "Fixture 1\nVA each", "q2": "Fixture 2\nQty",
                 "v2": "Fixture 2\nVA each", "add": "Added VA", "sub": "Sub-Panel\n(feeds)",
                 "xf": "Via\nXFMR", "pol": "Poles", "trip": "Trip (A)", "brk": "Breaker\nType",
                 "wov": "Wire\noverride", "ovr": "VA\noverride", "why": "Override reason",
                 "va": "Load (VA)", "amp": "Load (A)",
                 "wire": "Cable (auto)", "stat": "Status"}
        for k, col in Z.items():
            set_cell(ws, f"{col}9", heads[k], font(8, True, C_LABEL), AL_HDR, fl=F_ZHDR, border=GRID_LIGHT)
        for r in range(R0_, R1 + 1):
            ck = 2 * (r - R0_) + 1 + (0 if Z is ODD else 1)
            set_cell(ws, f"{Z['ckt']}{r}", ck, font(9, True, C_LABEL), AL_C, fl=F_ZHDR, border=GRID_LIGHT)
            set_cell(ws, f"{Z['ph']}{r}", None, font(9, False, "595959"), AL_C, fl=F_CALC, border=GRID_LIGHT)
            for k in INPUT_KEYS:
                fmt = {"q1": "0", "q2": "0", "v1": "#,##0", "v2": "#,##0", "add": "#,##0",
                       "pol": "0", "trip": "0", "wov": "@", "ovr": "#,##0"}.get(k)
                al = AL_L if k in ("desc", "cls", "sub", "brk", "why") else AL_C
                set_cell(ws, f"{Z[k]}{r}", None, font(9, False, C_INPUT), al, fmt, F_INPUT, GRID_LIGHT)
            set_cell(ws, f"{Z['va']}{r}", None, font(9), AL_R, "#,##0", F_CALC, GRID_LIGHT)
            set_cell(ws, f"{Z['amp']}{r}", None, font(9), AL_R, "0.0", F_CALC, GRID_LIGHT)
            set_cell(ws, f"{Z['wire']}{r}", None, font(8), AL_L, None, F_CALC, GRID_LIGHT)
            set_cell(ws, f"{Z['stat']}{r}", None, font(8), AL_L, None, F_CALC, GRID_LIGHT)
    zone_visible_formulas(ws, ODD, OH, L)
    zone_visible_formulas(ws, EVEN, EH, L)

    # ---- panel data inputs (left zone rows 2..8)
    lbl = font(8, False, C_LABEL)
    g1 = [("name", "Panel Name  (must match tab name)"), ("title", "Schedule Title"),
          ("loc", "Location"), ("manual", "Supply From  (manual / top level)"),
          ("mount", "Mounting"), ("encl", "Enclosure"), ("slots", f"No. of Circuits  (this sheet: up to {cap})")]
    for i, (k, t) in enumerate(g1):
        r = 2 + i
        set_cell(ws, f"C{r}", t, lbl, AL_RI)
        set_cell(ws, IN[k], hdr_vals.get(k), font(9, k == "name", C_INPUT), AL_L, None, F_INPUT, GRID_LIGHT)
    g2 = [("dist", "Distribution System"), ("aic", "A.I.C. Rating"), ("mtype", "Mains Type"),
          ("mains", "Mains (Bus) Rating, A"), ("mcb", "MCB Rating, A"),
          ("notes1", "Notes (top of schedule)"), ("notes2", "Notes (bottom of schedule)")]
    for i, (k, t) in enumerate(g2):
        r = 2 + i
        merge(ws, f"E{r}:H{r}")
        set_cell(ws, f"E{r}", t, lbl, AL_RI)
        rng = f"I{r}:K{r}" if r <= 6 else f"I{r}:O{r}"
        merge(ws, rng)
        border_range(ws, rng, all_=thin_light)
        set_cell(ws, IN[k], hdr_vals.get(k), font(9, False, C_INPUT), AL_L,
                 "0" if k in ("mains", "mcb") else None, F_INPUT)
        for cc in ws[rng][0]:
            cc.fill = F_INPUT
    g4 = [("hlin", "High-leg phase (Δ systems only; blank = Lookups default)"),
          ("mat", "Conductor: material + insulation basis (Cu 75C default)"), ("term", "Termination rating, °C"),
          ("amb", "Ambient temperature, °C"), ("ccc", "Current-carrying cond. per raceway")]
    for i, (k, t) in enumerate(g4):
        r = 2 + i
        merge(ws, f"L{r}:N{r}")
        set_cell(ws, f"L{r}", t, lbl, Alignment(horizontal="right", vertical="center", wrap_text=True))
        set_cell(ws, IN[k], hdr_vals.get(k), font(9, False, C_INPUT), AL_C, None, F_INPUT, GRID_LIGHT)
    merge(ws, "P7:Q7")
    set_cell(ws, "P7", "Breaker rating", lbl, AL_RI)
    merge(ws, "R7:S7")
    border_range(ws, "R7:S7", all_=thin_light)
    set_cell(ws, IN["brk100"], hdr_vals.get("brk100"), font(9, False, C_INPUT), AL_L, None, F_INPUT)
    ws["S7"].fill = F_INPUT
    g3 = [("System", f'={pv("ph")}&" ph / "&{pv("wires")}&" W   ·   "&{pv("vln")}&" / "&{pv("vll")}&" V"'
                     f'&IF({pv("hl")}<>"","   ·   high leg "&{pv("hl")}&" = "&TEXT({pv("vhl")},"0")&" V to N","")'),
          ("Fed from (auto)", f'=IF({pv("fed")}="","— not fed from a listed panel",'
                              f'{pv("fed")}&"   ·   CKT "&{pv("fedckt")}&"   ·   "&{pv("fedbrk")})'),
          ("Demand", f'=TEXT({pv("dem")},"#,##0")&" VA   ·   "&TEXT({pv("demA")},"0.0")&" A (max line)"'),
          ("Connected", f'=TEXT({pv("conn")},"#,##0")&" VA   ·   "&TEXT({pv("amax")},"0.0")&" A (max line)"'),
          ("Feeder cable", f'=IF({pv("fcab")}="","(enter MCB or bus rating)",{pv("fcab")}&"   ·   factor "'
                           f'&TEXT({pv("fct")},"0.00")&" @ "&{pv("term")}&" °C")')]
    for i, (t, fml) in enumerate(g3):
        r = 2 + i
        merge(ws, f"P{r}:Q{r}")
        set_cell(ws, f"P{r}", t, lbl, AL_RI)
        merge(ws, f"R{r}:S{r}")
        set_cell(ws, f"R{r}", fml, font(9), AL_L, fl=F_CALC)
        border_range(ws, f"R{r}:S{r}", all_=thin_light)
        ws[f"S{r}"].fill = F_CALC

    # ---- panel checks (right zone rows 2..8)
    E = EVEN
    chk_rows = [
        ("Demand vs MCB / bus",
         f'=IF({pv("rating")}="","No MCB or bus rating entered",TEXT({pv("demA")},"0.0")&" A est. demand (max line) = "'
         f'&TEXT({pv("loading")},"0%")&" of "&{pv("rating")}&" A "&IF(ISNUMBER({pv("mcb")}),"MCB","bus")'
         f'&IF(ISNUMBER({pv("mcb")}),"  ·  NEC min MCB "&TEXT({pv("mcbreq")},"0")&" A ("'
         f'&IF({pv("brkmode")}>=1,"100% rated","125% continuous")&")",""))',
         f'=IF({pv("rating")}="","WARN",IF(OR({pv("loading")}>1,AND(ISNUMBER({pv("mcb")}),{pv("mcbreq")}>N({pv("mcb")})+0.001)),"ERR",'
         f'IF({pv("loading")}>LK_MCBWarn,"WARN","OK")))'),
        ("MCB vs bus rating",
         f'=IF(AND(ISNUMBER({pv("mains")}),ISNUMBER({pv("mcb")})),IF({pv("mcb")}>{pv("mains")},"MCB "&{pv("mcb")}'
         f'&" A exceeds the "&{pv("mains")}&" A bus rating","MCB "&{pv("mcb")}&" A ≤ bus "&{pv("mains")}&" A"),'
         f'"Enter both ratings to check")',
         f'=IF(AND(ISNUMBER({pv("mains")}),ISNUMBER({pv("mcb")})),IF({pv("mcb")}>{pv("mains")},"ERR","OK"),"–")'),
        ("Line currents / balance",
         f'="A "&TEXT({pv("aA")},"0.0")&IF({pv("hot")}>=2,"  ·  B "&TEXT({pv("aB")},"0.0"),"")'
         f'&IF({pv("hot")}=3,"  ·  C "&TEXT({pv("aC")},"0.0"),"")&" A"&IF({pv("hot")}>=2,"   →  "'
         f'&TEXT({pv("imb")},"0%")&IF({pv("hl")}<>""," A-C lighting-leg imbalance"," max deviation"),"")',
         f'=IF({pv("hot")}<2,"–",IF({pv("imb")}>LK_ImbLimit,"WARN","OK"))'),
        ("Circuit flags",
         f'={pv("err")}&" circuit(s) ERR   ·   "&{pv("warn")}&" circuit(s) WARN"'
         f'&IF({pv("nov")}>0,"   ·   "&{pv("nov")}&" VA override(s)","")'
         f'&IF({pv("ntyped")}>0,"   ·   "&{pv("ntyped")}&" schedule cell(s) TYPED OVER: totals ignore them","")',
         f'=IF(OR({pv("err")}>0,{pv("ntyped")}>0),"ERR",IF({pv("warn")}>0,"WARN","OK"))'),
        ("High leg",
         f'=IF({pv("hl")}<>"","High leg "&{pv("hl")}&" = "&TEXT({pv("vhl")},"0")&" V to N   ·   "&{pv("hl1p")}'
         f'&" 1-pole load(s) on it"&IF({pv("hl")}<>"B","   ·   NEC 408.3(E)(1): B position unless meter equipment",""),'
         f'IF({pv("hlin")}<>"","High-leg phase ignored: "&{pv("dist")}&" has no high leg","No high leg on this system"))',
         f'=IF({pv("hl")}="",IF({pv("hlin")}<>"","WARN","–"),IF({pv("hl1p")}>0,"ERR",IF({pv("hl")}<>"B","WARN","OK")))'),
        ("Cable basis",
         f'={pv("mat")}&" "&{pv("ins")}&" °C basis · term "&{pv("term")}&" °C · "&{pv("amb")}&" °C x"&TEXT({pv("ambf")},"0.00")'
         f'&" · "&{pv("ccc")}&" CCC x"&TEXT({pv("cccf")},"0.00")&" → factor "&TEXT({pv("fct")},"0.00")'
         f'&IF({pv("ambf")}<=0,"  ·  ambient too high for this insulation",IF({pv("fct")}<LK_DerateMin,'
         f'"  ·  below "&TEXT(LK_DerateMin,"0%")&": split raceway or upsize",""))',
         f'=IF({pv("ambf")}<=0,"ERR",IF({pv("fct")}<LK_DerateMin,"WARN","OK"))'),
        ("Capacity / integrity",
         f'={pv("ncls")}&" load classes ({SCH_CLS_ROWS} shown)  ·  "&{pv("nfeed")}&" sub-panels fed (max {NFEED})"'
         f'&IF(N({IN["slots"]})>{pv("cap")},"  ·  {cap}-ckt sheet: use Template 84","")'
         f'&IF(ABS({pv("chk")})>0.5,"  ·  phase total ≠ class total","")'
         f'&IF(ISNUMBER(MATCH({pv("name")},LK_PanelNames,0)),"","  ·  not on Panels sheet")',
         f'=IF(OR({pv("ncls")}>{SCH_CLS_ROWS},{pv("nfeed")}>{NFEED},ABS({pv("chk")})>0.5,N({IN["slots"]})>{pv("cap")}),"ERR",'
         f'IF(ISNUMBER(MATCH({pv("name")},LK_PanelNames,0)),"OK","WARN"))'),
    ]
    for i, (t, txt, flag) in enumerate(chk_rows):
        r = 2 + i
        set_cell(ws, f"{E['desc']}{r}", t, lbl, AL_RI)
        rng = f"{E['cls']}{r}:{E['wire']}{r}"
        merge(ws, rng)
        set_cell(ws, f"{E['cls']}{r}", txt, font(8), AL_L)
        for cc in ws[rng][0]:
            cc.fill = F_CALC
        border_range(ws, rng, all_=thin_light)
        set_cell(ws, f"{E['stat']}{r}", flag, font(8, True), AL_C, fl=F_CALC, border=GRID_LIGHT)

    # ---- sub-panels fed table (right zone below circuits)
    t0 = R1 + 2
    set_cell(ws, f"{E['ckt']}{t0}", "SUB-PANELS FED FROM THIS PANEL", font(9, True, C_LABEL), AL_L)
    cols = [(E["ckt"], E["ckt"], "#"), (E["ph"], E["desc"], "Sub-panel"), (E["cls"], E["cls"], "CKT"),
            (E["q1"], E["sub"], "Breaker"), (E["xf"], E["trip"], "Conn. VA"), (E["brk"], E["amp"], "Demand VA"),
            (E["wire"], E["wire"], "Cable")]
    for a, b, t in cols:
        if a != b:
            merge(ws, f"{a}{t0+1}:{b}{t0+1}")
        set_cell(ws, f"{a}{t0+1}", t, font(8, True, C_LABEL), AL_C, fl=F_ZHDR)
        border_range(ws, f"{a}{t0+1}:{b}{t0+1}", all_=thin_light)
        for cc in ws[f"{a}{t0+1}:{b}{t0+1}"][0]:
            cc.fill = F_ZHDR
    for j in range(1, NFEED + 1):
        r = t0 + 1 + j
        fr = FEED_R0 + j - 1
        cab = (f'=IFERROR(IF({FB["fname"]}{fr}="","",INDEX({OH["cab"]}${R0_}:{OH["cab"]}${R1},'
               f'MATCH({j},{OH["frank"]}${R0_}:{OH["frank"]}${R1},0))),'
               f'IFERROR(INDEX({EH["cab"]}${R0_}:{EH["cab"]}${R1},MATCH({j},{EH["frank"]}${R0_}:{EH["frank"]}${R1},0)),""))')
        vals = [f'=IF({FB["fname"]}{fr}="","",{j})', f'={FB["fname"]}{fr}', f'={FB["fckt"]}{fr}',
                f'={FB["fbrk"]}{fr}', f'={FB["fva"]}{fr}', f'={FB["fdem"]}{fr}', cab]
        for (a, b, t), v in zip(cols, vals):
            if a != b:
                merge(ws, f"{a}{r}:{b}{r}")
            set_cell(ws, f"{a}{r}", v, font(8), AL_C, "#,##0")
            border_range(ws, f"{a}{r}:{b}{r}", all_=thin_light)

    # ---- legend / how-to (left zone below circuits)
    set_cell(ws, f"A{t0}", "HOW TO USE THIS SHEET", font(9, True, C_LABEL), AL_L)
    tips = [
        "Yellow cells with blue text are inputs. Grey cells are calculated. Never type in the schedule (middle).",
        "Load (VA) = Fixture 1 Qty × VA + Fixture 2 Qty × VA + Added VA.",
        "Multi-pole: enter the breaker on its FIRST circuit only; the next slot(s) on the same side are taken.",
        "Sub-panel: pick a panel in 'Sub-Panel (feeds)' in an empty slot, then set Poles and Trip.",
        "    Through a transformer (e.g. 120/208 V from a 240 V delta)? Type XFMR in 'Via XFMR'.",
        "Cable (auto) is picked from the trip: 310.16 ampacity, ambient/bundling factors, 240.4(D), 250.122 EGC.",
        "    'Wire override' forces a size (pick from the list); parallel sets are worked out for you.",
        "High-leg delta: orange rows are the high leg (208 V to N). No 1-pole loads there; set Poles first.",
        "New panel: copy 'Template' (42 ckts) or 'Template 84', rename the tab, type the same name",
        "    in Panel Name, and add it to the Panels sheet (column B).",
        "Smaller panel: set No. of Circuits, then hide the unused circuit rows to shorten the printout.",
        "VA override: type the VA you want in 'VA override' (and why). Totals, amps, demand and cable follow;",
        "    the row shows WARN and the override is listed below. Clear the cell to go back to the calculation.",
        "Reset a row: select its yellow cells (Description to Override reason) and press Delete.",
        "Reset the panel: pick 'Circuit_Inputs' in the Name Box (left of the formula bar), press Delete.",
        "Working / print view: the 1 and 2 buttons above the column letters hide or show the input zones.",
        "Never type on the schedule itself: the tool flags typed-over cells (Panel checks, Circuit flags).",
        "Do not insert or delete rows/columns on panel sheets (cross-panel links use fixed cells).",
        "Status: ERR = must fix · WARN = check · (pole of CKT x) = slot used by a multi-pole breaker.",
    ]
    for i, t in enumerate(tips):
        set_cell(ws, f"B{t0+1+i}", t, font(8, False, "404040"), AL_L)

    # ---- overrides in effect (left zone, below the tips)
    o0 = t0 + len(tips) + 2
    set_cell(ws, f"A{o0}", "OVERRIDES IN EFFECT", font(9, True, "9C5700"), AL_L)
    ocols = [("A", "A", "CKT"), ("B", "C", "Description"), ("D", "D", "Calculated VA"), ("E", "I", "Override VA"),
             ("J", "U", "Reason")]
    for a_, b_, t in ocols:
        if a_ != b_:
            merge(ws, f"{a_}{o0+1}:{b_}{o0+1}")
        set_cell(ws, f"{a_}{o0+1}", t, font(8, True, C_LABEL), AL_C, fl=F_ZHDR)
        border_range(ws, f"{a_}{o0+1}:{b_}{o0+1}", all_=thin_light)
    def opick(oc, ec, j):
        return (f"IFERROR(INDEX({oc}${R0_}:{oc}${R1},MATCH({j},{OH['orank']}${R0_}:{OH['orank']}${R1},0)),"
                f"IFERROR(INDEX({ec}${R0_}:{ec}${R1},MATCH({j},{EH['orank']}${R0_}:{EH['orank']}${R1},0)),\"\"))")
    for j in range(1, 13):
        r = o0 + 1 + j
        vals = ["=" + opick("$" + ODD["ckt"], "$" + EVEN["ckt"], j),
                "=" + opick("$" + ODD["desc"], "$" + EVEN["desc"], j) + '&""',
                "=" + opick(OH["own0"], EH["own0"], j),
                "=" + opick("$" + ODD["ovr"], "$" + EVEN["ovr"], j),
                "=" + opick("$" + ODD["why"], "$" + EVEN["why"], j) + '&""']
        for (a_, b_, t), v in zip(ocols, vals):
            if a_ != b_:
                merge(ws, f"{a_}{r}:{b_}{r}")
            set_cell(ws, f"{a_}{r}", v, font(8), AL_L if a_ in ("B", "J") else AL_C, "#,##0")
            border_range(ws, f"{a_}{r}:{b_}{r}", all_=thin_light)

    # ================= COMPANY-LOOK SCHEDULE =================
    n12, b12 = font(12), font(12, True)
    n10, b10 = font(10), font(10, True)

    def box(rng, al=AL_C, ft=n12, value=None, fmt=None, border=blk):
        if ":" in rng:
            merge(ws, rng)
            first = rng.split(":")[0]
        else:
            first = rng
        set_cell(ws, first, value, ft, al, fmt)
        border_range(ws, rng if ":" in rng else f"{rng}:{rng}", all_=border)
        return ws[first]

    s = S
    # header block (company rows 1-7 -> rows 1-8)
    box(f"{s['ckto']}1:{s['desco']}1", value=f'=UPPER({pv("title")})')
    box(f"{s['ckto']}2:{s['desco']}4", ft=font(40), value=f'={pv("name")}')
    for i, (t, ref) in enumerate([("LOCATION:", pv("loc")), ("SUPPLY FROM:", pv("supply")),
                                  ("MOUNTING:", f'IF(LEN({IN["mount"]})=0,"",UPPER({IN["mount"]}))'),
                                  ("ENCLOSURE:", f'IF(LEN({IN["encl"]})=0,"",UPPER({IN["encl"]}))')]):
        r = 1 + i
        box(f"{s['tripo']}{r}", AL_CW if i == 1 else AL_L, font(11), t)
        box(f"{s['polo']}{r}:{s['wireo']}{r}", AL_CW, font(11), "=" + ref)
    box(f"{s['L1o']}1:{s['L3e']}1", ft=font(18, name="Bookman Old Style"), value=f'={pv("sysbox")}')
    box(f"{s['L1o']}2:{s['L3e']}2", ft=font(11), value=f'={pv("dist")}')
    box(f"{s['L1o']}3:{s['L3e']}4", AL_CW, font(10, True, "C65911"), f'={pv("hlnote")}')
    for i, (t, ref, fmt) in enumerate([("LINE TO LINE VOLTAGE:", pv("vll"), "0"),
                                       ("LINE TO NEUTRAL VOLTAGE:", pv("vln"), "0"),
                                       ("# PHASE:", pv("ph"), "0"),
                                       ("WIRE SIZE:", pv("fcab"), None)]):
        r = 1 + i
        box(f"{s['wiree']}{r}:{s['brke']}{r}", Alignment(horizontal="left", vertical="center", wrap_text=True),
            font(10), t)
        box(f"{s['pole']}{r}:{s['tripe']}{r}",
            Alignment(horizontal="center", vertical="center", shrink_to_fit=True) if i == 3 else AL_CW,
            font(9 if i == 3 else 12), "=" + ref, fmt)
    for i, (t, ref) in enumerate([("A.I.C RATING: ", pv("aic")), ("MAINS TYPE: ", pv("mtype")),
                                  ("MAINS RATING: ", f'IF(ISNUMBER({pv("mains")}),{pv("mains")}&"A","")'),
                                  ("MCB RATING: ", f'IF(ISNUMBER({pv("mcb")}),{pv("mcb")}&"A"'
                                                   f'&IF({pv("brkmode")}>=1," (100%)",""),"")')]):
        r = 1 + i
        box(f"{s['desce']}{r}:{s['ckte']}{r}", AL_L, font(11), f'="{t}"&{ref}')
    box(f"{s['ckto']}5:{s['wireo']}8", AL_WRAP_TL, font(10),
        f'=IF(LEN({IN["notes1"]})=0,"","NOTES: "&{IN["notes1"]})')
    box(f"{s['L1o']}5:{s['L3e']}8", Alignment(horizontal="center", vertical="top"), n12, "LOAD VALUES")
    border_range(ws, f"{s['L1o']}5:{s['L3e']}8", bottom=med)
    box(f"{s['wiree']}5:{s['brke']}8", Alignment(horizontal="left", vertical="center", wrap_text=True),
        font(10), "NUMBER OF CIRCUITS:")
    box(f"{s['pole']}5:{s['tripe']}8", AL_C, n12, f'={pv("slots")}', "0")
    box(f"{s['desce']}5:{s['ckte']}8", AL_CW, font(10),
        f'=IF({pv("fed")}="","","FED FROM "&{pv("fed")}&", CKT "&{pv("fedckt")}&" ("&{pv("fedbrk")}&")")')

    # column headers row 9
    for k, t in {"ckto": "CKT#", "desco": "CKT DESCRIPTION", "tripo": "TRIP", "polo": "POLES",
                 "brko": "BRK TYPE", "wireo": "WIRE SIZE", "wiree": "WIRE SIZE", "brke": "BRK TYPE",
                 "pole": "POLES", "tripe": "TRIP", "desce": "CKT DESCRIPTION", "ckte": "CKT#"}.items():
        box(f"{s[k]}9", AL_CW, b12, t)
    for n, (letter, (a, z)) in enumerate(PHASE_COLS.items(), 1):
        box(f"{s[a]}9:{s[z]}9", AL_CW, font(12 if True else 12, True),
            f'="L-{n} VA"&IF({pv("hl")}="{letter}"," (HIGH LEG)","")')
    border_range(ws, f"{s['L1o']}9:{s['L3e']}9", top=med)

    # circuit rows
    VA_AL = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    for r in range(R0_, R1 + 1):
        box(f"{s['ckto']}{r}", AL_C, b12, f'=IF(${ODD["ckt"]}{r}<={pv("slots")},${ODD["ckt"]}{r},"")')
        box(f"{s['ckte']}{r}", AL_C, b12, f'=IF(${EVEN["ckt"]}{r}<={pv("slots")},${EVEN["ckt"]}{r},"")')
        for Z, H, dk, tk, pk, bk, wk in [(ODD, OH, "desco", "tripo", "polo", "brko", "wireo"),
                                         (EVEN, EH, "desce", "tripe", "pole", "brke", "wiree")]:
            def txt(rr):
                d, tr, po = f"${Z['desc']}{rr}", f"${Z['trip']}{rr}", f"${Z['pol']}{rr}"
                return {
                    "d": (f'IF(LEN(TRIM({d}))>0,UPPER({d}),IF({H["valid"]}{rr}=1,{H["sub"]}{rr},'
                          f'IF(AND({H["has"]}{rr}=0,${Z["ckt"]}{rr}<={pv("slots")}),{pv("empty")},"")))'),
                    "t": f'IF(ISNUMBER({tr}),{tr},"")',
                    "p": f'IF(OR(ISNUMBER({tr}),{H["VA"]}{rr}>0,{H["valid"]}{rr}=1,ISNUMBER({po})),{H["P"]}{rr},"")',
                    "b": f'UPPER(TRIM(${Z["brk"]}{rr}&""))',
                    "w": f'{H["wlab"]}{rr}'}
            here, up = txt(r), (txt(r - 1) if r > R0_ else None)
            src = f'{H["dsrc"]}{r}'
            pick = lambda k: (f'=IF({src}=0,{here[k]},IF({src}=1,{up[k]},""))' if up
                              else f'=IF({src}=0,{here[k]},"")')
            NOSHRINK = Alignment(horizontal="center", vertical="center")
            box(f"{s[dk]}{r}", Alignment(horizontal="center", vertical="center", shrink_to_fit=True), n12, pick("d"))
            box(f"{s[tk]}{r}", NOSHRINK, n12, pick("t"), '0"A"')
            box(f"{s[pk]}{r}", NOSHRINK, n12, pick("p"), "0")
            box(f"{s[bk]}{r}", Alignment(horizontal="center", vertical="center", shrink_to_fit=True), font(10), pick("b"))
            box(f"{s[wk]}{r}", Alignment(horizontal="center", vertical="center", shrink_to_fit=True), font(10), pick("w"))
        for letter, (a, z) in PHASE_COLS.items():
            for col, H in [(s[a], OH), (s[z], EH)]:
                box(f"{col}{r}", VA_AL, n12,
                    f'=IF(AND(${ROWPH}{r}="{letter}",{H["slot"]}{r}<>0),{H["slot"]}{r},"")', "0")
    for keys in (("desco", "tripo", "polo", "brko", "wireo"), ("wiree", "brke", "pole", "tripe", "desce")):
        for k in keys:
            for r in range(R0_, R1 + 1):
                c = ws[f"{s[k]}{r}"]
                c.border = Border(left=blk, right=blk, top=blk if r == R0_ else None, bottom=blk if r == R1 else None)
    for g in ("g1", "g2"):
        border_range(ws, f"{s[g]}9:{s[g]}{R1}", all_=none)
        border_range(ws, f"{s[g]}9:{s[g]}{R1}", left=blk, right=blk)

    # totals (company rows 39-41)
    T = L.TOT
    for i, t in enumerate(["TOTAL LOAD VA:", "TOTAL LOAD VA:", "TOTAL AMPS:"]):
        box(f"{s['polo']}{T+i}:{s['wireo']}{T+i}", AL_L, n12, t)
    for n, (letter, (a, z)) in enumerate(PHASE_COLS.items()):
        act = ["TRUE", f'{pv("hot")}>=2', f'{pv("hot")}=3'][n]
        for col in (s[a], s[z]):
            box(f"{col}{T}", VA_AL, n12, f'=IF({act},SUM({col}{R0_}:{col}{R1}),"")', "0")
        box(f"{s[a]}{T+1}:{s[z]}{T+1}", VA_AL, n12, f'=IF({act},{s[a]}{T}+{s[z]}{T},"")', "0")
        box(f"{s[a]}{T+2}:{s[z]}{T+2}", VA_AL, b12,
            f'=IF({act},{pv(["aA", "aB", "aC"][n])},"")', "0.0")
    border_range(ws, f"{s['L1o']}{T+2}:{s['L3e']}{T+2}", bottom=med)
    border_range(ws, f"{s['ckto']}{T}:{s['tripo']}{T+2}", all_=none)
    border_range(ws, f"{s['wiree']}{T}:{s['ckte']}{T+2}", all_=none)

    # load classification + totals (company rows 42-47)
    B = L.BOT
    box(f"{s['ckto']}{B}", AL_C, n12, "#")
    box(f"{s['desco']}{B}", AL_L, n12, "LOAD CLASSIFICATION")
    box(f"{s['tripo']}{B}:{s['g1']}{B}", AL_C, n12, "CONNECTED LOADS")
    box(f"{s['L1o']}{B}:{s['L1e']}{B}", AL_CW, n12, "DEMAND FACTOR")
    box(f"{s['L2o']}{B}:{s['L2e']}{B}", AL_CW, n12, "ESTIMATED DEMAND")
    border_range(ws, f"{s['L1o']}{B}:{s['L2e']}{B}", top=med)
    box(f"{s['L3o']}{B}:{s['tripe']}{L.LAST}", AL_WRAP_TL, n10,
        f'="NOTES:"&CHAR(10)&{pv("basis")}&CHAR(10)&{pv("brknote")}'
        f'&IF(LEN({IN["notes2"]})=0,"",CHAR(10)&{IN["notes2"]})')
    ordr = f"{CB['ord']}${CLS_R0}:{CB['ord']}${CLS_R0+NCLS-1}"
    cr = lambda key: f"{CB[key]}${CLS_R0}:{CB[key]}${CLS_R0+NCLS-1}"
    for j in range(1, SCH_CLS_ROWS + 1):
        r = B + j
        mt = f"MATCH({j},{ordr},0)"
        q = f"${s['desco']}{r}"
        box(f"{s['ckto']}{r}", AL_C, n12, f'=IF({q}="","",{j})')
        box(f"{s['desco']}{r}", AL_L, n12, f'=IFERROR(UPPER(INDEX({cr("cname")},{mt})),"")')
        box(f"{s['tripo']}{r}:{s['g1']}{r}", AL_C, n12, f'=IF({q}="","",INDEX({cr("conn")},{mt}))', "#,##0")
        box(f"{s['L1o']}{r}:{s['L1e']}{r}", AL_C, n12, f'=IF({q}="","",INDEX({cr("df")},{mt}))', "0.00%")
        box(f"{s['L2o']}{r}:{s['L2e']}{r}", AL_C, n12, f'=IF({q}="","",INDEX({cr("dem")},{mt}))', "#,##0")
    right = [("TOTAL CONNECTED LOADS:", pv("conn"), '#,##0" VA"'),
             ("TOTAL ESTIMATED DEMANDS:", pv("dem"), '#,##0" VA"'),
             ("TOTAL CONNECTED AMPS:", pv("amax"), '0.0" A"'),
             ("TOTAL EST. DEMAND AMPS:", pv("demA"), '0.0" A"'),
             ("LOADING (% MCB / BUS):", f'IF(ISNUMBER({pv("loading")}),{pv("loading")},"")', "0%"),
             ("PHASE IMBALANCE:", f'IF({pv("hot")}>=2,{pv("imb")},"")', "0%"),
             ("MCA:", pv("mca"), '0" A"')]
    for i in range(SCH_CLS_ROWS + 1):
        r = B + i
        t, ref, fmt = right[i] if i < len(right) else (None, None, None)
        box(f"{s['desce']}{r}", AL_R, font(11), t)
        box(f"{s['ckte']}{r}", Alignment(horizontal="right", vertical="center", shrink_to_fit=True), font(11),
            ("=" + ref) if ref else None, fmt)
    ws[f"{s['desce']}{B+6}"].comment = Comment(
        "As in the v005 sheets: the next standard breaker rating at or above the estimated "
        "demand amps (max line current x demand factor). Check feeder OCPD per NEC 215.3 "
        "(125% of continuous load) separately.", "Panel tool")

    # outer frame + white fill
    border_range(ws, f"{SCH_FIRST}1:{SCH_LAST}{L.LAST}", top=med, bottom=med, left=med, right=med)
    for row in ws.iter_rows(min_row=1, max_row=L.LAST, min_col=CI(SCH_FIRST), max_col=CI(SCH_LAST)):
        for c in row:
            if c.font is None or c.font.name not in (FONT, "Bookman Old Style"):
                c.font = n12
            c.fill = F_WHITE

    # ================= CONDITIONAL FORMATTING =================
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.formatting.rule import Rule
    for keys, H in ((("desco", "tripo", "polo", "brko", "wireo"), OH), (("wiree", "brke", "pole", "tripe", "desce"), EH)):
        rng = " ".join(f"{s[k]}{R0_ + 1}:{s[k]}{R1}" for k in keys)
        dx = DifferentialStyle(border=Border(top=blk, left=blk, right=blk))
        rule = Rule(type="expression", dxf=dx, formula=[f'${H["cont"]}{R0_ + 1}=""'])
        ws.conditional_formatting.add(rng, rule)
    for letter, (a, z) in PHASE_COLS.items():
        ws.conditional_formatting.add(f"{s[a]}{R0_}:{s[z]}{R1}",
                                      FormulaRule(formula=[f'AND(${ROWPH}{R0_}="{letter}",${ODD["ckt"]}{R0_}<={pv("slots")})'],
                                                  fill=fill(F_PH)))
    for Z, H in [(ODD, OH), (EVEN, EH)]:
        rng = f"{Z['ckt']}{R0_}:{Z['stat']}{R1}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"${Z['ckt']}{R0_}>{pv('slots')}"],
                                                       fill=fill("BFBFBF"), font=font(9, False, "808080"), stopIfTrue=True))
        ws.conditional_formatting.add(f"{Z['desc']}{R0_}:{Z['wov']}{R1}",
                                      FormulaRule(formula=[f'AND(${H["cont"]}{R0_}<>"",${H["has"]}{R0_}=0)'],
                                                  fill=fill("D9D9D9"), stopIfTrue=True))
        ws.conditional_formatting.add(f"{Z['ckt']}{R0_}:{Z['wire']}{R1}",
                                      FormulaRule(formula=[f'AND({pv("hl")}<>"",${ROWPH}{R0_}={pv("hl")})'],
                                                  fill=fill(F_HL)))
        st = f"{Z['stat']}{R0_}:{Z['stat']}{R1}"
        top = f"{Z['stat']}{R0_}"
        ws.conditional_formatting.add(st, FormulaRule(formula=[f'LEFT({top},3)="ERR"'], fill=fill("FDE2E1"), font=font(8, True, "C00000")))
        ws.conditional_formatting.add(st, FormulaRule(formula=[f'LEFT({top},4)="WARN"'], fill=fill("FFEB9C"), font=font(8, False, "9C5700")))
        ws.conditional_formatting.add(st, FormulaRule(formula=[f'{top}="OK"'], font=font(8, False, "2E7D32")))
        ws.conditional_formatting.add(st, FormulaRule(formula=[f'LEFT({top},1)="("'], font=font(8, False, "808080", italic=True)))
    fl = f"{EVEN['stat']}2:{EVEN['stat']}8"
    ws.conditional_formatting.add(fl, FormulaRule(formula=[f'{EVEN["stat"]}2="ERR"'], fill=fill("FDE2E1"), font=font(8, True, "C00000")))
    ws.conditional_formatting.add(fl, FormulaRule(formula=[f'{EVEN["stat"]}2="WARN"'], fill=fill("FFEB9C"), font=font(8, True, "9C5700")))
    ws.conditional_formatting.add(fl, FormulaRule(formula=[f'{EVEN["stat"]}2="OK"'], fill=fill("E2EFDA"), font=font(8, True, "2E7D32")))

    # ================= DATA VALIDATION =================
    def dv(formula, rng, style="stop", title=None, msg=None, typ="list", op=None, f2=None):
        d = DataValidation(type=typ, formula1=formula, formula2=f2, operator=op, allow_blank=True,
                           errorStyle=style, showErrorMessage=True)
        if title:
            d.errorTitle, d.error = title, msg
        ws.add_data_validation(d)
        d.add(rng)
    dv("LK_Titles", IN["title"], "warning")
    dv("LK_Mounting", IN["mount"], "warning")
    dv("LK_Enclosure", IN["encl"], "warning")
    dv("LK_Slots", IN["slots"], "stop", "No. of circuits", f"Pick from the list. This sheet holds up to {cap}.")
    dv("DD_Dist", IN["dist"], "stop", "Distribution system", "Pick a system from Lookups (add new ones there).")
    dv("LK_AIC", IN["aic"], "warning")
    dv("LK_MainsType", IN["mtype"], "warning")
    dv("LK_Bus", IN["mains"], "warning", "Bus rating", "Not a standard bus rating. Keep it?")
    dv("LK_Breakers", IN["mcb"], "warning", "MCB rating", "Not a standard breaker rating. Keep it?")
    dv('"A,B,C"', IN["hlin"], "stop", "High-leg phase",
       "A, B or C. NEC 408.3(E)(1) puts the high leg on B in panelboards (C only in meter equipment).")
    dv('"Cu 75C,Cu 90C,Al 75C,Al 90C"', IN["mat"], "stop", "Conductor",
       "Cu 75C (any NEC building wire: THWN, THHN/THWN-2, XHHW-2) is the default. "
       "Pick 90C only when 90 °C wire (THHN/THWN-2, XHHW-2) is specified; it is printed on the schedule.")
    dv('"Standard 80%,MCB 100%,All 100%"', IN["brk100"], "stop", "Breaker rating",
       "Blank = Standard 80%. Use 100% only where the breaker AND its assembly are listed for 100% operation.")
    dv('"60,75,90"', IN["term"], "stop", "Termination rating", "60, 75 or 90 °C (NEC 110.14(C)).")
    dv("0", IN["amb"], "stop", "Ambient", "Enter °C between 0 and 85.", typ="decimal", op="between", f2="85")
    dv("1", IN["ccc"], "stop", "Conductors per raceway", "Whole number ≥ 1.", typ="whole", op="greaterThanOrEqual")
    for Z in (ODD, EVEN):
        q = f"{Z['q1']}{R0_}"
        hlok = (f'OR({pv("hl")}="",${ROWPH}{R0_}<>{pv("hl")},N(${Z["pol"]}{R0_})>=2,'
                f'LEN(TRIM(${Z["sub"]}{R0_}))>0)')
        dv(f"AND(ISNUMBER({q}),{q}>=0,{hlok})", f"{Z['q1']}{R0_}:{Z['add']}{R1}", "stop",
           "Load entry",
           "Enter a number ≥ 0. On an orange (high-leg) row set Poles to 2 or 3 FIRST: "
           "1-pole loads are not allowed on the high leg (208 V to neutral).", typ="custom")
        dv("DD_Class", f"{Z['cls']}{R0_}:{Z['cls']}{R1}", "stop", "Load classification",
           "Pick a class from the list (add new classes on the Lookups sheet).")
        dv("DD_Panels", f"{Z['sub']}{R0_}:{Z['sub']}{R1}", "warning", "Sub-panel",
           "That panel is not on the Panels sheet. It will still work if a tab with this name exists.")
        dv('"XFMR"', f"{Z['xf']}{R0_}:{Z['xf']}{R1}", "stop", "Via transformer",
           "Type XFMR when this feeder supplies the sub-panel through a transformer.")
        dv('"1,2,3"', f"{Z['pol']}{R0_}:{Z['pol']}{R1}", "stop", "Poles", "1, 2 or 3.")
        dv("LK_Breakers", f"{Z['trip']}{R0_}:{Z['trip']}{R1}", "warning", "Trip rating", "Not a standard trip rating. Keep it?")
        dv("LK_BrkType", f"{Z['brk']}{R0_}:{Z['brk']}{R1}", "warning", "Breaker type", "Not in the Lookups list. Keep it?")
        dv("LK_WSize", f"{Z['wov']}{R0_}:{Z['wov']}{R1}", "stop", "Wire size", "Pick a size from the list (AWG / kcmil).")
        dv("0", f"{Z['ovr']}{R0_}:{Z['ovr']}{R1}", "stop", "VA override", "Enter the VA to use (≥ 0), or clear the cell.",
           typ="decimal", op="greaterThanOrEqual")

    ws[IN["name"]].comment = Comment("Must be identical to this sheet's tab name. Sub-panel links and the "
                                     "Panels/Topology sheets find panels by tab name.", "Panel tool")
    ws[f"{ODD['sub']}9"].comment = Comment("Pick a panel to feed it from this circuit (use an empty slot). "
                                           "Its load comes in per phase and its load classes roll up into "
                                           "this panel's demand calculation.", "Panel tool")
    ws[f"{ODD['xf']}9"].comment = Comment("Type XFMR when the sub-panel is fed through a transformer "
                                          "(e.g. 120/208 V panel from a 240 V high-leg delta). Its VA is then "
                                          "spread evenly over this breaker's poles and its current is worked "
                                          "out at this panel's voltage.", "Panel tool")
    ws[f"{ODD['wov']}9"].comment = Comment("Leave blank to let the tool pick the conductor. Pick a size to "
                                           "force it (e.g. upsizing for voltage drop); parallel sets follow.",
                                           "Panel tool")
    ws[IN["brk100"]].comment = Comment("Blank = Standard 80%: breakers carry 80% continuously, so they and their "
                                       "cables are sized at 125% of continuous load (NEC 210.20(A), 215.3). "
                                       "MCB 100% / All 100%: only where the breaker and enclosure are listed for 100% "
                                       "operation (usually larger frames; check the manufacturer). For a single feeder, "
                                       "set its Breaker Type to '100% rated' instead.", "Panel tool")
    ws[IN["hlin"]].comment = Comment("Only used on high-leg delta systems. Blank = the Lookups default (B). "
                                     "NEC 408.3(E)(1) requires B in panelboards; C is allowed only in meter "
                                     "equipment.", "Panel tool")

    # ================= INPUT VALUES =================
    for ck, data in circuits.items():
        Z = ODD if ck % 2 == 1 else EVEN
        r = R0_ + (ck - 1) // 2
        assert r <= R1, (sheet_name, ck)
        for k, v in data.items():
            ws[f"{Z[k]}{r}"] = v

    if hide_unused and isinstance(hdr_vals.get("slots"), int):
        for r in range(R0_ + hdr_vals["slots"] // 2, R1 + 1):
            ws.row_dimensions[r].hidden = True

    # named ranges on this sheet (copy with the sheet): CAD Data Link + reset selections
    from openpyxl.workbook.defined_name import DefinedName
    q = f"'{sheet_name}'"
    def local(name, ref):
        ws.defined_names[name] = DefinedName(name, attr_text=ref)
    local("Schedule", f"{q}!${SCH_FIRST}$1:${SCH_LAST}${L.LAST}")
    local("Circuit_Inputs", f"{q}!${ODD['desc']}${R0_}:${ODD['why']}${R1},{q}!${EVEN['desc']}${R0_}:${EVEN['why']}${R1}")

    # working / print view: input zones on outline level 1 (the 1 / 2 buttons above the column letters)
    for ci in list(range(1, CI(ODD["stat"]) + 2)) + list(range(CI(EVEN["ckt"]), CI(EVEN["stat"]) + 2)):
        ws.column_dimensions[CL(ci)].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryRight = False

    ws.oddFooter.right.text = f"panelbook {VERSION}"
    ws.oddFooter.right.size = 7

    # print setup
    ws.print_area = f"{SCH_FIRST}1:{SCH_LAST}{L.LAST}"
    ws.page_setup.orientation = "landscape" if cap <= 42 else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.sheet_view.selection[0].activeCell = IN["name"]
    ws.sheet_view.selection[0].sqref = IN["name"]
    return ws
