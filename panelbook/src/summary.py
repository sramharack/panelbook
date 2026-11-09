"""Panels register/summary sheet and Topology tree sheet."""
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from common import *


def ind(panel_ref, cell):
    """INDIRECT into a panel tab by name."""
    return f'INDIRECT("\'"&{panel_ref}&"\'!{cell}")'


def build_panels(wb, names):
    ws = wb.create_sheet("Panels")
    ws.sheet_properties.tabColor = "2E7D32"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "PANEL REGISTER & SUMMARY"
    ws["A1"].font = font(14, True, C_LABEL)
    ws["A2"] = ("Type each panel's TAB name in column B (one row per panel; no gaps). Everything else is calculated from "
                "the panel sheets. 'Fed From' comes from the Sub-Panel assignments on the upstream panel.")
    ws["A2"].font = font(9, False, "595959", italic=True)
    heads = {"idx": "#", "panel": "Panel\n(tab name)", "check": "Sheet check", "type": "Type",
             "loc": "Location", "fed": "Fed From", "fckt": "Feeder CKT", "fbrk": "Feeder Breaker",
             "sys": "System", "bus": "Bus (A)", "mcb": "MCB (A)", "connk": "Conn.\n(kVA)",
             "demk": "Est. Demand\n(kVA)", "demA": "Est. Demand\n(A, max line)", "loading": "Loading\n(% MCB/bus)",
             "aA": "Line A (A)", "aB": "Line B (A)", "aC": "Line C (A)", "imb": "Imbalance", "level": "Level",
             "err": "ERR", "warn": "WARN", "flags": "Flags"}
    widths = {"idx": 4, "panel": 12, "check": 11, "type": 16, "loc": 20, "fed": 11, "fckt": 9,
              "fbrk": 18, "sys": 28, "bus": 7, "mcb": 7, "connk": 8, "demk": 10, "demA": 10,
              "loading": 10, "aA": 7, "aB": 7, "aC": 7, "imb": 9, "level": 6, "err": 5, "warn": 6, "flags": 34}
    ws.row_dimensions[4].height = 32
    for k, col in RV.items():
        c = ws[f"{col}4"]
        c.value = heads[k]
        c.font = font(8, True, C_LABEL)
        c.fill = F_ZHDR
        c.alignment = AL_HDR
        c.border = GRID_LIGHT
        ws.column_dimensions[col].width = widths[k]
    ws.column_dimensions["X"].width = 2

    fm = f"${FM[0]}${REG_R0}:${FM[-1]}${REG_R1}"
    Bc = f"$B${REG_R0}:$B${REG_R1}"
    for i in range(REG_R1 - REG_R0 + 1):
        r = REG_R0 + i
        B = f"${RV['panel']}{r}"
        Bt = f"TRIM({B})"
        ok = f"${RH['ok']}{r}"
        V = {k: f"${col}{r}" for k, col in RV.items()}
        def g(key):
            return ind(Bt, f"${PV}${P[key]}")
        ws[V["idx"]] = f'=IF({B}="","",{i+1})'
        if i < len(names):
            ws[V["panel"]] = names[i]
        ws[V["check"]] = (f'=IF({B}="","",IF(ISERROR({g("name")}),"No such tab",'
                          f'IF({g("name")}<>{Bt},"Panel Name ≠ tab","OK")))')
        ws[f"{RH['ok']}{r}"] = f'=IF({B}="",0,IF({V["check"]}="No such tab",0,1))'
        ws[V["type"]] = f'=IF({ok}=1,{g("title")},"")'
        ws[V["loc"]] = f'=IF({ok}=1,{g("loc")},"")'
        ws[f"{RH['manual']}{r}"] = f'=IF({ok}=1,T({g("manual")}),"")'
        for j in range(NFEED):
            ws[f"{FM[j]}{r}"] = f'=IF({ok}=1,T({ind(Bt, f"${FB["fname"]}${FEED_R0+j}")}),"")'
        ws[f"{RH['cnt']}{r}"] = f'=IF({B}="",0,COUNTIF({fm},{Bt}))'
        ws[f"{RH['prow']}{r}"] = (f'=IF({RH["cnt"]}{r}=1,SUMPRODUCT(({fm}={Bt})*(ROW({fm})-{REG_R0-1})),0)')
        prow = f"{RH['prow']}{r}"
        ws[V["fed"]] = (f'=IF({B}="","",IF({RH["cnt"]}{r}=0,"",IF({RH["cnt"]}{r}>1,"(multiple)",'
                        f'INDEX({Bc},{prow}))))')
        pos = f"MATCH({Bt},INDEX({fm},{prow},0),0)"
        ws[V["fckt"]] = (f'=IF({prow}>0,IFERROR(INDIRECT("\'"&{V["fed"]}&"\'!${FB["fckt"]}$"&({FEED_R0-1}+{pos})),""),"")')
        ws[V["fbrk"]] = (f'=IF({prow}>0,IFERROR(INDIRECT("\'"&{V["fed"]}&"\'!${FB["fbrk"]}$"&({FEED_R0-1}+{pos})),""),"")')
        ws[V["sys"]] = f'=IF({ok}=1,{g("dist")},"")'
        ws[V["bus"]] = f'=IF({ok}=1,{g("mains")},"")'
        ws[V["mcb"]] = f'=IF({ok}=1,{g("mcb")},"")'
        ws[V["connk"]] = f'=IF({ok}=1,N({g("conn")})/1000,"")'
        ws[V["demk"]] = f'=IF({ok}=1,N({g("dem")})/1000,"")'
        ws[V["demA"]] = f'=IF({ok}=1,N({g("demA")}),"")'
        ws[V["loading"]] = f'=IF({ok}=1,{g("loading")},"")'
        ws[V["aA"]] = f'=IF({ok}=1,N({g("aA")}),"")'
        ws[V["aB"]] = f'=IF({ok}=1,IF(N({g("hot")})>=2,N({g("aB")}),""),"")'
        ws[V["aC"]] = f'=IF({ok}=1,IF(N({g("hot")})=3,N({g("aC")}),""),"")'
        ws[V["imb"]] = f'=IF({ok}=1,IF(N({g("hot")})>=2,N({g("imb")}),""),"")'
        ws[V["err"]] = f'=IF({ok}=1,N({g("err")}),"")'
        ws[V["warn"]] = f'=IF({ok}=1,N({g("warn")}),"")'
        # parent chain
        ws[f"{PAR[0]}{r}"] = f'=IF({prow}>0,{V["fed"]},"")'
        p1c = f"${PAR[0]}${REG_R0}:${PAR[0]}${REG_R1}"
        for j in range(1, 8):
            prev = f"{PAR[j-1]}{r}"
            ws[f"{PAR[j]}{r}"] = f'=IF({prev}="","",IFERROR(INDEX({p1c},MATCH({prev},{Bc},0)),""))'
        ws[V["level"]] = f'=IF({B}="","",' + "+".join(f'({PAR[j]}{r}<>"")' for j in range(8)) + ")"
        for j in range(7):
            ws[f"{IDXP[j]}{r}"] = f'=IF({PAR[j]}{r}="",0,IFERROR(MATCH({PAR[j]}{r},{Bc},0),0))'
        ws[f"{RH['idx']}{r}"] = f'=IF({B}="","",{i+1})'
        L = V["level"]
        terms = [f"{RH['idx']}{r}*51^(7-{L})"] + [f"{IDXP[j]}{r}*51^({8+j}-{L})" for j in range(7)]
        ws[f"{RH['key']}{r}"] = f'=IF(OR({B}="",N({L})>7),"",' + "+".join(terms) + ")"
        kc = f"${RH['key']}${REG_R0}:${RH['key']}${REG_R1}"
        key = f"{RH['key']}{r}"
        ws[f"{RH['rank']}{r}"] = f'=IF({key}="","",SUMPRODUCT(ISNUMBER({kc})*({kc}<{key}))+1)'
        ws[f"{RH['islast']}{r}"] = (f'=IF({key}="",1,IF({PAR[0]}{r}="",1,IF(SUMPRODUCT(({p1c}={PAR[0]}{r})'
                                    f'*ISNUMBER({kc})*({kc}>{key}))=0,1,0)))')
        ilc = f"${RH['islast']}${REG_R0}:${RH['islast']}${REG_R1}"
        for j in range(7):
            ws[f"{ISLP[j]}{r}"] = f'=IF({PAR[j]}{r}="",1,IFERROR(INDEX({ilc},MATCH({PAR[j]}{r},{Bc},0)),1))'
        segs = "&".join(f'IF(N({L})-1>={j+1},IF({ISLP[j]}{r}=1,"   ","│  "),"")' for j in reversed(range(7)))
        ws[f"{RH['tree']}{r}"] = (f'=IF({B}="","",IF({key}="",{Bt}&"  (feed loop?)",{segs}'
                                  f'&IF(N({L})>=1,IF({RH["islast"]}{r}=1,"└─ ","├─ "),"")&{Bt}))')
        fs = (f'IF({V["check"]}="Panel Name ≠ tab"," · Panel Name ≠ tab","")'
              f'&IF({RH["cnt"]}{r}>1," · Fed from 2+ panels","")'
              f'&IF({PAR[7]}{r}<>""," · Feed loop","")'
              f'&IF(LEFT({RH["sysx"]}{r},1)="⚠"," · Voltage change, no XFMR","")'
              f'&IF(N({V["err"]})>0," · "&{V["err"]}&" ERR","")'
              f'&IF(N({V["warn"]})>0," · "&{V["warn"]}&" WARN","")'
              f'&IF(AND(ISNUMBER({V["bus"]}),ISNUMBER({V["mcb"]})),IF(N({V["mcb"]})>N({V["bus"]})," · MCB > bus",""),"")'
              f'&IF(ISNUMBER({V["loading"]}),IF(N({V["loading"]})>1," · Overloaded",IF(N({V["loading"]})>LK_MCBWarn," · Loading "&TEXT({V["loading"]},"0%"),"")),"")'
              f'&IF(ISNUMBER({V["imb"]}),IF(N({V["imb"]})>LK_ImbLimit," · Imbalance "&TEXT({V["imb"]},"0%"),""),"")')
        ws[f"{RH['flagstr']}{r}"] = f'=IF({ok}=1,{fs},"")'
        # voltage change across the feeder (transformer-fed sub-panels)
        psys = f'IFERROR(INDEX(${RV["sys"]}${REG_R0}:${RV["sys"]}${REG_R1},MATCH({V["fed"]},{Bc},0))&"","")'
        ws[f"{RH['sysx']}{r}"] = (f'=IF(OR({ok}<>1,{prow}=0),"",IF(ISNUMBER(SEARCH("XFMR",{V["fbrk"]})),'
                                  f'"XFMR: "&{psys}&" → "&{V["sys"]},IF(OR(N(INDIRECT("\'"&{V["fed"]}&"\'!${PV}${P["vll"]}"))<>N({g("vll")}),N(INDIRECT("\'"&{V["fed"]}&"\'!${PV}${P["vln"]}"))<>N({g("vln")})),'
                                  f'"⚠ "&{psys}&" → "&{V["sys"]}&" (no XFMR marked)","")))')
        ws[V["flags"]] = (f'=IF({B}="","",IF({ok}=0,{V["check"]},IF({RH["flagstr"]}{r}="","OK",'
                          f'MID({RH["flagstr"]}{r},4,200))))')
        # styles
        for k, col in RV.items():
            c = ws[f"{col}{r}"]
            c.border = GRID_LIGHT
            c.font = font(9, k == "panel", C_INPUT if k == "panel" else "000000")
            c.fill = F_INPUT if k == "panel" else F_WHITE
            c.alignment = AL_L if k in ("panel", "type", "loc", "fed", "sys", "flags", "fbrk") else AL_C
        for k, fmt in [("connk", "#,##0.0"), ("demk", "#,##0.0"), ("demA", "#,##0"), ("loading", "0%"),
                       ("aA", "0"), ("aB", "0"), ("aC", "0"), ("imb", "0%"), ("bus", "0"), ("mcb", "0")]:
            ws[f"{RV[k]}{r}"].number_format = fmt
        for col in FM + list(RH.values()) + PAR + IDXP + ISLP:
            ws[f"{col}{r}"].font = font(8, False, "7F7F7F")
    # helper headers + hide
    for j, col in enumerate(FM):
        ws[f"{col}4"] = f"feeds{j+1}"
    for k, col in RH.items():
        ws[f"{col}4"] = k
    for j, col in enumerate(PAR):
        ws[f"{col}4"] = f"parent{j+1}"
    for j, col in enumerate(IDXP):
        ws[f"{col}4"] = f"idxP{j+1}"
    for j, col in enumerate(ISLP):
        ws[f"{col}4"] = f"lastP{j+1}"
    for ci in range(CI(FM[0]), CI(REG_LAST) + 1):
        ws.column_dimensions[CL(ci)].hidden = True
        ws[f"{CL(ci)}4"].font = font(8, False, "7F7F7F")

    # CF
    rr = f"{REG_R0}:{REG_R1}"
    flags = f"{RV['flags']}{REG_R0}:{RV['flags']}{REG_R1}"
    ws.conditional_formatting.add(flags, FormulaRule(formula=[f'AND({RV["flags"]}{REG_R0}<>"",{RV["flags"]}{REG_R0}<>"OK")'],
                                                     fill=fill("FFEB9C"), font=font(9, False, "9C5700")))
    ws.conditional_formatting.add(flags, FormulaRule(formula=[f'{RV["flags"]}{REG_R0}="OK"'], font=font(9, False, "2E7D32")))
    chk = f"{RV['check']}{REG_R0}:{RV['check']}{REG_R1}"
    ws.conditional_formatting.add(chk, FormulaRule(formula=[f'AND({RV["check"]}{REG_R0}<>"",{RV["check"]}{REG_R0}<>"OK")'],
                                                   fill=fill("FDE2E1"), font=font(9, True, "C00000")))
    ld = f"{RV['loading']}{REG_R0}:{RV['loading']}{REG_R1}"
    ws.conditional_formatting.add(ld, FormulaRule(formula=[f'AND(ISNUMBER({RV["loading"]}{REG_R0}),{RV["loading"]}{REG_R0}>1)'],
                                                  fill=fill("FDE2E1"), font=font(9, True, "C00000")))
    ws.conditional_formatting.add(ld, FormulaRule(formula=[f'AND(ISNUMBER({RV["loading"]}{REG_R0}),{RV["loading"]}{REG_R0}>LK_MCBWarn)'],
                                                  fill=fill("FFEB9C")))
    ws.freeze_panes = f"C{REG_R0}"

    wb.defined_names["LK_PanelNames"] = DefinedName("LK_PanelNames", attr_text=f"Panels!$B${REG_R0}:$B${REG_R1}")
    wb.defined_names["DD_Panels"] = DefinedName(
        "DD_Panels", attr_text=f"OFFSET(Panels!$B${REG_R0},0,0,MAX(1,COUNTA(Panels!$B${REG_R0}:$B${REG_R1})),1)")
    ws.print_area = f"A1:{RV['flags']}{REG_R1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_topology(wb):
    ws = wb.create_sheet("Topology")
    ws.sheet_properties.tabColor = "2E7D32"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "DISTRIBUTION TOPOLOGY"
    ws["A1"].font = font(14, True, C_LABEL)
    ws["A2"] = ("Built automatically from the Sub-Panel assignments on each panel sheet (panels must be listed on the "
                "Panels sheet). Loads shown for each panel include everything downstream of it.")
    ws["A2"].font = font(9, False, "595959", italic=True)

    # KPIs
    def reg(k):
        return f"Panels!${RV[k]}${REG_R0}:${RV[k]}${REG_R1}"
    kpis = [
        ("Panels listed", f"=COUNTA(Panels!$B${REG_R0}:$B${REG_R1})", "0"),
        ("Top-level (service) panels", f"=COUNTIFS({reg('level')},0)", "0"),
        ("Deepest level", f"=MAX({reg('level')})", "0"),
        ("Est. demand at top level (kVA)", f"=SUMIFS({reg('demk')},{reg('level')},0)", "#,##0.0"),
        ("Connected at top level (kVA)", f"=SUMIFS({reg('connk')},{reg('level')},0)", "#,##0.0"),
        ("Panels with flags", f'=SUMPRODUCT((Panels!$B${REG_R0}:$B${REG_R1}<>"")*({reg("flags")}<>"OK"))', "0"),
    ]
    for i, (t, fml, fmt) in enumerate(kpis):
        col = CL(2 + 2 * (i % 3))
        r = 4 + (i // 3) * 2
        lab = ws[f"{col}{r}"]; lab.value = t; lab.font = font(8, False, "595959")
        v = ws[f"{col}{r+1}"]; v.value = fml; v.font = font(14, True, C_LABEL); v.number_format = fmt
        v.alignment = AL_L

    heads = ["#", "Distribution tree", "Fed from", "Feeder breaker", "System", "Bus / MCB",
             "Conn. (kVA)", "Est. Demand (kVA)", "Est. Demand (A, max line)", "Loading", "Imbalance", "Flags"]
    widths = [4, 34, 22, 18, 44, 13, 10, 12, 11, 16, 10, 36]
    HR = 9
    ws.row_dimensions[HR].height = 28
    for i, (h, w) in enumerate(zip(heads, widths)):
        col = CL(1 + i)
        c = ws[f"{col}{HR}"]
        c.value = h; c.font = font(8, True, C_LABEL); c.fill = F_ZHDR; c.alignment = AL_HDR; c.border = GRID_LIGHT
        ws.column_dimensions[col].width = w
    ws.column_dimensions["N"].hidden = True
    T0 = HR + 1
    for t in range(1, REG_R1 - REG_R0 + 2):
        r = T0 + t - 1
        ix = f"$N{r}"
        ws[f"N{r}"] = f'=IFERROR(MATCH({t},Panels!${RH["rank"]}${REG_R0}:${RH["rank"]}${REG_R1},0),"")'
        def pick(k):
            return f"INDEX({reg(k)},{ix})"
        man = f"INDEX(Panels!${RH['manual']}${REG_R0}:${RH['manual']}${REG_R1},{ix})"
        tree = f"INDEX(Panels!${RH['tree']}${REG_R0}:${RH['tree']}${REG_R1},{ix})"
        vals = [
            f'=IF({ix}="","",{t})',
            f'=IF({ix}="","",{tree})',
            f'=IF({ix}="","",IF({pick("fed")}<>"",{pick("fed")}&"  ·  CKT "&{pick("fckt")},IF({man}="","(source not set)","◄ "&{man})))',
            f'=IF({ix}="","",{pick("fbrk")}&"")',
            f'=IF({ix}="","",IF(INDEX(Panels!${RH["sysx"]}${REG_R0}:${RH["sysx"]}${REG_R1},{ix})<>"",INDEX(Panels!${RH["sysx"]}${REG_R0}:${RH["sysx"]}${REG_R1},{ix}),{pick("sys")}&""))',
            f'=IF({ix}="","",IF(ISNUMBER({pick("bus")}),{pick("bus")}&" A","–")&" / "&IF(ISNUMBER({pick("mcb")}),{pick("mcb")}&" A","MLO"))',
            f'=IF({ix}="","",{pick("connk")})',
            f'=IF({ix}="","",{pick("demk")})',
            f'=IF({ix}="","",{pick("demA")})',
            f'=IF({ix}="","",{pick("loading")})',
            f'=IF({ix}="","",{pick("imb")})',
            f'=IF({ix}="","",{pick("flags")})',
        ]
        fmts = [None, None, None, None, None, None, "#,##0.0", "#,##0.0", "#,##0", "0%", "0%", None]
        for i, (v, fm) in enumerate(zip(vals, fmts)):
            col = CL(1 + i)
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = font(10, i == 1, "1F3864", name="Consolas") if i == 1 else font(9)
            c.alignment = AL_L if i in (1, 2, 3, 4, 5, 11) else AL_C
            if fm:
                c.number_format = fm
            c.border = Border(bottom=Side(style="hair", color="D9D9D9"))
        ws.row_dimensions[r].height = 16
    last = T0 + REG_R1 - REG_R0
    ws.conditional_formatting.add(f"J{T0}:J{last}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                end_value=1, color="5B9BD5", showValue=True))
    ws.conditional_formatting.add(f"L{T0}:L{last}", FormulaRule(formula=[f'AND(L{T0}<>"",L{T0}<>"OK")'],
                                                                font=font(9, True, "C00000")))
    ws.conditional_formatting.add(f"L{T0}:L{last}", FormulaRule(formula=[f'L{T0}="OK"'], font=font(9, False, "2E7D32")))
    ws.freeze_panes = f"A{T0}"
    ws.print_area = f"A1:L{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws
