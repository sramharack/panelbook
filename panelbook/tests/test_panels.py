"""Extra panels used only by the tests (never shipped)."""
HLD = "120/240 3Ph 4W High-Leg Delta"

TEST_PANELS = {
    # Regression from the old-tool check: expected line amps A 35.27 / B 31.10 / C 21.10
    "HLTEST": (dict(name="HLTEST", title="Panelboard", loc="Regression", manual="Test", slots=42,
                    dist=HLD, mtype="MCB", mains=100, mcb=100),
               {1: dict(desc="1P on A", cls="Other", add=1000, trip=20),
                3: dict(desc="1P on B (must be flagged)", cls="Other", add=500, trip=20),
                5: dict(desc="1P on C", cls="Other", add=800, trip=20),
                2: dict(desc="2P on A-B", cls="Other", add=3000, pol=2, trip=20),
                7: dict(desc="3P chiller", cls="Other", add=6000, pol=3, trip=20)}, 42, False),
    # High leg moved to C, aluminium, 40 C, 6 CCC, sub-panel and override edge cases
    "E1": (dict(name="E1", slots=42, dist=HLD, mtype="MCB", mains=225, mcb=200, hlin="C", mat="Al", amb=40, ccc=6),
           {1: dict(sub="E2", pol=2, trip=60),
            2: dict(sub="E3", pol=3, trip=60),
            7: dict(desc="1P on A, C is HL", cls="Other", add=900, trip=20),
            11: dict(desc="1P on C = HL", cls="Other", add=900, trip=20),
            8: dict(desc="override 4/0 at 400A", cls="Other", add=1000, pol=3, trip=400, wov="4/0"),
            14: dict(desc="override #8, 150A", cls="Other", add=1000, trip=150, wov="8"),
            20: dict(desc="bad override", cls="Other", add=100, trip=20, wov="7")}, 42, False),
    "E2": (dict(name="E2", slots=24, dist="120/240 1Ph 3W", mains=100, mcb=60),
           {1: dict(desc="rec", cls="Receptacle", q1=5, v1=180, trip=20)}, 42, False),
    "E3": (dict(name="E3", slots=24, dist="120/208 Wye", mains=100, mcb=60),
           {1: dict(desc="rec", cls="Receptacle", q1=5, v1=180, trip=20)}, 42, False),
    "E4": (dict(name="E4", slots=12, dist="115/230 3Ph 4W High-Leg Delta", mains=100, mcb=100),
           {3: dict(desc="1P on B", cls="Other", add=500, trip=20)}, 42, False),
    "E5": (dict(name="E5", slots=12, dist="120/208 Wye", hlin="B", mains=100, mcb=100),
           {1: dict(desc="x", cls="Other", add=100, trip=20)}, 42, False),
    # Cable basis: #12 on 20 A at 35 C with 6 current-carrying conductors
    "E6": (dict(name="E6", slots=12, dist="120/208 Wye", mains=100, mcb=100, mat="Cu 75C", amb=35, ccc=6),
           {1: dict(desc="20 A, 75 C basis", cls="Other", add=1000, trip=20)}, 42, False),
    "E7": (dict(name="E7", slots=12, dist="120/208 Wye", mains=100, mcb=100, mat="Cu 90C", amb=35, ccc=6),
           {1: dict(desc="20 A, 90 C basis", cls="Other", add=1000, trip=20)}, 42, False),
    # Continuous load: test (a) 125% governs; noncontinuous twin for comparison
    "E8": (dict(name="E8", slots=12, dist="120/208 Wye", mains=100, mcb=100),
           {1: dict(desc="continuous 20.8 A on 20 A", cls="Lighting", add=2500, trip=20),
            3: dict(desc="noncontinuous 20.8 A on 20 A", cls="Other", add=2500, trip=20),
            5: dict(desc="override too small", cls="Lighting", add=1000, trip=30, wov="12"),
            7: dict(desc="continuous 16.7 A on 20 A", cls="Lighting", add=2000, trip=20)}, 42, False),
    # Breaker rating: same loads with all breakers 100% rated, plus a per-circuit 100% feeder
    "E9": (dict(name="E9", slots=12, dist="120/208 Wye", mains=100, mcb=100, brk100="All 100%"),
           {1: dict(desc="continuous 20.8 A on 20 A", cls="Lighting", add=2500, trip=20),
            7: dict(desc="continuous 16.7 A on 20 A", cls="Lighting", add=2000, trip=20)}, 42, False),
    # MCB check: 32.4 kVA continuous 3P at 208 V = 90 A demand on a 100 A MCB
    "E10": (dict(name="E10", slots=12, dist="120/208 Wye", mains=225, mcb=100),
            {1: dict(desc="continuous 90 A 3P", cls="Lighting", add=32400, pol=3, trip=125)}, 42, False),
    "E11": (dict(name="E11", slots=12, dist="120/208 Wye", mains=225, mcb=100, brk100="MCB 100%"),
            {1: dict(desc="continuous 90 A 3P", cls="Lighting", add=32400, pol=3, trip=125)}, 42, False),
    "E12": (dict(name="E12", slots=12, dist="120/208 Wye", mains=100, mcb=100),
            {2: dict(desc="100% rated feeder", cls="Lighting", add=1900, trip=15, brk="100% rated")}, 42, False),
    # Overrides: CKT 1 calculated 1,000 VA, overridden to 2,400 VA; CKT 3 3P display test
    "E13": (dict(name="E13", slots=12, dist="120/208 Wye", mains=100, mcb=100),
            {1: dict(desc="override me", cls="Other", add=1000, trip=20, ovr=1500, why="site survey"),
             3: dict(desc="three pole load", cls="Other", add=6000, pol=3, trip=30),
             2: dict(desc="two pole load", cls="Other", add=3000, pol=2, trip=20)}, 42, False),
    # 240.85 warning: only real 2P loads on the high leg without a straight-rated breaker
    "E14": (dict(name="E14", slots=24, dist=HLD, mains=100, mcb=100),
            {2: dict(desc="SPARE", pol=2, trip=20),                                   # A-B spare, no load
             10: dict(desc="2P load, slash breaker", cls="Other", add=2000, pol=2, trip=20),  # B-C
             14: dict(desc="2P load, straight 240V", cls="Other", add=2000, pol=2, trip=20,
                     brk="Straight 240V")}, 42, False),                               # A-B
}
