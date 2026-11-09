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
}
