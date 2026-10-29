import json
import unittest
from pathlib import Path

from kkflow.parse_ordre import parse_ordrekvittering

ROD = Path(__file__).resolve().parent.parent
CFG = json.loads((ROD / "config.json").read_text(encoding="utf-8"))


class TestOrdrekvittering(unittest.TestCase):
    def test_ab_catering_31525653(self):
        o = parse_ordrekvittering(ROD / "eksempler" / "indbakke" / "ordrekvittering_31525653.html", CFG["kolonneregler"])
        self.assertEqual(o["ordre"], "31525653")
        self.assertEqual(o["leveringsdato"], "2026-10-27")
        self.assertEqual(o["kunde"], "39-69505002, FORENEDE SERVICE FRUGT")
        self.assertEqual(o["bruger"], "kaso@kokkeneskoekken.dk")
        self.assertEqual(o["leverandoer"], "AB Catering København")
        self.assertEqual(o["kolonne"], "Frugt")
        self.assertEqual(len(o["linjer"]), 5)
        self.assertEqual(o["linjer"][0], {"varenr": "224918", "vare": "Æble Ingrid Marie DK Øko", "antal": 20.0,
                                          "enhed": "STK", "enhedspris": 3.42, "linjepris": 68.40})
        self.assertAlmostEqual(o["total"], 386.23)
        self.assertEqual(o["advarsler"], [])

    def test_ren_tekst(self):
        sti = ROD / "tests" / "_tmp_ordre.txt"
        sti.write_text("Ordre: 123456\nLeveringsdato: 27/10/2026\nKunde: 39-1, X\n"
                       "224918 Æble Ingrid Marie DK Øko 20 STK 3,42 68,40\nTotal 68,40\nAB Catering København",
                       encoding="utf-8")
        try:
            o = parse_ordrekvittering(sti, CFG["kolonneregler"])
        finally:
            sti.unlink()
        self.assertEqual(o["ordre"], "123456")
        self.assertEqual(o["linjer"][0]["vare"], "Æble Ingrid Marie DK Øko")
        self.assertEqual(o["kolonne"], "BC catering")
        self.assertEqual(o["advarsler"], [])


if __name__ == "__main__":
    unittest.main()
