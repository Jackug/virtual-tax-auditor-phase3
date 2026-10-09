import unittest
import app


class ScoringTests(unittest.TestCase):
    def test_bands(self):
        self.assertEqual(app._risk_band(10), "Low")
        self.assertEqual(app._risk_band(40), "Medium")
        self.assertEqual(app._risk_band(70), "High")
        self.assertEqual(app._risk_band(90), "Critical")

    def test_positive_exposure_adds_configured_points(self):
        rule = {"id": 1, "name": "Rule A",
                "ranking_config_json": '{"base_score":20,"exposure_weight":15,"frequency_weight":10}'}
        score_no_exposure, _ = app._calculate_taxpayer_risk_score([(rule, 0, 0, "A")])
        score_with_exposure, _ = app._calculate_taxpayer_risk_score([(rule, 100, 0, "A")])
        self.assertEqual(score_no_exposure, 20)
        self.assertEqual(score_with_exposure, 35)

    def test_multiple_rules_are_capped(self):
        rule_a = {"id": 1, "name": "A",
                  "ranking_config_json": '{"base_score":80,"exposure_weight":30,"frequency_weight":25}'}
        rule_b = {"id": 2, "name": "B",
                  "ranking_config_json": '{"base_score":70,"exposure_weight":20,"frequency_weight":25}'}
        score, details = app._calculate_taxpayer_risk_score([
            (rule_a, 100, 0, "A"), (rule_b, 200, 0, "B")
        ])
        self.assertEqual(score, 100)
        self.assertEqual(len(details), 2)


if __name__ == "__main__":
    unittest.main()
