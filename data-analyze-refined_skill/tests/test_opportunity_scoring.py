import os
import sys
import unittest

import pandas as pd


SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from data_analyzer import DataAnalyzer  # noqa: E402


class TestOpportunityScoring(unittest.TestCase):
    def test_priority_scoring_has_required_columns_and_order(self):
        analyzer = DataAnalyzer()
        df = pd.DataFrame(
            {
                "城市": ["A市", "B市", "C市", "D市"],
                "城市总量": [1000000, 650000, 220000, 450000],
                "目标品牌份额(%)": [5.0, 18.0, 2.0, 35.0],
                "目标品牌量": [50000, 117000, 4400, 157500],
            }
        )

        ranked = analyzer.build_opportunity_priority(
            df=df,
            entity_col="城市",
            total_col="城市总量",
            share_col="目标品牌份额(%)",
            target_volume_col="目标品牌量",
            top_n=10,
        )

        expected_cols = {
            "城市",
            "城市总量",
            "目标品牌份额(%)",
            "影响分",
            "可行性分",
            "投入效率分",
            "综合优先级分",
            "优先级",
            "优先级理由",
        }
        self.assertTrue(expected_cols.issubset(set(ranked.columns)))

        scores = ranked["综合优先级分"].tolist()
        self.assertEqual(scores, sorted(scores, reverse=True))
        self.assertTrue((ranked["综合优先级分"] >= 0).all())
        self.assertTrue((ranked["综合优先级分"] <= 100).all())
        self.assertNotEqual(ranked.iloc[0]["城市"], "D市")


if __name__ == "__main__":
    unittest.main()
