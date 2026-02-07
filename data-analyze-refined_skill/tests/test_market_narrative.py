import os
import sys
import unittest


SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from main import DataAnalysisPipeline  # noqa: E402


class TestMarketNarrative(unittest.TestCase):
    def test_market_narrative_has_business_sections(self):
        pipeline = DataAnalysisPipeline()

        concentration = {
            "Top1占比": 28.2,
            "Top3占比": 66.5,
            "Top5占比": 89.8,
            "覆盖80所需实体数": 4,
            "覆盖90所需实体数": 6,
            "高值离群数": 2,
            "低值离群数": 0,
        }

        text = pipeline.build_market_narrative(
            dim_label="申报企业",
            metric_label="数量",
            concentration=concentration,
            head_names=["苏州碧迪医疗器械有限公司", "苏州林华医疗器械股份有限公司"],
            tail_names=["江西丰临医用器械有限公司"],
            has_time=False,
            has_price=False,
            has_structure=False,
        )

        required_keys = {
            "现状判断",
            "机会判断",
            "策略建议",
            "为何现在做",
            "预期收益",
            "主要风险",
            "风险对策",
        }
        self.assertEqual(required_keys, set(text.keys()))

        for key in required_keys:
            self.assertTrue(text[key])
            self.assertNotIn("中位", text[key])
            self.assertNotIn("标准差", text[key])


if __name__ == "__main__":
    unittest.main()
