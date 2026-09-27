import json
import unittest
from fiscal_tables import select_latest, select_first_estimate, extract_metric_history, metric_points


class Page:
    def __init__(self, tables):
        self.tables = tables

    def locator(self, selector):
        return self

    def evaluate_all(self, expression):
        return self.tables


class FiscalTablesTests(unittest.TestCase):
    def test_latest_prefers_ltm_in_either_column_order(self):
        points = [("Dec '17", '99'), ("Dec '24", '247'), ("Dec '25", '236'), ('LTM', '239')]
        self.assertEqual(select_latest(points), '239')
        self.assertEqual(select_latest(points[::-1]), '239')

    def test_latest_ignores_forecasts_and_missing_values(self):
        self.assertEqual(select_latest([('LTM', '—'), ("Dec '25", '236'), ("Dec '26 (E)", '300')]), '236')
        self.assertEqual(select_latest([("Dec '25", '0')]), '0')

    def test_forecast_is_first_estimate_not_historical_mean(self):
        self.assertEqual(select_first_estimate([("Dec '25 (A)", '3,830'), ("Dec '27 (E)", '5,000'), ("Dec '26 (E)", '€4,143.64')]), '€4,143.64')
        self.assertEqual(select_first_estimate([("Dec '25 (A)", '3,830')]), 'N/A')
        self.assertEqual(select_first_estimate([("Dec '26 (E)", '—'), ("Dec '27 (E)", '5,000')]), 'N/A')

    def test_history_keeps_missing_column_alignment_and_own_headers(self):
        page = Page([
            {'headers': ['Other', "Dec '99"], 'rows': [['Other', '10']]},
            {'headers': ['Cash Flow', "Dec '23", "Dec '24", "Dec '25", 'LTM'],
             'rows': [['Free Cash Flow', '1', '', '3', '4'], ['Free Cash Flow / Share', '8', '9', '10', '11']]},
        ])
        self.assertEqual(json.loads(extract_metric_history(page, 'Free Cash Flow')),
                         [{'label': "Dec '23", 'value': 1}, {'label': "Dec '25", 'value': 3}])
        self.assertEqual(metric_points(page, 'Free Cash Flow', 'Other'), [])


if __name__ == '__main__':
    unittest.main()
