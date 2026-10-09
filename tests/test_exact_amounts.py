"""Exercise the production amount parser without running the PDF pipeline."""
import ast
import re
import unittest
from pathlib import Path


def load_parser():
    source = Path(__file__).resolve().parents[1] / 'src/stage3_extract.py'
    names = {
        'clean_space', 'normalize_house_pdf_text', 'money_number',
        'amount_continuation_prefix', 'parse_amount_text', 'STANDARD_AMOUNT_RANGES',
        'DETAIL_LABEL_RE', 'CAP_GAINS_HEADER_MONEY_RE', 'OPEN_WITH_MONEY_RE',
        'OPEN_TRAILING_RE', 'MONEY_RE', 'HOUSE_BROKEN_UNICODE_START',
        'HOUSE_BROKEN_UNICODE_END', 'HOUSE_BROKEN_UNICODE_OFFSET',
    }
    nodes = []
    for node in ast.parse(source.read_text()).body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            nodes.append(node)
        elif isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in names for t in node.targets
        ):
            nodes.append(node)
    namespace = {'re': re}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), 'exec'), namespace)
    return namespace['parse_amount_text']


class AmountTests(unittest.TestCase):
    def test_exact_with_explicit_cents(self):
        parse = load_parser()
        for text, value in [('$2,000.00', 2000), ('$15,001.00', 15001), ('$318.74', 318.74)]:
            result = parse(text)
            self.assertEqual(result['amount_exact'], value)
            self.assertEqual(result['amount_status'], 'nonstandard_exact')
            self.assertIsNone(result['amount_min'])
            self.assertIsNone(result['amount_max'])

    def test_ranges_and_ambiguous_bounds(self):
        parse = load_parser()
        for text in ['$1,001 - $15,000', '$1,001.00 - $15,000.00']:
            result = parse(text)
            self.assertEqual(result['amount_status'], 'valid_range')
            self.assertEqual((result['amount_min'], result['amount_max']), (1001, 15000))
        for text in ['$15,001', '$15,001 -', '$15,001.00 -']:
            self.assertEqual(parse(text)['amount_status'], 'missing_range_bound')
        self.assertEqual(parse('Over $1,000,000')['amount_status'], 'valid_open_ended')
        self.assertEqual(parse('$1,001 -', '$15,000 Description: $2,000.00')['amount_status'], 'valid_range')


if __name__ == '__main__':
    unittest.main()
