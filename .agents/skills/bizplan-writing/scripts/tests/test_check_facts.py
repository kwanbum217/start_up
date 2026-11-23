"""check_facts.py 단위 테스트.

실행: python3 -I -m unittest discover -s .agents/skills/bizplan-writing/scripts/tests -p 'test_check_facts.py' -v
"""
import contextlib
import io
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import check_facts  # noqa: E402

FIXTURES = os.path.join(HERE, 'fixtures')


def fixture(name):
    return os.path.join(FIXTURES, name)


def run_main(old, new, show_all=False):
    argv = [old, new] + (['--all'] if show_all else [])
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
        code = check_facts.main(argv)
    return code, buf.getvalue()


def tokens_of(name):
    out = []
    for p in check_facts.paragraphs(fixture(name)):
        out.extend(t for t, _, _ in check_facts.extract_facts(p))
    return out


class SameAndChangedTest(unittest.TestCase):
    def test_same_text_has_no_diff(self):
        code, out = run_main(fixture('facts_base.txt'), fixture('facts_same.txt'))
        self.assertEqual(code, 0, out)
        self.assertIn('차이 0종', out)

    def test_changed_number_is_gone_and_added(self):
        code, out = run_main(fixture('facts_base.txt'), fixture('facts_number_changed.txt'))
        self.assertEqual(code, 1)
        gone_part, added_part = out.split('[새로 생김]')
        self.assertIn('57.3조원  1→0', gone_part)
        self.assertIn('55.1조원  0→1', added_part)

    def test_rewritten_sentences_with_same_numbers_have_no_diff(self):
        code, out = run_main(fixture('facts_base.txt'), fixture('facts_rewritten.txt'))
        self.assertEqual(code, 0, out)
        self.assertIn('차이 0종', out)

    def test_all_flag_shows_every_token(self):
        code, out = run_main(fixture('facts_base.txt'), fixture('facts_number_changed.txt'), show_all=True)
        self.assertEqual(code, 1)
        self.assertIn('57.3조원', out)
        self.assertNotIn('--all로 전부 보기', out)


class TokenShapeTest(unittest.TestCase):
    def test_date_is_one_token_not_split_into_numbers(self):
        tokens = tokens_of('facts_dates.txt')
        self.assertEqual(tokens.count('2026-06-26'), 2)
        self.assertEqual(tokens.count('2026년 6월'), 1)
        for bad in ('2026', '06', '26', '2026년'):
            self.assertNotIn(bad, tokens)

    def test_comma_percent_docnumber_fraction_tokens(self):
        tokens = tokens_of('facts_tokens.txt')
        self.assertIn('31,305건', tokens)
        self.assertIn('144/144', tokens)
        self.assertIn('제2026-260호', tokens)
        for bad in ('31', '305건', '2026', '05', '26'):
            self.assertNotIn(bad, tokens)


class CountChangeTest(unittest.TestCase):
    def test_repeat_count_decrease_is_reported(self):
        code, out = run_main(fixture('facts_repeat3.txt'), fixture('facts_repeat2.txt'))
        self.assertEqual(code, 1)
        self.assertIn('8사  3→2', out)


class ExitCodeTest(unittest.TestCase):
    def test_missing_file_exits_2(self):
        code, _ = run_main(fixture('facts_base.txt'), fixture('facts_there_is_no_such_file.txt'))
        self.assertEqual(code, 2)

    def test_wrong_argument_count_exits_2(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
            code = check_facts.main([fixture('facts_base.txt')])
        self.assertEqual(code, 2)


if __name__ == '__main__':
    unittest.main()
