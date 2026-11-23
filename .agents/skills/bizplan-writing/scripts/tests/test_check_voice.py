"""check_voice.py 단위 테스트.

실행: python3 -I -m unittest discover -s .agents/skills/bizplan-writing/scripts/tests -p 'test_check_voice.py' -v

-I로 실행하면 scripts 폴더가 sys.path에 들어오지 않으므로 아래에서 직접 넣는다.
fixtures/voice_*.txt는 한 줄이 한 문단이다(빈 줄은 문단으로 세지 않는다).
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

import check_voice  # noqa: E402

FIXTURES = os.path.join(HERE, 'fixtures')


def fixture(name):
    return os.path.join(FIXTURES, name)


def paras_of(name):
    return check_voice.paragraphs(fixture(name))


def hits_of(name):
    return check_voice.word_hits(paras_of(name))


def run_main(name, show_all=False):
    argv = [fixture(name)] + (['--all'] if show_all else [])
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
        code = check_voice.main(argv)
    return code, buf.getvalue()


class StrengthTableTest(unittest.TestCase):
    def test_strong_categories(self):
        for cat in ('상투어', '번역투', '기대 맺음', '작업 메모', '같은 부정 반복', '경구형 제목'):
            self.assertEqual(check_voice.meta(cat)[0], '강', cat)

    def test_weak_categories(self):
        for cat in ('막연한 수식', 'AI 문장부호', '나열 공식', '대조 구문', '같은 끝말 연속',
                    '같은 첫말 연속', '길이가 고른 문단 묶음'):
            self.assertEqual(check_voice.meta(cat)[0], '약', cat)

    def test_output_order_covers_all_categories(self):
        self.assertEqual(len(check_voice.ORDER), 13)
        self.assertEqual(set(check_voice.ORDER),
                         set(c for c, _, _, _ in check_voice.WORD_RULES) | set(check_voice.STRUCT_RULES))


class WorkMemoTest(unittest.TestCase):
    def test_drafting_marks_are_caught(self):
        hit = [i for i, _, _ in hits_of('voice_work_memo.txt')['작업 메모']]
        self.assertEqual(hit, [1, 2])

    def test_unconfirmed_fact_is_not_a_memo(self):
        hit = [i for i, _, _ in hits_of('voice_work_memo.txt')['작업 메모']]
        self.assertNotIn(3, hit)  # 「돈을 낼 의사 미확인」
        self.assertNotIn(4, hit)  # 「— 미확인」


class ProverbTitleTest(unittest.TestCase):
    def test_short_aphorism_title_is_caught(self):
        titles = check_voice.proverb_titles(paras_of('voice_proverb_title.txt'))
        self.assertEqual([i for i, _ in titles], [1, 4])

    def test_plain_bullet_titles_are_not_caught(self):
        titles = check_voice.proverb_titles(paras_of('voice_proverb_title.txt'))
        texts = [t for _, t in titles]
        self.assertNotIn('○ 시장: 계약 규모와 제품이 지금 다루는 공고', texts)
        self.assertNotIn('◦ 네 업종 공고 검색', texts)


class NegativeRepeatTest(unittest.TestCase):
    def test_same_negative_phrase_is_grouped(self):
        groups = check_voice.repeated_negations(paras_of('voice_negative_repeat.txt'))
        self.assertEqual(len(groups), 1, groups)
        key, nums, phrase = groups[0]
        self.assertEqual(key, ('BidBox', '매출'))
        self.assertEqual(nums, [1, 2])
        self.assertIn('BidBox 매출 아님', phrase)

    def test_required_notice_is_not_counted(self):
        groups = check_voice.repeated_negations(paras_of('voice_negative_repeat.txt'))
        for _, nums, phrase in groups:
            self.assertNotIn('낙찰 보증 없음', phrase)
            self.assertFalse(set(nums) & {3, 4, 5}, nums)


class ContrastTest(unittest.TestCase):
    def test_negative_contrast_is_caught(self):
        hit = [i for i, _, _ in hits_of('voice_contrast.txt')['대조 구문']]
        self.assertEqual(hit, [1])

    def test_plain_colon_and_unconfirmed_are_not_caught(self):
        hit = [i for i, _, _ in hits_of('voice_contrast.txt')['대조 구문']]
        self.assertNotIn(2, hit)  # 「○ 시장: …」는 항목 이름 뒤 콜론
        self.assertNotIn(3, hit)  # 「미확인」은 사실을 밝히는 말


class StockWordTest(unittest.TestCase):
    def test_new_stock_words_are_caught(self):
        hit = [i for i, _, _ in hits_of('voice_stock_words.txt')['상투어']]
        self.assertEqual(hit, [1, 2])  # 손쉽게, 본질

    def test_fact_words_are_not_caught(self):
        hit = [i for i, _, _ in hits_of('voice_stock_words.txt')['상투어']]
        for i in (3, 4, 5):
            self.assertNotIn(i, hit)  # 핵심·한눈에·누구나


class SameOpeningTest(unittest.TestCase):
    def test_three_paragraphs_with_same_first_word_is_caught(self):
        runs = check_voice.same_opening_runs(paras_of('voice_same_opening_run.txt'))
        self.assertEqual(runs, [(1, 3, 'BidBox는')])

    def test_two_paragraphs_are_not_caught(self):
        self.assertEqual(check_voice.same_opening_runs(paras_of('voice_same_opening_short.txt')), [])


class CommaListTest(unittest.TestCase):
    def test_comma_items_skip_numbers_and_parentheses(self):
        dist = check_voice.comma_item_counts(paras_of('voice_comma_list.txt'))
        self.assertEqual(dist, {2: 1, 3: 1, '4+': 0})


class OutputTest(unittest.TestCase):
    def test_exit_code_is_zero_and_headers_are_printed(self):
        code, out = run_main('voice_work_memo.txt')
        self.assertEqual(code, 0)
        self.assertIn('[강] 작업 메모 2곳', out)
        self.assertIn('[약] AI 문장부호', out)
        self.assertIn('[참고] 쉼표 나열 항목 수', out)
        self.assertIn('[요약] 강 합계 2곳', out)
        self.assertIn('[약 신호 겹침]', out)

    def test_all_flag_shows_every_hit(self):
        code, out = run_main('voice_work_memo.txt', show_all=True)
        self.assertEqual(code, 0)
        self.assertIn('제출 전 확인', out)
        self.assertIn('[확인 필요', out)


if __name__ == '__main__':
    unittest.main()
