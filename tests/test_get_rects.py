import unittest
from unittest.mock import patch

from tile_img_on_xl import get_rects, Size, Rect


class TestBasicFunction(unittest.TestCase):
    """get_rects 関数の基本機能のテスト
    """
    @patch('tile_img_on_xl.conf')
    def test_exact_multiple(self, mock_conf):
        """元サイズが max_w, max_h の整数倍
        """
        mock_conf.max_w = 2
        mock_conf.max_h = 3
        rects, counts = get_rects(Size(6, 6))
        expected = [
            Rect(0, 0, 2, 3),
            Rect(2, 0, 4, 3),
            Rect(4, 0, 6, 3),
            Rect(0, 3, 2, 6),
            Rect(2, 3, 4, 6),
            Rect(4, 3, 6, 6),
        ]
        self.assertEqual(rects, expected)
        self.assertEqual(counts, Size(3, 2))

    @patch('tile_img_on_xl.conf')
    def test_not_multiple(self, mock_conf):
        """元サイズが max_w, max_h で割り切れない
        """
        mock_conf.max_w = 6
        mock_conf.max_h = 4
        rects, counts = get_rects(Size(14, 10))
        expected = [
            Rect(0, 0, 5, 4),
            Rect(5, 0, 10, 4),
            Rect(10, 0, 14, 4),
            Rect(0, 4, 5, 8),
            Rect(5, 4, 10, 8),
            Rect(10, 4, 14, 8),
            Rect(0, 8, 5, 10),
            Rect(5, 8, 10, 10),
            Rect(10, 8, 14, 10),
        ]
        self.assertEqual(rects, expected)
        self.assertEqual(counts, Size(3, 3))

    @patch('tile_img_on_xl.conf')
    def test_no_slice(self, mock_conf):
        """max_w, max_h が 0
        """
        mock_conf.max_w = 0
        mock_conf.max_h = 0
        rects, counts = get_rects(Size(23, 31))
        expected = [
            Rect(0, 0, 23, 31),
        ]
        self.assertEqual(rects, expected)
        self.assertEqual(counts, Size(1, 1))
