import unittest

from pyxl.functions import *

class TestExcelFunctions(unittest.TestCase):

    def test_ACCRINT(self):
        self.assertAlmostEqual(ACCRINT(39508, 39691, 39569, 0.1, 1000, 2, 0), 16.666667, 6)
        self.assertAlmostEqual(ACCRINT(DATE(2008, 3, 5), 39691, 39569, 0.1, 1000, 2, 0, False), 15.555556, 6)
        self.assertAlmostEqual(ACCRINT(DATE(2008, 4, 5), 39691, 39569, 0.1, 1000, 2, 0, True), 7.2222222, 6)

    def test_ACCRINTM(self):
        self.assertAlmostEqual(ACCRINTM(39539, 39614, 0.1, 1000, 3), 20.54794521)

    def test_ACOS(self):
        self.assertAlmostEqual(ACOS(-0.5), 2.094395102)

    def test_ACOSH(self):
        self.assertAlmostEqual(ACOSH(1), 0)
        self.assertAlmostEqual(ACOSH(10), 2.9932228)

    def test_ACOT(self):
        self.assertAlmostEqual(ACOT(2), 0.4636, 4)

    def test_ACOTH(self):
        self.assertAlmostEqual(ACOTH(6), 0.168, 3)

    # def test_ASC(self):
    #     self.assertEqual(ASC('ＡＢＣａｂｃ０１２！＃＄アイウガギグ　'), 'ABCabc012!#$ｱｲｳｶﾞｷﾞｸﾞ ')

    def test_BAHTTEXT(self):
        self.assertEqual(BAHTTEXT(123), 'หนึ่งร้อยยี่สิบสามบาทถ้วน')
        self.assertEqual(BAHTTEXT(123.456), 'หนึ่งร้อยยี่สิบสามบาทสี่สิบหกสตางค์')
        self.assertEqual(BAHTTEXT(-1), 'ลบหนึ่งบาทถ้วน')
        self.assertEqual(BAHTTEXT(-1.78), 'ลบหนึ่งบาทเจ็ดสิบแปดสตางค์')
        self.assertEqual(BAHTTEXT(0), 'ศูนย์บาทถ้วน')
        self.assertEqual(BAHTTEXT(0.49), 'สี่สิบเก้าสตางค์')
        self.assertEqual(BAHTTEXT(-0.25), 'ลบยี่สิบห้าสตางค์')
        # self.assertEqual(BAHTTEXT(10000000000), 'หนึ่งหมื่นล้านบาทถ้วน')

    def test_BASE(self):
        self.assertEqual(BASE(9999, 2), '10011100001111')
        self.assertEqual(BASE(9999, 3), '111201100')
        self.assertEqual(BASE(9999, 8), '23417')
        self.assertEqual(BASE(9999, 11), '7570')
        self.assertEqual(BASE(9999, 16), '270F')
        self.assertEqual(BASE(9999, 25), 'FOO')

    def test_BIN2DEC(self):
        self.assertEqual(BIN2DEC(1100100), '100')
        self.assertEqual(BIN2DEC(1111111111), '-1')

    def test_BIN2HEX(self):
        self.assertEqual(BIN2HEX(1100100), '64')

    def test_BIN2OCT(self):
        self.assertEqual(BIN2OCT(1100100), '144')

    def test_CHAR(self):
        self.assertEqual(CHAR(65), 'A')
        self.assertEqual(CHAR(33), '!')
    
    def test_CONCAT(self):
        self.assertEqual(CONCAT("HELLO", " ","WORLD"), "HELLO WORLD")

    def test_CONCATENATE(self):
        self.assertEqual(CONCATENATE("HELLO", " ","WORLD"), "HELLO WORLD")

    def test_DATEVALUE(self):
        self.assertEqual(DATEVALUE('8/22/2011'), 40777)
        self.assertEqual(DATEVALUE('22-MAY-2011'), 40685)
        self.assertEqual(DATEVALUE('2011/02/23'), 40597)

    def test_DEC2BIN(self):
        self.assertEqual(DEC2BIN(73), '1001001')

    def test_DEC2HEX(self):
        self.assertEqual(DEC2HEX(12648430), 'C0FFEE')

    def test_DEC2OCT(self):
        self.assertEqual(DEC2OCT(73), '111')

    def test_DEGREES(self):
        self.assertAlmostEqual(DEGREES(PI()), 180)
        self.assertAlmostEqual(DEGREES(ACOS(-0.5)), 120)

    def test_FISHER(self):
        self.assertAlmostEqual(FISHER(0.75), 0.9729551)

    def test_FISHERINV(self):
        self.assertAlmostEqual(FISHERINV(0.972955), 0.75)

    def test_FLOOR(self):
        self.assertEqual(FLOOR(1.5), 1)
        self.assertEqual(FLOOR(7, 5), 5)

    def test_FLOOR_MATH(self):
        self.assertEqual(FLOOR_MATH(1.5), 1)
        self.assertEqual(FLOOR_MATH(7, 5), 5)
        self.assertEqual(FLOOR_MATH(-7.3, 3), -9)
        self.assertEqual(FLOOR_MATH(-7.3, 3, -1), -6)

    def test_GAUSS(self):
        self.assertAlmostEqual(GAUSS(1), 0.3413447461)
        self.assertAlmostEqual(GAUSS(2), 0.4772498681)
        self.assertAlmostEqual(GAUSS(3), 0.498650102)
        self.assertAlmostEqual(GAUSS(4), 0.4999683288)
        self.assertAlmostEqual(GAUSS(5), 0.4999997133)

    def test_GCD(self):
        self.assertAlmostEqual(GCD(2, 4), 2)
        self.assertAlmostEqual(GCD(2, 4, 5), 1)
        self.assertAlmostEqual(GCD(5, 2), 1)
        self.assertAlmostEqual(GCD(24, 36), 12)
        self.assertAlmostEqual(GCD(7, 1), 1)
        self.assertAlmostEqual(GCD(5, 0), 5)

    def test_HEX2BIN(self):
        self.assertEqual(HEX2BIN('A5'), '10100101')

    def test_HEX2DEC(self):
        self.assertEqual(HEX2DEC('A5'), '165')

    def test_HEX2OCT(self):
        self.assertEqual(HEX2OCT('A5'), '245')

    # def test_JIS(self):
    #     self.assertEqual(JIS('ABCabc012!#$ｱｲｳｶﾞｷﾞｸﾞ '), 'ＡＢＣａｂｃ０１２！＃＄アイウガギグ　')

    def test_OCT2BIN(self):
        self.assertEqual(OCT2BIN(65), '110101')

    def test_OCT2DEC(self):
        self.assertEqual(OCT2DEC(65), '53')

    def test_OCT2HEX(self):
        self.assertEqual(OCT2HEX(65), '35')

if __name__ == '__main__':
    unittest.main()