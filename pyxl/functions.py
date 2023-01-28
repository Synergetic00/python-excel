import calendar
import datetime
import math
import re
import string

class _ExcelError(Exception):
    ''''''
    pass

MONTH_VALUES = {
    'JAN': 1,
    'FEB': 2,
    'MAR': 3,
    'APR': 4,
    'MAY': 5,
    'JUN': 6,
    'JUL': 7,
    'AUG': 8,
    'SEP': 9,
    'OCT': 10, 
    'NOV': 11,
    'DEC': 12
}

PYTHON_EXCEL_ORDINAL_DIFF = 693594

def _clamp(num, min_value, max_value):
   return int(max(min(num, int(max_value)), int(min_value)))

def ABS(number):
    '''Returns the bsolute value of a number'''
    try: cnumber = int(str(number))
    except: raise TypeError('Only integers are allowed')
    return abs(cnumber)

DAY_COUNT_BASIS = {
    'UsPsa30_360'  : 0,
    'ActualActual' : 1,
    'Actual360'    : 2,
    'Actual365'    : 3,
    'Europe30_360' : 4
}

METHOD_360_US = {
    'ModifyStartDate': 0,
    'ModifyBothDates': 1
}

NUM_DENUM_POSITION = {
    'Denumerator': 0,
    'Numerator': 1
}
ACCR_INT_CALC_METHOD = {
    'FromFirstToSettlement': 0,
    'FromIssueToSettlement': 1
}

def _is_leap_year(year):
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

def _extract_date_from_value(date_value):
    date = datetime.date.fromordinal(date_value + PYTHON_EXCEL_ORDINAL_DIFF)
    return date.day, date.month, date.year

def _is_feb_29_between_consecutive_years(issue, settlement):
    _, i_month, i_year = _extract_date_from_value(issue)
    _, s_month, s_year = _extract_date_from_value(settlement)
    if (i_year == s_year) and _is_leap_year(i_year):
        return (i_month <= 2) and (s_month > 2)
    if (i_year == s_year):
        return False
    if (i_year == (s_year + 1)):
        if _is_leap_year(i_year):
            return i_month <= 2
        else:
            if _is_leap_year(s_year):
                return s_month > 2
            else:
                return False
    else:
        return False

def _consider_as_bisestile(issue, settlement):
    _, _, i_year = _extract_date_from_value(issue)
    s_day, s_month, s_year = _extract_date_from_value(settlement)
    return ((i_year == s_year) and _is_leap_year(i_year)) or (s_month == 2 and s_day == 29) or _is_feb_29_between_consecutive_years(issue, settlement)

def _less_or_equal_to_a_year_apart(issue, settlement):
    i_day, i_month, i_year = _extract_date_from_value(issue)
    s_day, s_month, s_year = _extract_date_from_value(settlement)
    return (i_year == s_year) or ((s_year == (i_year + 1)) and ((i_month > s_month) or ((i_month == s_month) and i_day >= s_day)))

def _actual_coup_days(settl, mat, freq):
    pcd = _find_previous_coupon_date(settl, mat, freq, DAY_COUNT_BASIS['ActualActual'])
    ncd = _find_next_coupon_date(settl, mat, freq, DAY_COUNT_BASIS['ActualActual'])
    return abs((ncd - pcd).days)

def _find_previous_coupon_date(settl, mat, freq, basis) -> datetime.date:
    return _find_coupon_dates(settl, mat, freq, basis)[0]

def _find_next_coupon_date(settl, mat, freq, basis) -> datetime.date:
    return _find_coupon_dates(settl, mat, freq, basis)[1]

def _find_coupon_dates(settl, mat, freq, basis):
    end_month = _is_last_day_of_month(mat)
    num_months = -_freq_2_months(freq)
    return _find_pcd_ncd(mat, settl, num_months, basis, end_month)

def _last_day_of_month(month, year):
    return calendar.monthrange(year, month)[-1]

def _is_last_day_of_month(date_value):
    day, month, year = _extract_date_from_value(date_value)
    return _last_day_of_month(month, year) == day

def _freq_2_months(freq):
    return 12 / freq

def _find_pcd_ncd(start_date: int, end_date: int, num_months, basis, return_last_month):
    pcd, ncd, _ = _dates_aggregate_1(start_date, end_date, num_months, basis, 0, 0, return_last_month)
    return pcd, ncd

def _change_month(date, num_months, basis, return_last_day):
    orig_date = date
    o_day, o_month, o_year = _extract_date_from_value(orig_date)
    o_month += num_months
    if o_month > 12:
        o_month %= 12
        o_year += 1
    
    if o_month == 2:
        feb_days = 29 if _is_leap_year(o_year) else 28
        o_day = _clamp(o_day, 1, feb_days)
    new_date = DATE(o_year, int(o_month), int(o_day))
    _, n_month, n_year = _extract_date_from_value(new_date)
    last_day = _last_day_of_month(n_month, n_year)
    if return_last_day:
        return DATE(n_year, n_month, last_day)
    else:
        return new_date

def _dates_aggregate_1(start_date: int, end_date: int, num_months, basis, f, acc, return_last_month):
    front_date = start_date
    trailing_date = end_date
    s1 = (front_date > end_date) or (front_date == end_date)
    s2 = (end_date > front_date) or (end_date == front_date)
    stop = s1 if num_months > 0 else s2
    while stop == False:
        trailing_date = front_date
        front_date = _change_month(front_date, num_months, basis, return_last_month)
        func = f(front_date, trailing_date)
        acc = acc + func
        s1 = (front_date > end_date) or (front_date == end_date)
        s2 = (end_date > front_date) or (end_date == front_date)
        stop = s1 if num_months > 0 else s2
    return front_date, trailing_date, acc

def _coup_days(basis, settl, mat, freq):
    if basis == DAY_COUNT_BASIS['ActualActual']:
        return _actual_coup_days(settl, mat, freq)
    elif basis == DAY_COUNT_BASIS['Actual365']:
        return 365 / freq
    return 360 / freq

def _coup_pcd(basis, settl, mat, freq):
    return _find_previous_coupon_date(settl, mat, freq, basis)

def _coup_ncd(basis, settl, mat, freq):
    return _find_previous_coupon_date(settl, mat, freq, basis)

def _number_of_coupons(settl, mat, freq, basis):
    pcdate = _find_previous_coupon_date(settl, mat, freq, basis)
    _, pcm, pcy = _extract_date_from_value(pcdate)
    _, mm, my = _extract_date_from_value(mat)
    months = (my - pcy) * 12 + (mm - pcm)
    return months * freq / 12

def _coup_num(basis, settl, mat, freq):
    return _number_of_coupons(basis, settl, mat, freq)

def _date_diff_360(start_date, end_date):
    s_day, s_month, s_year = _extract_date_from_value(start_date)
    e_day, e_month, e_year = _extract_date_from_value(end_date)
    return ((e_year - s_year) * 360) + ((e_month - s_month) * 30) + (e_day - s_day)

def _date_diff_365(start_date, end_date):
    s_day, s_month, s_year = _extract_date_from_value(start_date)
    e_day, e_month, e_year = _extract_date_from_value(end_date)
    if (s_day > 28) and (s_month == 2):
        s_day = 28
    if (e_day > 28) and (e_month == 2):
        e_day = 28
    startd = datetime.date(s_year, s_month, s_day)
    endd = datetime.date(e_year, e_month, e_day)
    return (e_year - s_year) * 365 + abs((endd - startd).days)

def _is_last_day_of_february(date_value):
    day, month, year = _extract_date_from_value(date_value)
    if (month != 2): return False
    if _is_leap_year(year): return day == 29
    else: return day == 28

def _date_diff_360_us(start_date, end_date, method_360):
    s_day, s_month, s_year = _extract_date_from_value(start_date)
    e_day, e_month, e_year = _extract_date_from_value(end_date)
    if _is_last_day_of_february(end_date) and (_is_last_day_of_february(start_date) or method_360 == METHOD_360_US['ModifyStartDate']):
        e_day = 30
    if e_day == 31 and (s_day >= 30 or method_360 == METHOD_360_US['ModifyBothDates']):
        e_day = 30
    if s_day == 31:
        s_day = 30
    if (_is_last_day_of_february(start_date)):
        s_day = 30
    d1 = DATE(s_year, s_month, s_day)
    d2 = DATE(e_year, e_month, e_day)
    return _date_diff_360(d1, d2)

def _date_diff_360_eu(start_date, end_date):
    s_day, s_month, s_year = _extract_date_from_value(start_date)
    e_day, e_month, e_year = _extract_date_from_value(end_date)
    if sd1 == 31:
        sd1 = 30
    if e_day == 31:
        e_day = 30
    d1 = datetime.date(s_year, s_month, s_day)
    d2 = datetime.date(e_year, e_month, e_day)
    return _date_diff_360(d1, d2)

def _coup_days_bs(basis, settl: datetime.date, mat, freq):
    cpcd: datetime.date = _coup_pcd(settl, mat, freq)
    if basis == DAY_COUNT_BASIS['UsPsa30_360']:
        return _date_diff_360_us(cpcd, settl, METHOD_360_US['ModifyStartDate'])
    elif basis == DAY_COUNT_BASIS['Europe30_360']:
        return _date_diff_360_eu(cpcd, settl)
    return abs((settl - cpcd).days)

def _coup_days_nc(basis, settl: datetime.date, mat, freq):
    cncd: datetime.date = _coup_ncd(settl, mat, freq)
    if basis == DAY_COUNT_BASIS['UsPsa30_360']:
        pdc = _find_previous_coupon_date(settl, mat, freq, DAY_COUNT_BASIS['UsPsa30_360'])
        ndc = _find_next_coupon_date(settl, mat, freq, DAY_COUNT_BASIS['UsPsa30_360'])
        tot_days_in_coup = _date_diff_360_us(pdc, ndc, METHOD_360_US['ModifyBothDates'])
        days_to_settl = _date_diff_360_us(pdc, settl, METHOD_360_US['ModifyStartDate'])
        return tot_days_in_coup - days_to_settl
    elif basis == DAY_COUNT_BASIS['Europe30_360']:
        return _date_diff_360_eu(settl, cncd)
    return abs((settl - cncd).days)

def _days_between(basis, issue, settl, position):
    if basis == DAY_COUNT_BASIS['UsPsa30_360']:
        return _date_diff_360_us(issue, settl, METHOD_360_US['ModifyStartDate'])
    elif basis == DAY_COUNT_BASIS['ActualActual']:
        return abs(settl - issue)
    elif basis == DAY_COUNT_BASIS['Actual360']:
        if position == NUM_DENUM_POSITION['Numerator']:
            return abs(settl - issue)
        else:
            return _date_diff_360_us(issue, settl, METHOD_360_US['ModifyStartDate'])
    elif basis == DAY_COUNT_BASIS['Actual365']:
        if position == NUM_DENUM_POSITION['Numerator']:
            return abs(settl - issue)
        else:
            return _date_diff_365(issue, settl)
    elif basis == DAY_COUNT_BASIS['Europe30_360']:
        return _date_diff_360_eu(issue, settl)

def _days_in_year(basis, issue, settl):
    if basis == DAY_COUNT_BASIS['ActualActual']:
        if _less_or_equal_to_a_year_apart(issue, settl):
            _, _, i_year = _extract_date_from_value(issue)
            _, _, s_year = _extract_date_from_value(settl)
            total_years = (s_year - i_year) + 1
            d1 = datetime.date(i_year, 1, 1)
            d2 = datetime.date(s_year + 1, 1, 1)
            total_days = (d2 - d1).days
            return total_days / total_years
        else:
            return 366 if _consider_as_bisestile(issue, settl) else 365
    elif basis == DAY_COUNT_BASIS['Actual365']:
        return 365
    return 360

def ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis=0, calc_method=True):
    '''Returns the accrued interest for a security that pays periodic interest'''
    if rate <= 0 or par <= 0: raise _ExcelError('Either rate or par are lower than or equal to zero')
    if frequency not in [1, 2, 4]: raise _ExcelError('Incorrect frequency value, must be: 1, 2 or 4')
    if basis not in [0, 1, 2, 3, 4]: raise _ExcelError('Incorrect basis value, must be: 0, 1, 2, 3 or 4')
    if issue >= settlement: raise _ExcelError('Issue is greater than or equal to settlement')
    num_months = _freq_2_months(frequency)
    num_months_neg = -num_months
    end_month_bond = _is_last_day_of_month(first_interest)
    if calc_method:
        calc_method = ACCR_INT_CALC_METHOD['FromIssueToSettlement']
    else:
        calc_method = ACCR_INT_CALC_METHOD['FromFirstToSettlement']
    if settlement > first_interest and calc_method == ACCR_INT_CALC_METHOD['FromIssueToSettlement']:
        pcd, _ = _find_pcd_ncd(first_interest, settlement, num_months, basis, end_month_bond)
    else:
        pcd = _change_month(first_interest, num_months_neg, basis, end_month_bond)
    first_date = issue if issue > pcd else pcd
    days_between = _days_between(basis, first_date, settlement, NUM_DENUM_POSITION['Numerator'])
    days_coup = _coup_days(basis, pcd, first_interest, frequency)
    def _agg_function(apcd, ancd):
        a_first_date = issue if issue > apcd else apcd
        if (basis == DAY_COUNT_BASIS['UsPsa30_360']):
            psa_method = METHOD_360_US['ModifyStartDate'] if issue > apcd else METHOD_360_US['ModifyBothDates']
            a_days = _date_diff_360_us(a_first_date, ancd, psa_method)
        else:
            a_days = _days_between(a_first_date, ancd, NUM_DENUM_POSITION['Numerator'])
        if (basis == DAY_COUNT_BASIS['UsPsa30_360']):
            a_coup_days = _date_diff_360_us(apcd, ancd, METHOD_360_US['ModifyBothDates'])
        else:
            if (basis == DAY_COUNT_BASIS['Actual365']):
                a_coup_days = 365 / frequency
            else:
                a_coup_days = _days_between(apcd, ancd, NUM_DENUM_POSITION['Denumerator'])
        if apcd > issue or apcd == issue:
            result = calc_method
        else:
            result = (a_days / a_coup_days)
        return result
    _, _, a = _dates_aggregate_1(pcd, issue, num_months_neg, basis, _agg_function, (days_between/days_coup), end_month_bond)
    return par * rate / frequency * a

def ACCRINTM(issue, settlement, rate, par, basis=0):
    '''Returns the accrued interest for a security that pays interest at maturity'''
    days_between = _days_between(basis, issue, settlement, NUM_DENUM_POSITION['Numerator'])
    days_in_year = _days_in_year(basis, issue, settlement)
    return par * rate * (days_between / days_in_year)

def ACOS(number):
    '''Returns the arccosine of a number'''
    try: cnumber = float(number)
    except ValueError: raise _ExcelError('Only floats are allowed')
    return math.acos(cnumber)

def ACOSH(number):
    '''Returns the inverse hyperbolic cosine of a number'''
    try: cnumber = float(number)
    except ValueError: raise _ExcelError('Only floats are allowed')
    if cnumber < 1: raise _ExcelError(f'Value should be greater than or equal to 1.')
    return math.acosh(cnumber)

def ACOT(number):
    '''Returns the arccotangent of a number'''
    try: cnumber = float(number)
    except ValueError: raise _ExcelError('Only floats are allowed')
    return PI() / 2 - math.atan(cnumber)

def ACOTH(number):
    '''Returns the hyperbolic arccotangent of a number'''
    try: cnumber = float(number)
    except ValueError: raise _ExcelError('Only floats are allowed')
    if cnumber < 1: raise _ExcelError(f'Value should be greater than or equal to 1.')
    return math.atanh(1 / cnumber)

def AGGREGATE():
    '''Returns an aggregate in a list or database'''
    pass

def ADDRESS():
    '''Returns a reference as text to a single cell in a worksheet'''
    pass

def AMORDEGRC():
    '''Returns the depreciation for each accounting period by using a depreciation coefficient'''
    pass

def AMORLINC():
    '''Returns the depreciation for each accounting period'''
    pass

def AND(value, *values):
    '''Returns TRUE if all of its arguments are TRUE'''
    pass

ARABIC_DICT = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}

def ARABIC(text):
    '''Converts a Roman number to Arabic, as a number'''
    parsed_text = str(text).strip().upper()
    if any(chr not in ARABIC_DICT.keys() for chr in parsed_text):
        raise _ExcelError('Input is not valid Roman numerals')
    if len(parsed_text) == 0: return 0
    result = 0
    for i,c in enumerate(parsed_text):
        if (i+1) == len(parsed_text) or ARABIC_DICT[c] >= ARABIC_DICT[parsed_text[i+1]]:
            result += ARABIC_DICT[c]
        else:
            result -= ARABIC_DICT[c]
    return result

def AREAS():
    '''Returns the number of areas in a reference'''
    pass

def ARRAYTOTEXT():
    '''Returns an array of text values from any specified range'''
    pass

FULL_HALF_WIDTH_MAP = {
    'ガ': 'ｶﾞ',
    'ギ': 'ｷﾞ',
    'グ': 'ｸﾞ',
    'ゲ': 'ｹﾞ',
    'ゴ': 'ｺﾞ',
    'ザ': 'ｻﾞ',
    'ジ': 'ｼﾞ',
    'ズ': 'ｽﾞ',
    'ゼ': 'ｾﾞ',
    'ゾ': 'ｿﾞ',
    'ダ': 'ﾀﾞ',
    'ヂ': 'ﾁﾞ',
    'ヅ': 'ﾂﾞ',
    'デ': 'ﾃﾞ',
    'ド': 'ﾄﾞ',
    'バ': 'ﾊﾞ',
    'ビ': 'ﾋﾞ',
    'ブ': 'ﾌﾞ',
    'ベ': 'ﾍﾞ',
    'ボ': 'ﾎﾞ',
    'パ': 'ﾊﾟ',
    'ピ': 'ﾋﾟ',
    'プ': 'ﾌﾟ',
    'ペ': 'ﾍﾟ',
    'ポ': 'ﾎﾟ',
    'ヴ': 'ｳﾞ',
    'ヷ': 'ﾜﾞ',
    'ヺ': 'ｦﾞ',
    'ア': 'ｱ',
    'イ': 'ｲ',
    'ウ': 'ｳ',
    'エ': 'ｴ',
    'オ': 'ｵ',
    'カ': 'ｶ',
    'キ': 'ｷ',
    'ク': 'ｸ',
    'ケ': 'ｹ',
    'コ': 'ｺ',
    'サ': 'ｻ',
    'シ': 'ｼ',
    'ス': 'ｽ',
    'セ': 'ｾ',
    'ソ': 'ｿ',
    'タ': 'ﾀ',
    'チ': 'ﾁ',
    'ツ': 'ﾂ',
    'テ': 'ﾃ',
    'ト': 'ﾄ',
    'ナ': 'ﾅ',
    'ニ': 'ﾆ',
    'ヌ': 'ﾇ',
    'ネ': 'ﾈ',
    'ノ': 'ﾉ',
    'ハ': 'ﾊ',
    'ヒ': 'ﾋ',
    'フ': 'ﾌ',
    'ヘ': 'ﾍ',
    'ホ': 'ﾎ',
    'マ': 'ﾏ',
    'ミ': 'ﾐ',
    'ム': 'ﾑ',
    'メ': 'ﾒ',
    'モ': 'ﾓ',
    'ヤ': 'ﾔ',
    'ユ': 'ﾕ',
    'ヨ': 'ﾖ',
    'ラ': 'ﾗ',
    'リ': 'ﾘ',
    'ル': 'ﾙ',
    'レ': 'ﾚ',
    'ロ': 'ﾛ',
    'ワ': 'ﾜ',
    'ヲ': 'ｦ',
    'ン': 'ﾝ',
    'ァ': 'ｧ',
    'ィ': 'ｨ',
    'ゥ': 'ｩ',
    'ェ': 'ｪ',
    'ォ': 'ｫ',
    'ッ': 'ｯ',
    'ャ': 'ｬ',
    'ュ': 'ｭ',
    'ョ': 'ｮ',
    '。': '｡',
    '、': '､',
    'ー': 'ｰ',
    '「': '｢',
    '」': '｣',
    '・': '･'
}

def ASC(text: str) -> str:
    '''Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters'''
    output = ''
    for char in text:
        code = ord(char)
        if code == 0x3000:
            output += chr(0x20)
        elif code >= (0x21 + 0xFEE0) and code <= (0x7F + 0xFEE0):
            output += chr(code - 0xFEE0)
        elif char in FULL_HALF_WIDTH_MAP.keys():
            output += FULL_HALF_WIDTH_MAP[char]
        else:
            output += char
    return output

def ASIN():
    '''Returns the arcsine of a number'''
    pass

def ASINH():
    ''''''
    pass

def ATAN():
    ''''''
    pass

def ATAN2():
    ''''''
    pass

def ATANH():
    ''''''
    pass

def AVEDEV():
    ''''''
    pass

def AVERAGE():
    ''''''
    pass

def AVERAGEA():
    ''''''
    pass

def AVERAGEIF():
    ''''''
    pass

def AVERAGEIFS():
    ''''''
    pass

TH_0      = 'ศูนย์'
TH_1      = 'หนึ่ง'
TH_2      = 'สอง'
TH_3      = 'สาม'
TH_4      = 'สี่'
TH_5      = 'ห้า'
TH_6      = 'หก'
TH_7      = 'เจ็ด'
TH_8      = 'แปด'
TH_9      = 'เก้า'
TH_10     = 'สิบ'
TH_11     = 'เอ็ด'
TH_20     = 'ยี่'
TH_1E2    = 'ร้อย'
TH_1E3    = 'พัน'
TH_1E4    = 'หมื่น'
TH_1E5    = 'แสน'
TH_1E6    = 'ล้าน'
TH_DOT0   = 'ถ้วน'   # WHOLE_NUMBER_TEXT
TH_BAHT   = 'บาท'   # PRIMARY_UNIT
TH_SATANG = 'สตางค์' # SECONDARY_UNIT
TH_MINUS  = 'ลบ'

'ศูนย์ บาทถ้วน'

TH_DIGITS = [TH_0, TH_1, TH_2, TH_3, TH_4, TH_5, TH_6, TH_7, TH_8, TH_9, TH_10]
TH_UNITS = [TH_10, TH_1E2, TH_1E3, TH_1E4, TH_1E5, TH_1E6]

MAX_POSITION = 6
UNIT_POSITION = 0
TEN_POSITION = 1

_is_zero_value = lambda number : number == 0
_is_unit_position = lambda position : position == UNIT_POSITION
_is_ten_position = lambda position : position % MAX_POSITION == TEN_POSITION
_is_millions_position = lambda position : (position >= MAX_POSITION and position % MAX_POSITION == 0)
_is_last_position = lambda position, length : position + 1 < length

def _get_digit(position, number, length):
    number_text = TH_DIGITS[number]
    if _is_zero_value(number):
        return ''
    if _is_ten_position(position) and number == 1:
        number_text = ''
    if _is_ten_position(position) and number == 2:
        number_text = TH_20
    if _is_millions_position(position) and _is_last_position(position, length) and number == 1:
        number_text = TH_11
    if length == 2 and _is_last_position(position, length) and number == 1:
        number_text = TH_11
    if length > 1 and _is_unit_position(position) and number == 1:
        number_text = TH_11
    return number_text

def _get_unit(position, number):
    unit_text = ''
    if not _is_unit_position(position):
        unit_text = TH_UNITS[abs(position - 1) % MAX_POSITION]
    if _is_zero_value(number) and not _is_millions_position(position):
        unit_text = ''
    return unit_text


def _get_int_output(digits):
    output = ''
    for pos, digit in enumerate(digits[::-1]):
        output = f'{_get_digit(pos, int(digit), len(digits))}{_get_unit(pos, digit)}{output}'
    return output

def BAHTTEXT(number):
    output = ''
    negative = number < 0
    number = abs(number)
    formatted = '{:.2f}'.format(float(number))
    int_digits, frac_digits = [str(d) for d in formatted.split('.')]
    int_output = _get_int_output(int_digits)
    frac_output = _get_int_output(frac_digits)
    if int_digits == '0' and frac_digits == '00':
        output += TH_0 + TH_BAHT + TH_DOT0
    if int_output:
        output += int_output + TH_BAHT
    if int_output and frac_digits == '00':
        output += TH_DOT0
    if frac_digits != '00' and frac_output:
        output += frac_output + TH_SATANG
    if negative:
        output = TH_MINUS + output
    return output

BASE_DIGITS = string.digits + string.ascii_uppercase

def BASE(number, radix, min_length=0):
    if radix < 2 or radix > 36:
        raise _ExcelError('Valid values are between 2 and 36 inclusive.')
    value = int(str(number))
    if value < 0 or value > 9.0072E+15:
        raise _ExcelError('Valid values are between 0 and 9.0072E+15 inclusive.')
    output = []
    while value:
        output.append(BASE_DIGITS[value % radix])
        value = value // radix
    output.reverse()
    return ''.join(output)

def BESSELI(x, n):
    try: x, n = int(x), int(n)
    except: raise _ExcelError('Values are nonnumeric')
    if n < 0: raise _ExcelError('Value has to be greater than or equal to zero')

def BESSELJ():
    ''''''
    pass

def BESSELK():
    ''''''
    pass

def BESSELY():
    ''''''
    pass

def BETADIST():
    ''''''
    pass

def BETA_DIST():
    ''''''
    pass

def BETAINV():
    ''''''
    pass

def BETA_INV():
    ''''''
    pass

def BIN2DEC(number):
    number = int(str(number), 2)
    return str(number) if number < 512 else str(number - 1024)

def BIN2HEX(number, places=0):
    return DEC2HEX(BIN2DEC(number))

def BIN2OCT(number, places=0):
    return DEC2OCT(BIN2DEC(number))

def BINOMDIST():
    ''''''
    pass

def BINOM_DIST():
    ''''''
    pass

def BINOM_DIST_RANGE():
    ''''''
    pass

def BINOM_INV():
    ''''''
    pass

def BITAND():
    ''''''
    pass

def BITLSHIFT():
    ''''''
    pass

def BITOR():
    ''''''
    pass

def BITRSHIFT():
    ''''''
    pass

def BITXOR():
    ''''''
    pass

def CALL():
    ''''''
    pass

def CEILING():
    ''''''
    pass

def CEILING_MATH():
    ''''''
    pass

def CEILING_PRECISE():
    ''''''
    pass

def CELL():
    ''''''
    pass

def CHAR(number):
    try: cnumber = int(number)
    except: raise _ExcelError('Input has to be a number')
    cnumber = _clamp(cnumber, 1, 255)
    return chr(cnumber)

def CHIDIST():
    ''''''
    pass

def CHIINV():
    ''''''
    pass

def CHITEST():
    ''''''
    pass

def CHISQ_DIST():
    ''''''
    pass

def CHISQ_DIST_RT():
    ''''''
    pass

def CHISQ_INV():
    ''''''
    pass

def CHISQ_INV_RT():
    ''''''
    pass

def CHISQ_TEST():
    ''''''
    pass

def CHOOSE():
    ''''''
    pass

def CLEAN():
    ''''''
    pass

def CODE():
    ''''''
    pass

def COLUMN():
    ''''''
    pass

def COLUMNS():
    ''''''
    pass

def COMBIN():
    ''''''
    pass

def COMBINA():
    ''''''
    pass

def COMPLEX():
    ''''''
    pass

def CONCAT(*params):
    return CONCATENATE(*params)
    pass

def CONCATENATE(*params):
    return ''.join(params)
    pass

def CONFIDENCE():
    ''''''
    pass

def CONFIDENCE_NORM():
    ''''''
    pass

def CONFIDENCE_T():
    ''''''
    pass

def CONVERT():
    ''''''
    pass

def CORREL():
    ''''''
    pass

def COS():
    ''''''
    pass

def COSH():
    ''''''
    pass

def COT():
    ''''''
    pass

def COTH():
    ''''''
    pass

def COUNT():
    ''''''
    pass

def COUNTA():
    ''''''
    pass

def COUNTBLANK():
    ''''''
    pass

def COUNTIF():
    ''''''
    pass

def COUNTIFS():
    ''''''
    pass

def COUPDAYBS():
    ''''''
    pass

def COUPDAYS():
    ''''''
    pass

def COUPDAYSNC():
    ''''''
    pass

def COUPNCD():
    ''''''
    pass

def COUPNUM():
    ''''''
    pass

def COUPPCD():
    ''''''
    pass

def COVAR():
    ''''''
    pass

def COVARIANCE_P():
    ''''''
    pass

def COVARIANCE_S():
    ''''''
    pass

def CRITBINOM():
    ''''''
    pass

def CSC():
    ''''''
    pass

def CSCH():
    ''''''
    pass

def CUBEKPIMEMBER():
    ''''''
    pass

def CUBEMEMBER():
    ''''''
    pass

def CUBEMEMBERPROPERTY():
    ''''''
    pass

def CUBERANKEDMEMBER():
    ''''''
    pass

def CUBESET():
    ''''''
    pass

def CUBESETCOUNT():
    ''''''
    pass

def CUBEVALUE():
    ''''''
    pass

def CUMIPMT():
    ''''''
    pass

def CUMPRINC():
    ''''''
    pass

def DATE(year, month, day):
    return DATEVALUE(f'{month}/{day}/{year}')

def DATEDIF():
    ''''''
    pass

def DATEVALUE(date_text: str):
    date_text = date_text.upper()
    if re.match(r'[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}', date_text):
        split = date_text.split('/')
        day = int(split[1])
        month = int(split[0])
        year = int(split[2])
    elif re.match(r'[0-9]{1,2}-[A-Z]{3}-[0-9]{4}', date_text):
        split = date_text.split('-')
        day = int(split[0])
        month = MONTH_VALUES[split[1]]
        year = int(split[2])
    elif re.match(r'[0-9]{4}/[0-9]{1,2}/[0-9]{1,2}', date_text):
        split = date_text.split('/')
        day = int(split[2])
        month = int(split[1])
        year = int(split[0])
    elif re.match(r'[0-9]{1,2}/[A-Z]{3}', date_text):
        split = date_text.split('/')
        day = int(split[0])
        month = MONTH_VALUES[split[1]]
        year = datetime.date.today().year
    return datetime.datetime(year, month, day).toordinal() - PYTHON_EXCEL_ORDINAL_DIFF

def DAVERAGE():
    ''''''
    pass

def DAY():
    ''''''
    pass

def DAYS():
    ''''''
    pass

def DAYS360():
    ''''''
    pass

def DB():
    ''''''
    pass

def DBCS():
    ''''''
    pass

def DCOUNT():
    ''''''
    pass

def DCOUNTA():
    ''''''
    pass

def DDB():
    ''''''
    pass

def DEC2BIN(number, places=0):
    '''Converts a decimal number to binary'''
    return bin(int(str(number)))[2:]

def DEC2HEX(number, places=0):
    '''Converts a decimal number to hexadecimal'''
    return hex(int(str(number)))[2:].upper()

def DEC2OCT(number, places=0):
    '''Converts a decimal number to octal'''
    return oct(int(str(number)))[2:]

def DECIMAL():
    ''''''
    pass

def DEGREES(angle):
    '''Converts radians into degrees.'''
    return angle * 180 / PI()

def DELTA():
    ''''''
    pass

def DEVSQ():
    ''''''
    pass

def DGET():
    ''''''
    pass

def DISC():
    ''''''
    pass

def DMAX():
    ''''''
    pass

def DMIN():
    ''''''
    pass

def DOLLAR():
    ''''''
    pass

def DOLLARDE():
    ''''''
    pass

def DOLLARFR():
    ''''''
    pass

def DPRODUCT():
    ''''''
    pass

def DSTDEV():
    ''''''
    pass

def DSTDEVP():
    ''''''
    pass

def DSUM():
    ''''''
    pass

def DURATION():
    ''''''
    pass

def DVAR():
    ''''''
    pass

def DVARP():
    ''''''
    pass

def EDATE():
    ''''''
    pass

def EFFECT():
    ''''''
    pass

def ENCODEURL():
    ''''''
    pass

def EOMONTH():
    ''''''
    pass

def ERF():
    ''''''
    pass

def ERF_PRECISE():
    ''''''
    pass

def ERFC():
    ''''''
    pass

def ERFC_PRECISE():
    ''''''
    pass

def ERROR_TYPE():
    ''''''
    pass

def EUROCONVERT():
    ''''''
    pass

def EVEN():
    ''''''
    pass

def EXACT():
    ''''''
    pass

def EXP():
    ''''''
    pass

def EXPON_DIST():
    ''''''
    pass

def EXPONDIST():
    ''''''
    pass

def FACT():
    ''''''
    pass

def FACTDOUBLE():
    ''''''
    pass

def FALSE():
    ''''''
    pass

def F_DIST():
    ''''''
    pass

def FDIST():
    ''''''
    pass

def F_DIST_RT():
    ''''''
    pass

def FILTER():
    ''''''
    pass

def FILTERXML():
    ''''''
    pass

def FIND():
    ''''''
    pass

def FINDB():
    ''''''
    pass

def F_INV():
    ''''''
    pass

def F_INV_RT():
    ''''''
    pass

def FINV():
    ''''''
    pass

def FISHER(x):
    try: cx = float(x)
    except: raise _ExcelError('Input is not a numeric value')
    if cx <= -1 or cx >= 1: raise _ExcelError('Value is out of bounds: [-1, 1]')
    return (1 / 2) * math.log((1 + cx) / (1 - cx))

def FISHERINV(y):
    try: cy = float(y)
    except: raise _ExcelError('Input is not a numeric value')
    return ((math.e ** (2 * y)) - 1) / ((math.e ** (2 * y)) + 1)

def FIXED():
    ''''''
    pass

def FLOOR(x, significance=1):
    try: cx = float(x)
    except: raise _ExcelError('Input is not a numerical value')
    try: csignificance = float(significance)
    except: raise _ExcelError('Significance is not a numerical value')
    return math.floor(cx / csignificance) * csignificance

def FLOOR_MATH(x, significance=1, mode=1):
    try: cx = float(x)
    except: raise _ExcelError('Input is not a numerical value')
    try: csignificance = float(significance)
    except: raise _ExcelError('Significance is not a numerical value')
    if abs(mode) != 1: raise _ExcelError('Mode is not 1 or -1')
    if mode == 1 or cx > 0: return math.floor(cx / csignificance) * csignificance
    return math.ceil(cx / csignificance) * csignificance

def FLOOR_PRECISE():
    ''''''
    pass

def FORECAST():
    ''''''
    pass

def FORECAST_ETS():
    ''''''
    pass

def FORECAST_ETS_CONFINT():
    ''''''
    pass

def FORECAST_ETS_SEASONALITY():
    ''''''
    pass

def FORECAST_ETS_STAT():
    ''''''
    pass

def FORECAST_LINEAR():
    ''''''
    pass

def FORMULATEXT():
    ''''''
    pass

def FREQUENCY():
    ''''''
    pass

def F_TEST():
    ''''''
    pass

def FTEST():
    ''''''
    pass

def FV():
    ''''''
    pass

def FVSCHEDULE():
    ''''''
    pass

def GAMMA():
    ''''''
    pass

def GAMMA_DIST():
    ''''''
    pass

def GAMMADIST():
    ''''''
    pass

def GAMMA_INV():
    ''''''
    pass

def GAMMAINV():
    ''''''
    pass

def GAMMALN():
    ''''''
    pass

def GAMMALN_PRECISE():
    ''''''
    pass

def _erf(x):
    '''Error function'''
    cof = [
        -1.3026537197817094, 6.4196979235649026e-1, 1.9476473204185836e-2, -9.561514786808631e-3,
        -9.46595344482036e-4, 3.66839497852761e-4, 4.2523324806907e-5, -2.0278578112534e-5,
        -1.624290004647e-6, 1.303655835580e-6, 1.5626441722e-8, -8.5238095915e-8,
        6.529054439e-9, 5.059343495e-9, -9.91364156e-10, -2.27365122e-10,
        9.6467911e-11, 2.394038e-12, -6.886027e-12, 8.94487e-13,
        3.13092e-13, -1.12708e-13, 3.81e-16, 7.106e-15,
        -1.523e-15, -9.4e-17, 1.21e-16, -2.8e-17
    ]
    isneg = False
    d = 0
    dd = 0
    t, ty, tmp, res = None, None, None, None
    if x < 0:
        x = -x
        isneg = True
    t = 2 / (2 + x)
    ty = 4 * t - 2
    for j in range(len(cof) - 1, 0, -1):
        tmp = d
        d = ty * d - dd + cof[j]
        dd = tmp
    res = t * math.exp(-x * x + 0.5 * (cof[0] + ty * d) - dd)
    return res - 1 if isneg else 1 - res


def _cdf(x, mean, std):
    '''Cumulative distribution function'''
    return 0.5 * (1 + _erf((x - mean) / math.sqrt(2 * std * std)))

def GAUSS(z):
    return _cdf(z, 0, 1) - 0.5

def _gcd(a, b):
    while b:
        a, b = b, a % b
    return a

def GCD(number, *numbers):
    output = _gcd(number, numbers[0])
    for val in numbers[1:]:
        output = _gcd(output, val)
    return output

def GEOMEAN():
    ''''''
    pass

def GESTEP():
    ''''''
    pass

def GETPIVOTDATA():
    ''''''
    pass

def GROWTH():
    ''''''
    pass

def HARMEAN():
    ''''''
    pass

def HEX2BIN(number, places=0):
    '''Converts a hexadecimal number to binary'''
    return DEC2BIN(HEX2DEC(number))

def HEX2DEC(number):
    '''Converts a hexadecimal number to decimal'''
    return str(int(str(number), 16))

def HEX2OCT(number, places=0):
    '''Converts a hexadecimal number to octal'''
    return DEC2OCT(HEX2DEC(number))

def HLOOKUP():
    ''''''
    pass

def HOUR():
    ''''''
    pass

def HYPERLINK():
    ''''''
    pass

def HYPGEOM_DIST():
    ''''''
    pass

def HYPGEOMDIST():
    ''''''
    pass

def IF():
    ''''''
    pass

def IFERROR():
    ''''''
    pass

def IFNA():
    ''''''
    pass

def IFS():
    ''''''
    pass

def IMABS():
    ''''''
    pass

def IMAGINARY():
    ''''''
    pass

def IMARGUMENT():
    ''''''
    pass

def IMCONJUGATE():
    ''''''
    pass

def IMCOS():
    ''''''
    pass

def IMCOSH():
    ''''''
    pass

def IMCOT():
    ''''''
    pass

def IMCSC():
    ''''''
    pass

def IMCSCH():
    ''''''
    pass

def IMDIV():
    ''''''
    pass

def IMEXP():
    ''''''
    pass

def IMLN():
    ''''''
    pass

def IMLOG10():
    ''''''
    pass

def IMLOG2():
    ''''''
    pass

def IMPOWER():
    ''''''
    pass

def IMPRODUCT():
    ''''''
    pass

def IMREAL():
    ''''''
    pass

def IMSEC():
    ''''''
    pass

def IMSECH():
    ''''''
    pass

def IMSIN():
    ''''''
    pass

def IMSINH():
    ''''''
    pass

def IMSQRT():
    ''''''
    pass

def IMSUB():
    ''''''
    pass

def IMSUM():
    ''''''
    pass

def IMTAN():
    ''''''
    pass

def INDEX():
    ''''''
    pass

def INDIRECT():
    ''''''
    pass

def INFO():
    ''''''
    pass

def INT():
    ''''''
    pass

def INTERCEPT():
    ''''''
    pass

def INTRATE():
    ''''''
    pass

def IPMT():
    ''''''
    pass

def IRR():
    ''''''
    pass

def ISBLANK():
    ''''''
    pass

def ISERR():
    ''''''
    pass

def ISERROR():
    ''''''
    pass

def ISEVEN():
    ''''''
    pass

def ISFORMULA():
    ''''''
    pass

def ISLOGICAL():
    ''''''
    pass

def ISNA():
    ''''''
    pass

def ISNONTEXT():
    ''''''
    pass

def ISNUMBER():
    ''''''
    pass

def ISODD():
    ''''''
    pass

def ISREF():
    ''''''
    pass

def ISTEXT():
    ''''''
    pass

def ISO_CEILING():
    ''''''
    pass

def ISOWEEKNUM():
    ''''''
    pass

def ISPMT():
    ''''''
    pass

def JIS(text: str) -> str:
    '''Converts half-width (single-byte) letters within a character string to full-width (double-byte) characters.'''
    output = ''
    char_list = list(text)
    char_list = char_list[::-1]
    carry = ''
    for i in range(len(char_list)):
        char = char_list[i]
        code = ord(char)
        if code in [65438, 65439]:
            carry = char
            continue
        outchar = char + carry
        if code == 0x20:
            output = chr(0x3000) + output
        elif code >= 0x21 and code <= 0x7F:
            output = chr(code + 0xFEE0) + output
        elif outchar in FULL_HALF_WIDTH_MAP.values():
            output = list(filter(lambda x: FULL_HALF_WIDTH_MAP[x] == outchar, FULL_HALF_WIDTH_MAP))[0] + output
        else:
            output = char + output
        if code not in [65438, 65439]:
            carry = ''
    return output

def KURT(number: int | str, *numbers: int | str):
    '''Returns the kurtosis of a data set.'''
    try:
        data = [int(val) for val in [number] + list(numbers)]
    except:
        raise TypeError('Only integers are allowed')
    mean = sum(data) / len(data)
    n = len(data)
    s = math.sqrt((sum([(x - mean) ** 2 for x in data])) / (n - 1))
    first_constant = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))
    main_sum = sum([(((x - mean) / s) ** 4) for x in data])
    second_constant = (3 * ((n - 1) ** 2)) / ((n - 2) * (n - 3))
    return (first_constant * main_sum) - second_constant

def LARGE():
    ''''''
    pass

def LCM():
    ''''''
    pass

def LEFT():
    ''''''
    pass

def LEFTB():
    ''''''
    pass

def LEN():
    ''''''
    pass

def LENB():
    ''''''
    pass

def LET():
    ''''''
    pass

def LINEST():
    ''''''
    pass

def LN():
    ''''''
    pass

def LOG():
    ''''''
    pass

def LOG10():
    ''''''
    pass

def LOGEST():
    ''''''
    pass

def LOGINV():
    ''''''
    pass

def LOGNORM_DIST():
    ''''''
    pass

def LOGNORMDIST():
    ''''''
    pass

def LOGNORM_INV():
    ''''''
    pass

def LOOKUP():
    ''''''
    pass

def LOWER():
    ''''''
    pass

def MATCH():
    ''''''
    pass

def MAX():
    ''''''
    pass

def MAXA():
    ''''''
    pass

def MAXIFS():
    ''''''
    pass

def MDETERM():
    ''''''
    pass

def MDURATION():
    ''''''
    pass

def MEDIAN():
    ''''''
    pass

def MID():
    ''''''
    pass

def MIDB():
    ''''''
    pass

def MIN():
    ''''''
    pass

def MINIFS():
    ''''''
    pass

def MINA():
    ''''''
    pass

def MINUTE():
    ''''''
    pass

def MINVERSE():
    ''''''
    pass

def MIRR():
    ''''''
    pass

def MMULT():
    ''''''
    pass

def MOD():
    ''''''
    pass

def MODE():
    ''''''
    pass

def MODE_MULT():
    ''''''
    pass

def MODE_SNGL():
    ''''''
    pass

def MONTH():
    ''''''
    pass

def MROUND():
    ''''''
    pass

def MULTINOMIAL():
    ''''''
    pass

def MUNIT():
    ''''''
    pass

def N():
    ''''''
    pass

def NA():
    ''''''
    pass

def NEGBINOM_DIST():
    ''''''
    pass

def NEGBINOMDIST():
    ''''''
    pass

def NETWORKDAYS():
    ''''''
    pass

def NETWORKDAYS_INTL():
    ''''''
    pass

def NOMINAL():
    ''''''
    pass

def NORM_DIST():
    ''''''
    pass

def NORMDIST():
    ''''''
    pass

def NORMINV():
    ''''''
    pass

def NORM_INV():
    ''''''
    pass

def NORM_S_DIST():
    ''''''
    pass

def NORMSDIST():
    ''''''
    pass

def NORM_S_INV():
    ''''''
    pass

def NORMSINV():
    ''''''
    pass

def NOT():
    ''''''
    pass

def NOW():
    ''''''
    pass

def NPER():
    ''''''
    pass

def NPV():
    ''''''
    pass

def NUMBERVALUE():
    ''''''
    pass

def OCT2BIN(number):
    '''Converts an octal number to binary'''
    return DEC2BIN(OCT2DEC(number))

def OCT2DEC(number):
    '''Converts an octal number to decimal'''
    return str(int(str(number), 8))

def OCT2HEX(number):
    '''Converts an octal number to hexadecimal'''
    return DEC2HEX(OCT2DEC(number))

def ODD():
    ''''''
    pass

def ODDFPRICE():
    ''''''
    pass

def ODDFYIELD():
    ''''''
    pass

def ODDLPRICE():
    ''''''
    pass

def ODDLYIELD():
    ''''''
    pass

def OFFSET():
    ''''''
    pass

def OR():
    ''''''
    pass

def PDURATION():
    ''''''
    pass

def PEARSON():
    ''''''
    pass

def PERCENTILE_EXC():
    ''''''
    pass

def PERCENTILE_INC():
    ''''''
    pass

def PERCENTILE():
    ''''''
    pass

def PERCENTRANK_EXC():
    ''''''
    pass

def PERCENTRANK_INC():
    ''''''
    pass

def PERCENTRANK():
    ''''''
    pass

def PERMUT():
    ''''''
    pass

def PERMUTATIONA():
    ''''''
    pass

def PHI():
    ''''''
    pass

def PHONETIC():
    ''''''
    pass

def PI():
    return 3.14159265358979

def PMT():
    ''''''
    pass

def POISSON_DIST():
    ''''''
    pass

def POISSON():
    ''''''
    pass

def POWER():
    ''''''
    pass

def PPMT():
    ''''''
    pass

def PRICE():
    ''''''
    pass

def PRICEDISC():
    ''''''
    pass

def PRICEMAT():
    ''''''
    pass

def PROB():
    ''''''
    pass

def PRODUCT():
    ''''''
    pass

def PROPER():
    ''''''
    pass

def PV():
    ''''''
    pass

def QUARTILE():
    ''''''
    pass

def QUARTILE_EXC():
    ''''''
    pass

def QUARTILE_INC():
    ''''''
    pass

def QUOTIENT():
    ''''''
    pass

def RADIANS():
    ''''''
    pass

def RAND():
    ''''''
    pass

def RANDARRAY():
    ''''''
    pass

def RANDBETWEEN():
    ''''''
    pass

def RANK_AVG():
    ''''''
    pass

def RANK_EQ():
    ''''''
    pass

def RANK():
    ''''''
    pass

def RATE():
    ''''''
    pass

def RECEIVED():
    ''''''
    pass

def REGISTER_ID():
    ''''''
    pass

def REPLACE():
    ''''''
    pass

def REPLACEB():
    ''''''
    pass

def REPT():
    ''''''
    pass

def RIGHT():
    ''''''
    pass

def RIGHTB():
    ''''''
    pass

def ROMAN():
    ''''''
    pass

def ROUND():
    ''''''
    pass

def ROUNDDOWN():
    ''''''
    pass

def ROUNDUP():
    ''''''
    pass

def ROW():
    ''''''
    pass

def ROWS():
    ''''''
    pass

def RRI():
    ''''''
    pass

def RSQ():
    ''''''
    pass

def RTD():
    ''''''
    pass

def SEARCH():
    ''''''
    pass

def SEARCHB():
    ''''''
    pass

def SEC():
    ''''''
    pass

def SECH():
    ''''''
    pass

def SECOND():
    ''''''
    pass

def SEQUENCE():
    ''''''
    pass

def SERIESSUM():
    ''''''
    pass

def SHEET():
    ''''''
    pass

def SHEETS():
    ''''''
    pass

def SIGN():
    ''''''
    pass

def SIN():
    ''''''
    pass

def SINH():
    ''''''
    pass

def SKEW():
    ''''''
    pass

def SKEW_P():
    ''''''
    pass

def SLN():
    ''''''
    pass

def SLOPE():
    ''''''
    pass

def SMALL():
    ''''''
    pass

def SORT():
    ''''''
    pass

def SORTBY():
    ''''''
    pass

def SQRT():
    ''''''
    pass

def SQRTPI():
    ''''''
    pass

def STANDARDIZE():
    ''''''
    pass

def STOCKHISTORY():
    ''''''
    pass

def STDEV():
    ''''''
    pass

def STDEV_P():
    ''''''
    pass

def STDEV_S():
    ''''''
    pass

def STDEVA():
    ''''''
    pass

def STDEVP():
    ''''''
    pass

def STDEVPA():
    ''''''
    pass

def STEYX():
    ''''''
    pass

def SUBSTITUTE():
    ''''''
    pass

def SUBTOTAL():
    ''''''
    pass

def SUM():
    ''''''
    pass

def SUMIF():
    ''''''
    pass

def SUMIFS():
    ''''''
    pass

def SUMPRODUCT():
    ''''''
    pass

def SUMSQ():
    ''''''
    pass

def SUMX2MY2():
    ''''''
    pass

def SUMX2PY2():
    ''''''
    pass

def SUMXMY2():
    ''''''
    pass

def SWITCH():
    ''''''
    pass

def SYD():
    ''''''
    pass

def T():
    ''''''
    pass

def TAN():
    ''''''
    pass

def TANH():
    ''''''
    pass

def TBILLEQ():
    ''''''
    pass

def TBILLPRICE():
    ''''''
    pass

def TBILLYIELD():
    ''''''
    pass

def T_DIST():
    ''''''
    pass

def T_DIST_2T():
    ''''''
    pass

def T_DIST_RT():
    ''''''
    pass

def TDIST():
    ''''''
    pass

def TEXT():
    ''''''
    pass

def TEXTJOIN():
    ''''''
    pass

def TIME():
    ''''''
    pass

def TIMEVALUE():
    ''''''
    pass

def T_INV():
    ''''''
    pass

def T_INV_2T():
    ''''''
    pass

def TINV():
    ''''''
    pass

def TODAY():
    ''''''
    pass

def TRANSPOSE():
    ''''''
    pass

def TREND():
    ''''''
    pass

def TRIM():
    ''''''
    pass

def TRIMMEAN():
    ''''''
    pass

def TRUE():
    ''''''
    pass

def TRUNC():
    ''''''
    pass

def T_TEST():
    ''''''
    pass

def TTEST():
    ''''''
    pass

def TYPE():
    ''''''
    pass

def UNICHAR():
    ''''''
    pass

def UNICODE():
    ''''''
    pass

def UNIQUE():
    ''''''
    pass

def UPPER():
    ''''''
    pass

def VALUE():
    ''''''
    pass

def VALUETOTEXT():
    ''''''
    pass

def VAR():
    ''''''
    pass

def VAR_P():
    ''''''
    pass

def VAR_S():
    ''''''
    pass

def VARA():
    ''''''
    pass

def VARP():
    ''''''
    pass

def VARPA():
    ''''''
    pass

def VDB():
    ''''''
    pass

def VLOOKUP():
    ''''''
    pass

def WEBSERVICE():
    ''''''
    pass

def WEEKDAY():
    ''''''
    pass

def WEEKNUM():
    ''''''
    pass

def WEIBULL():
    ''''''
    pass

def WEIBULL_DIST():
    ''''''
    pass

def WORKDAY():
    ''''''
    pass

def WORKDAY_INTL():
    ''''''
    pass

def XIRR():
    ''''''
    pass

def XLOOKUP():
    ''''''
    pass

def XMATCH():
    ''''''
    pass

def XNPV():
    ''''''
    pass

def XOR():
    ''''''
    pass

def YEAR():
    ''''''
    pass

def YEARFRAC():
    ''''''
    pass

def YIELD(settlement, maturity, rate, pr, redemption, frequency, basis=0):
    ''''''
    pass

def YIELDDISC():
    ''''''
    pass

def YIELDMAT():
    ''''''
    pass

def Z_TEST():
    ''''''
    pass

def ZTEST():
    ''''''
    pass
