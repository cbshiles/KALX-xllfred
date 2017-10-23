// xllfred.h - Federal Reserve Economic Data
// Copyright (c) 2012 KALX, LLC. All rights reserved. No warranty is made.
// Uncomment the following line to use features for Excel2007 and above.
//#define EXCEL12
#include <ctime>
#include "xll/xll.h"

using namespace xll;

#define FRED_AGENT L"xllfred/0.1"
#define FRED_URL L"api.stlouisfed.org"
#define FRED_API_KEY L"1947fc010512c76edbe4b8f6f8e0bdd5"

#ifndef CATEGORY
#define CATEGORY _T("FRED")
#endif

#define FRED_PREFIX L"/fred/"
#define FRED_CATEGORY FRED_PREFIX L"category?category_id=%d"
#define FRED_CATEGORY_CHILDREN FRED_PREFIX L"category/children?category_id=%d"
#define FRED_CATEGORY_RELATED FRED_PREFIX L"category/related?category_id=%d"
#define FRED_CATEGORY_SERIES FRED_PREFIX L"category/series?category_id=%d"

#define FRED_SERIES FRED_PREFIX L"series?series_id=%s"
#define FRED_SERIES_OBSERVATIONS FRED_PREFIX L"series/observations?series_id=%s"

typedef xll::traits<XLOPERX>::xcstr xcstr; // pointer to const string
typedef xll::traits<XLOPERX>::xstring xstring; // pointer to const string
typedef xll::traits<XLOPERX>::xword xword; // use for OPER and FP indices

// pointer to yyyy-mm-dd from Excel date
inline const wchar_t* fred_date(double t)
{
	static wchar_t buf[11];
	OPER o;

	o = Excel<XLOPER>(xlfYear, OPER(t));
	ensure (o.xltype == xltypeNum);
	int y = static_cast<int>(o.val.num);

	o = Excel<XLOPER>(xlfMonth, OPER(t));
	ensure (o.xltype == xltypeNum);
	int m = static_cast<int>(o.val.num);

	o = Excel<XLOPER>(xlfDay, OPER(t));
	ensure (o.xltype == xltypeNum);
	int d = static_cast<int>(o.val.num);

	_snwprintf_s(buf, sizeof(buf), L"%4d-%02d-%02d", y, m, d);

	return buf; 
}
// non Excel based conversion routions
inline OPERX fred2date(const wchar_t* date)
{
	int y, m, d;

	ensure (3 == swscanf(date, L"%d-%d-%d", &y, &m, &d));

//	if (y < 1970)

	tm tm;
	memset(&tm, 0, sizeof(tm));
	tm.tm_year = y - 1900;
	tm.tm_mon = m - 1;
	tm.tm_mday = d;
	tm.tm_isdst = -1;

	time_t t = mktime(&tm);

	return OPERX(25569 + t/86400.);
}
inline OPERX fred2num(const wchar_t* num)
{
	double x;

//	ensure (1 == swscanf(num, L"%lf", &x));
	if (1 != swscanf(num, L"%lf", &x))
		x = 0;

	return OPERX(x);
}
int fred2excel(const wchar_t* date)
{
	int y, m, d;

	ensure (3 == swscanf(date, L"%d-%d-%d", &y, &m, &d));

	if (d == 29 && m == 02 && y==1900)
        return 60;

    long e = 
            int(( 1461 * ( y + 4800 + int(( m - 14 ) / 12) ) ) / 4) +
            int(( 367 * ( m - 2 - 12 * ( ( m - 14 ) / 12 ) ) ) / 12) -
            int(( 3 * ( int(( y + 4900 + int(( m - 14 ) / 12) ) / 100) ) ) / 4) +
            d - 2415019 - 32075;

    if (e < 60) {
        --e;
    }

    return static_cast<int>(e);
}