// xllfred.cpp - Federal Reserve Economic Data
#include <sstream>
#include "rapidxml.hpp"
#include "xllfred.h"
#include "xll/utility/message.h"
#include "xll/utility/strings.h"
#include "xll/utility/winhttp.h"

using namespace xll;
using namespace rapidxml;

// put in "xll/utility/strings.h"
// convert char* to xcstr
#ifdef EXCEL12
#define STRX(x) strings::mbstowcs(std::basic_string<char>(x)).c_str()
#else
#define STRX(x) x
#endif

static wchar_t buf[BUFSIZ];

static HINTERNET session()
{
	static WinHttp::Open inet(FRED_AGENT);
	static WinHttp::Connect conn(inet, FRED_URL, INTERNET_DEFAULT_HTTP_PORT);

	return conn;
}

// send request, get data, return parsed document
inline
std::vector<char> request(const wchar_t* query, size_t size = 0)
{
	if (size == 0)
		size = wcslen(query);

	std::basic_string<wchar_t> request(query, size);
	request += L"&api_key=";
	request += FRED_API_KEY;

	WinHttp::OpenRequest req(session(), L"GET", request.c_str());
	ensure (req.SendRequest());
	ensure (req.ReceiveResponse());

	wchar_t status[8];
	DWORD dw = sizeof(status);
	WinHttpQueryHeaders(req, WINHTTP_QUERY_STATUS_CODE, 0, &status, &dw, 0);

	std::vector<char> data;
	while (WinHttpQueryDataAvailable(req, &dw) && dw) {
		data.resize(data.size() + dw);
		WinHttpReadData(req, &data[0] + data.size() - dw, dw, &dw);
	}
	data.push_back(0);

	if (_wtoi(status) != HTTP_STATUS_OK) {
		xml_document<> doc;
		doc.parse<0>(&data[0]);
		xml_node<> *node = doc.first_node("error");
		ensure (node);
		xml_attribute<> *attr = node->first_attribute("message");
		ensure (attr);
		
		throw std::runtime_error(attr->value());
	}


	return data;
}

static AddInX xai_fred(
	DocumentX(CATEGORY)
	.Documentation(
		_T("This add-in gives you access to the Federal Reserve Economic Database ")
		_T("that provides 45,000 economic time series from 41 sources. ")
		_T("If you know the category id then <codeInline>FRED.CATEGORY</codeInline> provides more ")
		_T("information about that category. ")
		_T("Use <codeInline>FRED.CATEGORY.CHILDREN</codeInline> to get the hierarchy of available categories ")
		_T("and <codeInline>FRED.CATEGORY.SERIES</codeInline> to find the name of a particular data series. ")
		_T("<codeInline>FRED.SERIES</codeInline> returns general information about a series and ")
		_T("<codeInline>FRED.SERIES.OBSERVATIONS</codeInline> retrieves the actual data. ")
		,
		xml::externalLink(_T("St. Louis Fed Web Services"), _T("http://api.stlouisfed.org/"))
	)
);

static AddInX xai_fred_category(
	FunctionX(XLL_LPOPERX, _T("?xll_fred_category"), _T("FRED.CATEGORY"))
	.Arg(XLL_USHORTX, _T("id"), _T("is the category id. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get the parent_id and name corresponding to id."))
	.Documentation(_T("Calls fred/category?category_id=id."))
);
LPOPERX WINAPI
xll_fred_category(USHORT id)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = OPERX();

		swprintf_s(buf, BUFSIZ, FRED_CATEGORY, id);
		std::vector<char> data = request(buf);

		xml_document<> doc;
		doc.parse<0>(&data[0]);

		xml_node<> *node = doc.first_node();
		ensure (node);

		for (xml_node<> *cat = node->first_node(); cat; cat = node->next_sibling()) {
			xword r = o.rows();
			o.resize(r + 1, 2);
			for (xml_attribute<> *attr = cat->first_attribute(); attr; attr = attr->next_attribute()) {
				if (!strcmp("parent_id", attr->name()))
					o(r, 0) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
				else if (!strcmp("name", attr->name()))
					o(r, 1) = OPERX(STRX(attr->value()));
			}
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}

static AddInX xai_fred_category_children(
	FunctionX(XLL_LPOPERX, _T("?xll_fred_category_children"), _T("FRED.CATEGORY.CHILDREN"))
	.Arg(XLL_USHORTX, _T("id"), _T("is the category id. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get the category id and name of children."))
	.Documentation(_T("Calls fred/category/children?category_id=id."))
);
LPOPERX WINAPI
xll_fred_category_children(USHORT id)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = OPERX();

		swprintf_s(buf, BUFSIZ, FRED_CATEGORY_CHILDREN, id);
		std::vector<char> data = request(buf);

		xml_document<> doc;
		doc.parse<0>(&data[0]);

		xml_node<> *node = doc.first_node();
		ensure (node);

		for (xml_node<> *cat = node->first_node(); cat; cat = cat->next_sibling()) {
			xword r = o.rows();
			o.resize(r + 1, 2);
			for (xml_attribute<> *attr = cat->first_attribute(); attr; attr = attr->next_attribute()) {
				if (!strcmp("id", attr->name()))
					o(r, 0) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
				else if (!strcmp("name", attr->name()))
					o(r, 1) = OPERX(STRX(attr->value()));
			}
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}

static AddInX xai_fred_category_series(
	FunctionX(XLL_LPOPERX, _T("?xll_fred_category_series"), _T("FRED.CATEGORY.SERIES"))
	.Arg(XLL_USHORTX, _T("id"), _T("is the category id. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get the series id, start, end, frequency, units, and title of category id."))
	.Documentation(_T("Calls fred/category/series?category_id=id."))
);
LPOPERX WINAPI
xll_fred_category_series(USHORT id)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = OPERX();

		swprintf_s(buf, BUFSIZ, FRED_CATEGORY_SERIES, id);
		std::vector<char> data = request(buf);

		xml_document<> doc;
		doc.parse<0>(&data[0]);

		xml_node<> *node = doc.first_node();
		ensure (node);

		for (xml_node<> *ser = node->first_node(); ser; ser = ser->next_sibling()) {
			xword r = o.rows();
			o.resize(r + 1, 6);
			for (xml_attribute<> *attr = ser->first_attribute(); attr; attr = attr->next_attribute()) {
				if (!strcmp("observation_start", attr->name()))
					o(r, 1) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
				else if (!strcmp("observation_end", attr->name()))
					o(r, 2) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
				else if (!strcmp("frequency", attr->name()))
					o(r, 3) = OPERX(STRX(attr->value()));
				else if (!strcmp("id", attr->name()))
					o(r, 0) = OPERX(STRX(attr->value()));
				else if (!strcmp("units", attr->name()))
					o(r, 4) = OPERX(STRX(attr->value()));
				else if (!strcmp("title", attr->name()))
					o(r, 5) = OPERX(STRX(attr->value()));
			}
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}

static AddInX xai_fred_series(
	FunctionX(XLL_LPOPERX, _T("?xll_fred_series"), _T("FRED.SERIES"))
	.Arg(XLL_CSTRINGX, _T("name"), _T("is the series name. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get the start, end, frequency, units and title of series name."))
	.Documentation(_T("Calls fred/series?series_id=id."))
);
LPOPERX WINAPI
xll_fred_series(xcstr series)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = OPERX();

		if (*series == '0') {
			o = ErrX(xlerrNA);
		}
		else {
			swprintf_s(buf, BUFSIZ, FRED_SERIES, strings::WCS(series).c_str());
			std::vector<char> data = request(buf);

			xml_document<> doc;
			doc.parse<0>(&data[0]);

			xml_node<> *node = doc.first_node();
			ensure (node);

			for (xml_node<> *ser = node->first_node(); ser; ser = ser->next_sibling()) {
				xword r = o.rows();
				o.resize(r + 1, 5);
				for (xml_attribute<> *attr = ser->first_attribute(); attr; attr = attr->next_attribute()) {
					if (!strcmp("observation_start", attr->name()))
						o(r, 0) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
					else if (!strcmp("observation_end", attr->name()))
						o(r, 1) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
					else if (!strcmp("frequency", attr->name()))
						o(r, 2) = OPERX(STRX(attr->value()));
					else if (!strcmp("units", attr->name()))
						o(r, 3) = OPERX(STRX(attr->value()));
					else if (!strcmp("title", attr->name()))
						o(r, 4) = OPERX(STRX(attr->value()));
				}
			}
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}


#ifdef EXCEL12

DWORD WINAPI
xla_fred_series_observations(LPVOID arg)
{
	OPER12& dh = *(LPOPER12)arg;

	try {
		ensure (dh[0].xltype & xltypeStr);
		std::vector<char> data = request(dh[0].val.str + 1, dh[0].val.str[0]);
		dh[0] = OPER12();
		OPER12& o(dh[0]);

 		xml_document<> doc;
		doc.parse<0>(&data[0]);

		xml_node<> *node = doc.first_node();
		ensure (node); // observations

		for (xml_node<> *cat = node->first_node(); cat; cat = cat->next_sibling()) {
			xword r = o.rows();
			o.resize(r + 1, 2);
			for (xml_attribute<> *attr = cat->first_attribute(); attr; attr = attr->next_attribute()) {
				if (!strcmp("date", attr->name()))
					o(r, 0) = fred2excel(STRX(attr->value()));
				else if (!strcmp("value", attr->name()))
					o(r, 1) = fred2num(STRX(attr->value()));
			}
		}

		// return result to xll_echoa
		Excel<XLOPER12>(xlAsyncReturn, dh[1], dh[0]); // note handle, then data
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		dh[0] = Err12(xlerrNA);
	}

	delete &dh;

	return 0; // do not use ThreadExit with a C++ dll linked to static CRT
}

static AddIn12 xai_fred_series_observations(
	Function12(XLL_VOID12, L"?xll_fred_series_observations", L"FRED.SERIES.OBSERVATIONS")
	.Arg(XLL_CSTRINGX, L"name", L"is the series name. ")
	.Arg(XLL_DOUBLEX, L"_start", L"is the series observation start time. ")
	.Arg(XLL_DOUBLEX, L"_end", L"is the series observation end time. ")
	.Arg(XLL_BOOLX XLL_ASYNCHRONOUS12, L"_desc", L"return times in descending order. ")
	.Category(CATEGORY)
	.FunctionHelp(L"Get the values of series between start and end time.")
	.Documentation(L"Calls fred/series/observations?series_id=id.")
);
void WINAPI
xll_fred_series_observations(xcstr series, double start, double end, BOOL desc, LPXLOPER12 ph)
{
#pragma XLLEXPORT
	try {
		wchar_t buf[512];
		OPER12& dh = *new OPER12(2, 1);

		// ensure (ph->xltype == xltypeBigData);
		dh[1] = *ph;

		swprintf_s(buf, BUFSIZ, FRED_SERIES_OBSERVATIONS, strings::WCS(series).c_str());
		dh[0] = OPER12(buf);
		
		if (start) {
			dh[0] = Excel<XLOPER12>(xlfConcatenate, dh[0], OPER12(L"&observation_start="));
			dh[0] = Excel<XLOPER12>(xlfConcatenate, dh[0], OPER12(fred_date(start),10));
		}

		if (end) {
			dh[0] = Excel<XLOPER12>(xlfConcatenate, dh[0], OPER12(L"&observation_end="));
			dh[0] = Excel<XLOPER12>(xlfConcatenate, dh[0], OPER12(fred_date(end),10));
		}

		if (desc)
			dh[0] = Excel<XLOPER12>(xlfConcatenate, dh[0], OPER12(L"&sort_order=desc"));

		if (!CreateThread(0, 0, xla_fred_series_observations, &dh, 0, 0)) { // use _begin/_endthread???
			// xll_echow not called
			delete &dh;
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());
	}

	return;
}

#else // EXCEL4

static AddInX xai_fred_series_observations(
	FunctionX(XLL_LPOPERX, _T("?xll_fred_series_observations"), _T("FRED.SERIES.OBSERVATIONS"))
	.Arg(XLL_CSTRINGX, _T("id"), _T("is the series id. "))
	.Arg(XLL_DOUBLEX, _T("start?"), _T("is the series observation start time. "))
	.Arg(XLL_DOUBLEX, _T("end?"), _T("is the series observation end time. "))
	.Arg(XLL_BOOLX, _T("desc?"), _T("return times in descending order. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get the values of series between start and end time."))
	.Documentation(_T("Calls fred/series/observations?series_id=id."))
);
LPOPERX WINAPI
xll_fred_series_observations(xcstr series, double start, double end, BOOL desc)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = OPERX();

		swprintf_s(buf, BUFSIZ, FRED_SERIES_OBSERVATIONS, strings::WCS(series).c_str());
		
		if (start) {
			wcscat_s(buf, L"&observation_start=");
			wcsncat_s(buf, fred_date(start), 10);
		}

		if (end) {
			wcscat_s(buf, L"&observation_end=");
			wcsncat_s(buf, fred_date(end), 10);
		}

		if (desc)
			wcscat_s(buf, L"&sort_order=desc");

		std::vector<char> data = request(buf);

 		xml_document<> doc;
		doc.parse<0>(&data[0]);

		xml_node<> *node = doc.first_node();
		ensure (node); // observations

		for (xml_node<> *cat = node->first_node(); cat; cat = cat->next_sibling()) {
			xword r = o.rows();
			o.resize(r + 1, 2);
			for (xml_attribute<> *attr = cat->first_attribute(); attr; attr = attr->next_attribute()) {
				if (!strcmp("date", attr->name()))
					o(r, 0) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
				else if (!strcmp("value", attr->name()))
					o(r, 1) = ExcelX(xlfValue, OPERX(STRX(attr->value())));
			}
		}
	}
	catch (const std::exception& ex) {
		xstring err = Error::Message(GetLastError()).what();
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}

#endif // EXCEL12

inline void
rescale_horizontal_axis(const OPERX& chart, const OPERX& min, const OPERX& max)
{
	ExcelX(xlcSelect, chart);
	ExcelX(xlcSelect, OPERX(_T("Axis 2")));
	ExcelX(xlcScale, min, max);
}

// rescale horizontal chart axis
static AddInX xai_rescale_horizontal_axis(_T("?xll_rescale_horizontal_axis"), _T("FRED.RESCALE.HORIZONTAL.AXIS"));
int WINAPI
xll_rescale_horizontal_axis(void)
{
#pragma XLLEXPORT
	try {
		rescale_horizontal_axis(
			OPERX(_T("Chart 2")), 
			OPERX(ExcelX(xlfEvaluate, OPERX(_T("=FRED!min_date")))), 
			OPERX(ExcelX(xlfEvaluate, OPERX(_T("=FRED!max_date"))))
		);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}
	
	return 1;
}
