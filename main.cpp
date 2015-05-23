#define UNICODE
#pragma comment(linker,"/opt:nowin98")
#include<windows.h>
#include <ole2.h>

TCHAR szClassName[] = TEXT("CreateExcelFile");

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
	va_list marker;
	va_start(marker, cArgs);
	if (!pDisp)
	{
		MessageBox(NULL, TEXT("NULL IDispatch passed to AutoWrap()"), TEXT("Error"), 0x10010);
		_exit(0);
	}
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	TCHAR buf[200];
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr))
	{
		wsprintf(buf, TEXT("IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx"), ptName, hr);
		MessageBox(NULL, buf, TEXT("AutoWrap()"), 0x10010);
		_exit(0);
		return hr;
	}
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	for (int i = 0; i<cArgs; i++)
	{
		pArgs[i] = va_arg(marker, VARIANT);
	}
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;
	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr))
	{
		wsprintf(buf, TEXT("IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx"), ptName, dispID, hr);
		MessageBox(NULL, buf, TEXT("AutoWrap()"), 0x10010);
		_exit(0);
		return hr;
	}
	va_end(marker);
	delete[] pArgs;
	return hr;
}

BOOL CreateExcelFile(LPCTSTR lpszFilePath)
{
	CoInitialize(NULL);
	CLSID clsid;

	// Get CLSID for our server...
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	if (FAILED(hr))
	{
		::MessageBox(NULL, TEXT("CLSIDFromProgID() failed"), TEXT("Error"), 0x10010);
		return FALSE;
	}

	// Start server and get IDispatch...
	IDispatch *pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pXlApp);
	if (FAILED(hr))
	{
		::MessageBox(NULL, TEXT("Excel not registered properly"), TEXT("Error"), 0x10010);
		return FALSE;
	}

	// Get Workbooks collection
	IDispatch *pXlBooks;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
		pXlBooks = result.pdispVal;
	}

	// Call Workbooks.Add() to get a new workbook...
	IDispatch *pXlBook;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Add", 0);
		pXlBook = result.pdispVal;
	}

	// Create a 15x15 safearray of variants...
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[2];
		sab[0].lLbound = 1; sab[0].cElements = 15;
		sab[1].lLbound = 1; sab[1].cElements = 15;
		arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	}

	// Fill safearray with some values...
	for (int i = 1; i <= 15; i++)
	{
		for (int j = 1; j <= 15; j++)
		{
			// Create entry value for (i,j)
			VARIANT tmp;
			tmp.vt = VT_I4;
			tmp.lVal = i*j;
			// Add to safearray...
			long indices[] = { i, j };
			SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
		}
	}

	// Get ActiveSheet object
	IDispatch *pXlSheet;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
		pXlSheet = result.pdispVal;
	}

	// Get Range object for the Range A1:O15...
	IDispatch *pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(L"A1:O15");

		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
		VariantClear(&parm);

		pXlRange = result.pdispVal;
	}

	// Set range with our safearray...
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, arr);

	//Save the work book.
	{
		VARIANT result;
		VariantInit(&result);
		VARIANT fname;
		fname.vt = VT_BSTR;
		fname.bstrVal = ::SysAllocString(lpszFilePath);
		AutoWrap(DISPATCH_METHOD, &result, pXlSheet, L"SaveAs", 1, fname);
	}

	// Tell Excel to quit (i.e. App.Quit)
	AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);

	// Release references...
	pXlRange->Release();
	pXlSheet->Release();
	pXlBook->Release();
	pXlBooks->Release();
	pXlApp->Release();
	VariantClear(&arr);

	// Uninitialize COM for this thread...
	CoUninitialize();

	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	switch (msg)
	{
	case WM_CREATE:
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("Excelƒtƒ@ƒCƒ‹‚ðì¬"), WS_VISIBLE | WS_CHILD, 10, 10, 256, 32, hWnd, (HMENU)100, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 100)
		{
			EnableWindow(hButton, FALSE);
			OPENFILENAME ofn = { sizeof(OPENFILENAME) };
			TCHAR szFileName[MAX_PATH] = TEXT("");
			ofn.hwndOwner = hWnd;
			ofn.lpstrFilter = TEXT("Excel Files (*.xlsx)\0*.xlsx\0All Files (*.*)\0*.*\0");
			ofn.lpstrFile = szFileName;
			ofn.nMaxFile = MAX_PATH;
			ofn.lpstrDefExt = TEXT("xlsx");
			ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
			if (GetSaveFileName(&ofn))
			{
				CreateExcelFile(szFileName);
			}
			EnableWindow(hButton, TRUE);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0, IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("CreateExcelFile"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
		);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return msg.wParam;
}
