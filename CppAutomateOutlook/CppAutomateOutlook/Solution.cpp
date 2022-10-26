#include "stdafx.h"

#define _WIN32_DCOM
#include <stdio.h>
#include <windows.h>
#include <OAIdl.h>
#include <objbase.h>
#include "Solution.h"
#pragma comment(lib, "Ole32.lib")

//
//   FUNCTION: AutoWrap(int, VARIANT*, IDispatch*, LPOLESTR, int,...)
//
//   PURPOSE: Automation helper function. It simplifies most of the low-level 
//      details involved with using IDispatch directly. Feel free to use it 
//      in your own implementations. One caveat is that if you pass multiple 
//      parameters, they need to be passed in reverse-order.
//
//   PARAMETERS:
//      * autoType - Could be one of these values: DISPATCH_PROPERTYGET, 
//      DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPATCH_METHOD.
//      * pvResult - Holds the return value in a VARIANT.
//      * pDisp - The IDispatch interface.
//      * ptName - The property/method name exposed by the interface.
//      * cArgs - The count of the arguments.
//
//   RETURN VALUE: An HRESULT value indicating whether the function succeeds 
//      or not. 
//
//   EXAMPLE: 
//      AutoWrap(DISPATCH_METHOD, NULL, pDisp, L"call", 2, parm[1], parm[0]);
//
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, 
				 LPOLESTR ptName, int cArgs...) 
{
	// Begin variable-argument list
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) 
	{
		_putws(L"NULL IDispatch passed to AutoWrap()");
		_exit(0);
		return E_INVALIDARG;
	}

	// Variables used
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;

	// Get DISPID for name passed
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr))
	{
		wprintf(L"IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx\n", 
			ptName, hr);
		_exit(0);
		return hr;
	}

	// Allocate memory for arguments
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for(int i=0; i < cArgs; i++) 
	{
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts
	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
		autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) 
	{
		wprintf(L"IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx\n", 
			ptName, dispID, hr);
		_exit(0);
		return hr;
	}

	// End variable-argument section
	va_end(marker);

	delete[] pArgs;

	return hr;
}


DWORD WINAPI AutomateOutlookByCOMAPI(LPVOID lpParam)
{
	// Initializes the COM library on the current thread and identifies 
	// the concurrency model as single-thread apartment (STA). 
	// [-or-] CoInitialize(NULL);
	// [-or-] CoCreateInstance(NULL);
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	// Define vtMissing for optional parameters in some calls.
	VARIANT vtMissing;
	vtMissing.vt = VT_EMPTY;

	// Get CLSID of the server
	CLSID clsid;
	HRESULT hr;

	// Option 1. Get CLSID from ProgID using CLSIDFromProgID.
	LPCOLESTR progID = L"Outlook.Application";
	hr = CLSIDFromProgID(progID, &clsid);
	if (FAILED(hr))
	{
		wprintf(L"CLSIDFromProgID(\"%s\") failed w/err 0x%08lx\n", progID, hr);
		return 1;
	}
	// Option 2. Build the CLSID directly.
	/*const IID CLSID_Application =
	{0x0006F03A,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
	clsid = CLSID_Application;*/

	// Get the IDispatch interface of the running instance

	IUnknown* pUnk = NULL;
	IDispatch* pOutlookApp = NULL;
	hr = GetActiveObject(
		clsid, NULL, (IUnknown**)&pUnk
	);

	if (FAILED(hr))
	{
		wprintf(L"GetActiveObject failed with w/err 0x%08lx\n", hr);
		return 1;
	}

	hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pOutlookApp);
	if (FAILED(hr))
	{
		wprintf(L"QueryInterface failed with w/err 0x%08lx\n", hr);
		return 1;
	}

	_putws(L"Outlook.Application is found");

	IDispatch* comAddins = NULL;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pOutlookApp, L"COMAddins", 0);
		comAddins = result.pdispVal;
	}

	IDispatch* myAddin = NULL;
	{
		VARIANT x;
		x.vt = VT_BSTR;
		x.bstrVal = SysAllocString(L"OutlookAddIn");

		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_METHOD, &result, comAddins, L"Item", 1, x);
		myAddin = result.pdispVal;

		VariantClear(&x);
	}

	IDispatch* myAddinObj = NULL;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, myAddin, L"Object", 0);
		myAddinObj = result.pdispVal;
	}

	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_METHOD, &result, myAddinObj, L"GetArrayFromCSharp", 0);

		
		_putws(result.bstrVal);
	}


	if (comAddins != NULL)
	{
		comAddins->Release();
	}
	if (myAddin != NULL)
	{
		myAddin->Release();
	}
	if (myAddinObj != NULL)
	{
		myAddinObj->Release();
	}
	if (pOutlookApp != NULL)
	{
		pOutlookApp->Release();
	}
	// Uninitialize COM for this thread.
	CoUninitialize();
}