// CppAutomateOutlook.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <stdio.h>
#include <windows.h>
#include "Solution.h"


int _tmain(int argc, _TCHAR* argv[])
{
	HANDLE hThread = CreateThread(NULL, 0, AutomateOutlookByCOMAPI2, NULL, 0, NULL);
	WaitForSingleObject(hThread, INFINITE);
	CloseHandle(hThread);
	return 0;
}

