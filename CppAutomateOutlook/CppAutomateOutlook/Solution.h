#pragma once


//   FUNCTION: AutomateOutlookByCOMAPI(LPVOID)
//
//   PURPOSE: Automate Microsoft Outlook using C++ and the COM APIs.
//
//   PARAMETERS:
//      * lpParam - The thread data passed to the function using the 
//      lpParameter parameter when creating a thread. 
//      (http://msdn.microsoft.com/en-us/library/ms686736.aspx)
//
//   RETURN VALUE: The return value indicates the success or failure of the 
//      function. 
//
DWORD WINAPI AutomateOutlookByCOMAPI(LPVOID lpParam);

DWORD WINAPI AutomateOutlookByCOMAPI2(LPVOID lpParam);