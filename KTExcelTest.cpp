#include <stdio.h>
#include <stdlib.h>

#include <windows.h>
#include <tchar.h>

#include "KTExcel.h"

#include <assert.h>

int main(int argc, char* argv[])
{
	HRESULT hRes = ::CoInitialize(NULL);
// If you are running on NT 4.0 or higher you can use the following call instead to 
// make the EXE free threaded. This means that calls come in on a random RPC thread.
//	HRESULT hRes = ::CoInitializeEx(NULL, COINIT_MULTITHREADED);
	assert(SUCCEEDED(hRes));
  
  KTLoadTemplateExcelFile(_T("D:\\test\\WTLExcelSource\\KTExcel\\Report.xls"));
  for (int i = 24; i < 29; ++i) {
    for (int j = 3; j < 27; ++j) {
      KTSetCellValue(i, j, j / 10 + i / 100.f);
    }
  }  
  KTSaveExcelFile(_T("d:\\test.xls"));
  
  KTCloseTemplateExcelFile();
  ::CoUninitialize();
  return 0;
}