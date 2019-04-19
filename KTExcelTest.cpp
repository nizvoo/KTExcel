#include <stdio.h>
#include <stdlib.h>

#include <windows.h>
#include <tchar.h>

#include "KTExcel.h"

#include <assert.h>


static BOOL KTGetUserAppPath(TCHAR* buf, DWORD maxlen)
{
  DWORD ret = GetModuleFileName(NULL, buf, MAX_PATH);
  if (ret > maxlen) return FALSE;
  if (ret != _tcslen(buf)) return FALSE;

  for (size_t i = _tcslen(buf) - 1; i >= 0; --i) {
    if (buf[i] == _T('/') || buf[i] == _T('\\')) {
      buf[i] = _T('\0');
      break;
    }
  }
  return TRUE;
}

int main(int argc, char* argv[])
{
  TCHAR template_filename[MAX_PATH + 1] = {0};
  TCHAR filename[MAX_PATH + 1] = {0};
  TCHAR path[MAX_PATH + 1] = {0};

  if (!KTGetUserAppPath(path, MAX_PATH)) {
    return -1;
  }
  _ftprintf(stdout, _T("%s\n"), path);
  
  KTInitExcel(path);

	HRESULT hRes = ::CoInitialize(NULL);

	assert(SUCCEEDED(hRes));

  _sntprintf(template_filename, MAX_PATH, _T("%s/UserTemplate.xls"), path);
  _sntprintf(filename, MAX_PATH, _T("%s/MyUserTemplate.xls"), path);

  _ftprintf(stdout, _T("%s\n%s\n"), template_filename, filename);

  if (KTExcelStatus()) {

    KTLoadTemplateExcelFile(template_filename);

    TCHAR text[MAX_PATH + 1] = {0};
    for (int i = 1; i < 4; ++i) {
      for (int j = 1; j < 4; ++j) {
        KTGetCellValue(i, j, "string", text, MAX_PATH);
        _ftprintf(stdout, _T("%s\t"), text);
      }
      _ftprintf(stdout, _T("\n"));
    }
    _ftprintf(stdout, _T("\n"));

    KTSetSheetIndex(0);
    KTSetCellValue(5, 1, "int", _T("1"));
    KTSetCellValue(5, 2, "int", _T("2"));
    KTSetCellValue(5, 3, "int", _T("3"));

    KTSetCellValue(6, 1, "float", _T("0.1"));
    KTSetCellValue(6, 2, "float", _T("0.2"));
    KTSetCellValue(6, 3, "float", _T("0.3"));

    KTSetSheetIndex(1);

    KTSetCellValue(5, 1, "int", _T("11"));
    KTSetCellValue(5, 2, "int", _T("12"));
    KTSetCellValue(5, 3, "int", _T("13"));

    KTSetCellValue(6, 1, "float", _T("1.1"));
    KTSetCellValue(6, 2, "float", _T("1.2"));
    KTSetCellValue(6, 3, "float", _T("1.3"));


    KTSaveExcelFile(filename);

    KTCloseTemplateExcelFile();
  } else {
    fprintf(stdout, "Invoking the EXCEL COM object fail.\n");
  }

  KTUnInitExcel();
  
  ::CoUninitialize();
  return 0;
}