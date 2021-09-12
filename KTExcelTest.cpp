/*
 * (C) Copyright AOE Studio 2010 - All Rights Reserved.
 *
 * This software is the confidential and proprietary information
 * of AOE Studio  ("Confidential Information").  You
 * shall not disclose such Confidential Information and shall use
 * it only in accordance with the terms of the license agreement
 * you entered into with AOE Studio
 *
 */

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
  DWORD start_tm;

  start_tm = GetTickCount();

  if (!KTGetUserAppPath(path, MAX_PATH)) {
    return -1;
  }
  _ftprintf(stdout, _T("%s\n"), path);

  KTInitExcel(path);

  HRESULT hRes = ::CoInitializeEx(NULL, COINIT_MULTITHREADED);

  assert(SUCCEEDED(hRes));

  _sntprintf(template_filename, MAX_PATH, _T("%s/UserTemplate.xls"), path);
  _sntprintf(filename, MAX_PATH, _T("%s/MyUserTemplate.xls"), path);

  _ftprintf(stdout, _T("%s\n%s\n"), template_filename, filename);

  if (!KTExcelStatus()) return -1;


  int handle = KTLoadTemplateExcelFile(template_filename);
  if (handle >= 0) {

    TCHAR text[MAX_PATH + 1] = {0};
    for (int i = 1; i < 4; ++i) {
      for (int j = 1; j < 4; ++j) {
        KTGetCellValue(handle, i, j, "string", text, MAX_PATH);
        _ftprintf(stdout, _T("%s\t"), text);
      }
      _ftprintf(stdout, _T("\n"));
    }
    _ftprintf(stdout, _T("\n"));

    KTSetSheetIndex(handle, 0);
    KTSetCellValue(handle, 5, 1, "int", _T("1"));
    KTSetCellValue(handle, 5, 2, "int", _T("2"));
    KTSetCellValue(handle, 5, 3, "int", _T("3"));

    KTSetCellValue(handle, 6, 1, "float", _T("0.1"));
    KTSetCellValue(handle, 6, 2, "float", _T("0.2"));
    KTSetCellValue(handle, 6, 3, "float", _T("0.3"));

    KTSetSheetIndex(handle, 1);

    DWORD start_batch_tm = GetTickCount();
    DWORD slice_tm = GetTickCount();
    char str[MAX_PATH + 1] = {0};
    int cols = 20;
    int total = 1000 * 1000;
    int sheet_text_size = total * 10 * cols;
    char* sheet_text = new char[sheet_text_size];
    memset(sheet_text, 0x00, sizeof(char) * sheet_text_size);
    int str_len = 0;
    for (int i = 0; i < total; ++i) {
      for (int j = 0; j < cols; ++j) {
        _snprintf(str, MAX_PATH, "%0.2f", i * 0.1f + j * 0.01f);
        strncpy(sheet_text + str_len, str, strlen(str));
        str_len += strlen(str);
        if (j < total - 1) {
          strncpy(sheet_text + str_len, "\t", strlen("\t"));
          str_len += strlen("\t");
        }
      }
      strncpy(sheet_text + str_len, "\n", strlen("\n"));
      str_len += strlen("\n");

      if (i % 1000 == 0) {
        DWORD use_tm = GetTickCount() - slice_tm;
        printf("%d\t%0.2f s\t%0.2f s\t%0.2f%%\n", i, use_tm / 1000.f, (GetTickCount() - start_batch_tm) * 0.001f, 100 * (1 + i) / (float)total);
        slice_tm = GetTickCount();
      }
    }
    printf("%s\n", sheet_text);
    KTPasteCellUserString(handle, sheet_text);
    delete [] sheet_text;
    sheet_text = NULL;

    printf("Preparing other cell\n");

    DWORD use_batch_tm = GetTickCount() - start_batch_tm;
    KTSetSheetIndex(handle, 0);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("OverSizeX"));
    KTSetCellValue(handle, 5, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("LowSizeX"));
    KTSetCellValue(handle, 6, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("OverSizeY"));
    KTSetCellValue(handle, 7, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("LowSizeY"));
    KTSetCellValue(handle, 8, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("OverArea"));
    KTSetCellValue(handle, 9, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("LowArea"));
    KTSetCellValue(handle, 10, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("ShiftX"));
    KTSetCellValue(handle, 11, 6, "string", text);

    _sntprintf(text, MAX_PATH, _T("%s"), _T("ShiftY"));
    KTSetCellValue(handle, 12, 6, "string", text);


    KTSaveExcelFile(handle, filename);

    KTCloseTemplateExcelFile(handle);
    printf("Use time:%0.3f S, use batch time:0.3f S\n", (GetTickCount() - start_tm) / 1000.0, use_batch_tm * 0.001);

  } else {
    fprintf(stdout, "Invoking the EXCEL COM object fail.\n");
  }

  KTUnInitExcel();

  ::CoUninitialize();
  return 0;
}