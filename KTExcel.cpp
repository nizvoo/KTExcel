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

#include <atlbase.h>
#include <atlstr.h>

#include <shellapi.h>

#include <atlwin.h>

#include <atlcoll.h>

#include "KTExcel.h"

#import "MSO.DLL" rename( "RGB", "MSORGB" )

using namespace Office;

#import "VBE6EXT.OLB"

#import "EXCEL.EXE" \
  rename( "DialogBox", "ExcelDialogBox" ) \
  rename( "RGB", "ExcelRGB" ) \
  rename( "CopyFile", "ExcelCopyFile" ) \
  rename( "ReplaceText", "ExcelReplaceText" )


static void ErrorfV1(const TCHAR* text)
{
  _ftprintf(stdout, _T("%s\n"), text);
}

static void Errorf(LPCTSTR pszFormat, ...)
{
  CString		str;

  va_list	argList;

  va_start(argList, pszFormat);

  str.FormatV(pszFormat, argList);

  ::MessageBox(NULL, (LPCTSTR)str, _T("WTLExcel Error"), MB_ICONHAND | MB_OK);

  return;
}

struct user_excel_st
{
  Excel::_ApplicationPtr app;
  Excel::_WorkbookPtr book;
  Excel::_WorksheetPtr sheet;
  Excel::RangePtr range;
  int sheet_index;
  user_excel_st()
  {
    sheet_index = 1;
  }
};

struct user_excel_st* user_excel_list[MAX_PATH] = {0};
int                   user_excel_count = 0;

extern "C" int KTAPI KTLoadTemplateExcelFile(const TCHAR* filename)
{
  // Load the Excel application in the background.
  user_excel_st* user_excel = new user_excel_st;

  if (FAILED(user_excel->app.CreateInstance(_T("Excel.Application")))) {
    Errorf(_T("Failed to initialize Excel::_Application!"));
    return FALSE;
  }

  _variant_t	varOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

  if (!user_excel->app || !user_excel->app->Workbooks) {
    Errorf(_T("Workbooks is empty\n"));
    return FALSE;
  }

  user_excel->book = user_excel->app->Workbooks->Open(filename);
  if (user_excel->book == NULL) {
    Errorf(_T("Failed to open Excel file!"));
    return FALSE;
  }
  user_excel->app->PutVisible(0, FALSE);

  user_excel->sheet = user_excel->book->Sheets->Item[1];

  user_excel->range = user_excel->sheet->Cells;

  int curr_pos = user_excel_count++;

  if (curr_pos > MAX_PATH - 1) {
    curr_pos = 0;
  }

  user_excel_list[curr_pos] = user_excel;

  return curr_pos;
}

void KTAPI KTSetSheetIndex(int handle, int sheet)
{
  user_excel_st* user_excel = user_excel_list[handle];
  user_excel->sheet_index = sheet + 1;
  user_excel->sheet = user_excel->book->Sheets->Item[sheet + 1];
  user_excel->range = user_excel->sheet->Cells;
}

int KTAPI KTGetSheetIndex(int handle)
{
  user_excel_st* user_excel = user_excel_list[handle];
  return user_excel->sheet_index - 1;
}

extern "C" void KTAPI KTSetCellValue(int handle, int row, int col, const char* type, const TCHAR* data)
{
  user_excel_st* user_excel = user_excel_list[handle];

  if (stricmp(type, "float") == 0) {
    float v = _ttof(data);
    user_excel->range->Item[row][col] = v;
  } else if (stricmp(type, "int") == 0) {
    int v = _ttof(data);
    user_excel->range->Item[row][col] = v;
  } else if (stricmp(type, "string") == 0) {
    user_excel->range->Item[row][col] = data;
  }
}

/* http://msdn.microsoft.com/en-us/library/x295h94e.aspx */
extern "C" BOOL KTAPI KTGetCellValue(int handle, int row, int col, const char* type, TCHAR* data, int dlc)
{
  user_excel_st* user_excel = user_excel_list[handle];

  BOOL res = FALSE;
  if (stricmp(type, "float") == 0) {
    float v = user_excel->range->Item[row][col];
    _sntprintf(data, dlc, _T("%0.3f"), v);
    res = TRUE;
  } else if (stricmp(type, "int") == 0) {
    float v = user_excel->range->Item[row][col];
    _sntprintf(data, dlc, _T("%d"), (int)v);
    res = TRUE;
  } else if (stricmp(type, "string") == 0) {
    _variant_t item = user_excel->range->Item[row][col];
    _bstr_t bstrText(item);

    _sntprintf(data, dlc, _T("%s"), bstrText.GetBSTR());
    res = TRUE;
  }
  return res;
}

extern "C" BOOL KTAPI KTSaveExcelFile(int handle, const TCHAR* filename)
{
  user_excel_st* user_excel = user_excel_list[handle];

  user_excel->app->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_FALSE);

  user_excel->sheet->SaveAs(filename);
  user_excel->app->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_TRUE);

  return TRUE;
}

extern "C" void KTAPI KTCloseTemplateExcelFile(int handle)
{
  user_excel_st* user_excel = user_excel_list[handle];

  if (user_excel) {
    user_excel->book->Close(VARIANT_FALSE);

    user_excel->app->Quit();

    delete user_excel;
    user_excel = NULL;
  }
}

extern "C" BOOL LoadExcelFile(const TCHAR* filename)
{
  Excel::_ApplicationPtr pApplication;

  if (FAILED(pApplication.CreateInstance(_T("Excel.Application"))))
  {
    Errorf(_T("Failed to initialize Excel::_Application!"));
    return FALSE;
  }

  _variant_t	varOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
  _ftprintf(stdout, _T("%s\n"), filename);
  if (!pApplication || !pApplication->Workbooks) {
    Errorf(_T("Workbooks is empty\n"));
    return FALSE;
  }

  Excel::_WorkbookPtr pBook = pApplication->Workbooks->Open(filename);
  if (pBook == NULL)
  {
    Errorf(_T("Failed to open Excel file!"));
    return FALSE;
  }
  pApplication->PutVisible(0, FALSE);

  Excel::_WorksheetPtr pSheet = pBook->Sheets->Item[1];

  if (pSheet == NULL)
  {
    Errorf(_T("Failed to get first Worksheet!"));
    return FALSE;
  }

  Excel::RangePtr pRange = pSheet->Cells;

  for (int i = 24; i < 29; ++i) {
    for (int j = 3; j < 27; ++j) {
      pRange->Item[i][j] = j / 10 + i / 100.f;
    }
  }
 
  pApplication->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_FALSE);

  pSheet->SaveAs(_T("./TestExcel.xls"));
  pApplication->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_TRUE);

  pBook->Close(VARIANT_FALSE);

  pApplication->Quit();

  return TRUE;
}

extern "C" BOOL KTAPI KTInitExcel(const TCHAR* text)
{
  return FALSE;
}

extern "C"  void KTAPI KTUnInitExcel()
{
}

extern "C" BOOL KTAPI KTExcelStatus()
{
  Excel::_ApplicationPtr app;

  if (FAILED(app.CreateInstance(_T("Excel.Application")))) {
    return FALSE;
  }

  app->Quit();
  return TRUE;
}