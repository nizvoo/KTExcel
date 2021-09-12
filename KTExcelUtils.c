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

#include <windows.h>
#include <tchar.h>

#include "KTExcel.h"


static HINSTANCE inst;

typedef int (KTAPI* TKTLoadTemplateExcelFile)(const TCHAR* filename);
typedef void (KTAPI* TKTSetCellValue)(int handle, int row, int col, const char* type, const TCHAR* data);
typedef BOOL(KTAPI* TKTGetCellValue)(int handle, int row, int col, const char* type, const TCHAR* data, int dlc);
typedef void (KTAPI* TKTSetCellFloatValue)(int handle, int row, int col, float data);
typedef BOOL(KTAPI* TKTSaveExcelFile)(int handle, const TCHAR* filename);
typedef void (KTAPI* TKTCloseTemplateExcelFile)(int handle);
typedef BOOL(KTAPI* TKTExcelStatus)();
typedef void (KTAPI* TKTSetSheetIndex)(int handle, int sheet);
typedef int (KTAPI*  TKTGetSheetIndex)(int handle);
typedef BOOL (KTAPI* TKTPasteCellUserStringProc)(int handle, const char* text);


static TKTLoadTemplateExcelFile KTLoadTemplateExcelFileProc = NULL;
static TKTSetCellValue KTSetCellValueProc = NULL;
static TKTGetCellValue KTGetCellValueProc = NULL;
static TKTSetCellFloatValue KTSetCellValueFloatProc = NULL;
static TKTSaveExcelFile KTSaveExcelFileProc = NULL;
static TKTCloseTemplateExcelFile KTCloseTemplateExcelFileProc = NULL;
static TKTExcelStatus KTExcelStatusProc = NULL;
static TKTSetSheetIndex KTSetSheetIndexProc = NULL;
static TKTGetSheetIndex KTGetSheetIndexProc = NULL;
static TKTPasteCellUserStringProc KTPasteCellUserStringProc = NULL;


BOOL KTAPI KTInitExcel(const TCHAR* path)
{
  TCHAR filename[MAX_PATH + 1] = { 0 };
  DWORD error_id;
  _sntprintf(filename, MAX_PATH, _T("%s/ESExcel.dll"), path);

  inst = LoadLibrary(filename);

  error_id = GetLastError();
  if (!inst) return FALSE;

  KTLoadTemplateExcelFileProc = (TKTLoadTemplateExcelFile)GetProcAddress(inst, "KTLoadTemplateExcelFile");
  KTSetCellValueProc = (TKTSetCellValue)GetProcAddress(inst, "KTSetCellValue");
  KTSetCellValueFloatProc = (TKTSetCellFloatValue)GetProcAddress(inst, "KTSetCellFloatValue"); 
  KTGetCellValueProc = (TKTGetCellValue)GetProcAddress(inst, "KTGetCellValue");
  KTSaveExcelFileProc = (TKTSaveExcelFile)GetProcAddress(inst, "KTSaveExcelFile");
  KTCloseTemplateExcelFileProc = (TKTCloseTemplateExcelFile)GetProcAddress(inst, "KTCloseTemplateExcelFile");
  KTExcelStatusProc = (TKTExcelStatus)GetProcAddress(inst, "KTExcelStatus");
  KTGetSheetIndexProc = (TKTGetSheetIndex)GetProcAddress(inst, "KTGetSheetIndex");
  KTSetSheetIndexProc = (TKTSetSheetIndex)GetProcAddress(inst, "KTSetSheetIndex");
  KTPasteCellUserStringProc = (TKTPasteCellUserStringProc)GetProcAddress(inst, "KTPasteCellUserString");

  if (!KTPasteCellUserStringProc) {
    printf("Binding KTPasteCellUserStringProc function not OK.\n");
    FreeLibrary(inst);
    inst = NULL;
    return FALSE;
  }

  if (!KTSetCellValueFloatProc) {
    printf("Binding KTSetCellValueFloatProc function not OK.\n");
    FreeLibrary(inst);
    inst = NULL;
    return FALSE;
  }
  if (!KTExcelStatusProc) {
    FreeLibrary(inst);
    inst = NULL;
    printf("Load Excel Status Entry Point fail.\n");
  }
  return TRUE;
}

void KTAPI KTUnInitExcel()
{
  if (inst)
    FreeLibrary(inst);
}

BOOL KTAPI KTExcelStatus()
{
  BOOL res = FALSE;

  if (!inst) return FALSE;

  if (KTExcelStatusProc) {
    res = KTExcelStatusProc();
  }

  return res;
}

int KTAPI KTLoadTemplateExcelFile(const TCHAR* filename)
{
  //if (!KTLoadTemplateExcelFileProc) return FALSE;

  return KTLoadTemplateExcelFileProc(filename);
}

void KTAPI KTSetCellValue(int handle, int row, int col, const char* type, const TCHAR* data)
{
  //if (!KTSetCellValueProc) return;

  KTSetCellValueProc(handle, row, col, type, data);
}

void KTAPI KTSetCellFloatValue(int handle, int row, int col, float data)
{
  if (!KTSetCellValueFloatProc) return;

  KTSetCellValueFloatProc(handle, row, col, data);
}

BOOL KTAPI KTSaveExcelFile(int handle, const TCHAR* filename)
{
  if (!KTSaveExcelFileProc) return FALSE;

  return KTSaveExcelFileProc(handle, filename);
}

void KTAPI KTCloseTemplateExcelFile(int handle)
{
  //if (!KTCloseTemplateExcelFileProc) return;

  KTCloseTemplateExcelFileProc(handle);
}

BOOL KTAPI KTGetCellValue(int handle, int row, int col, const char* type, TCHAR* data, int dlc)
{
  //if (!KTGetCellValueProc) return FALSE;

  return KTGetCellValueProc(handle, row, col, type, data, dlc);
}

int KTAPI KTGetSheetIndex(int handle)
{
  //if (!KTGetSheetIndexProc) return 0;

  return KTGetSheetIndexProc(handle);
}

void KTAPI KTSetSheetIndex(int handle, int sheet)
{
  if (KTSetSheetIndexProc)
    KTSetSheetIndexProc(handle, sheet);
}

BOOL KTAPI KTPasteCellUserString(int handle, const char* text)
{
  if (!KTPasteCellUserStringProc) return FALSE;
  return KTPasteCellUserStringProc(handle, text);
}