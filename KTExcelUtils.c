#include <windows.h>
#include <tchar.h>

#include "KTExcel.h"


static HINSTANCE inst;

typedef BOOL (KTAPI* TKTLoadTemplateExcelFile)(const TCHAR* filename);
typedef void (KTAPI* TKTSetCellValue)(int row, int col, const char* type, const TCHAR* data);
typedef BOOL (KTAPI* TKTGetCellValue)(int row, int col, const char* type, const TCHAR* data, int dlc);
typedef BOOL (KTAPI* TKTSaveExcelFile)(const TCHAR* filename);
typedef void (KTAPI* TKTCloseTemplateExcelFile)();

static TKTLoadTemplateExcelFile KTLoadTemplateExcelFileProc = NULL;
static TKTSetCellValue KTSetCellValueProc = NULL;
static TKTGetCellValue KTGetCellValueProc = NULL;
static TKTSaveExcelFile KTSaveExcelFileProc = NULL;
static TKTCloseTemplateExcelFile KTCloseTemplateExcelFileProc = NULL;


BOOL KTAPI KTInitExcel(const TCHAR* path)
{
  TCHAR filename[MAX_PATH + 1] = {0};
  DWORD error_id;
  _sntprintf(filename, MAX_PATH, _T("%s/KTExcel.dll"), path);
  
  inst  = LoadLibrary(filename);

  error_id = GetLastError();
  if (!inst) return FALSE;

  KTLoadTemplateExcelFileProc = (TKTLoadTemplateExcelFile)GetProcAddress(inst, "KTLoadTemplateExcelFile");
  KTSetCellValueProc = (TKTSetCellValue)GetProcAddress(inst, "KTSetCellValue");
  KTGetCellValueProc = (TKTGetCellValue)GetProcAddress(inst, "KTGetCellValue");
  KTSaveExcelFileProc = (TKTSaveExcelFile)GetProcAddress(inst, "KTSaveExcelFile");
  KTCloseTemplateExcelFileProc = (TKTCloseTemplateExcelFile)GetProcAddress(inst, "KTCloseTemplateExcelFile");
  return TRUE;
}

void KTAPI KTUnInitExcel()
{
  if (inst)
    FreeLibrary(inst);
}

BOOL KTAPI KTExcelStatus()
{
  if (!inst) return FALSE;

  return TRUE;
}

BOOL KTAPI KTLoadTemplateExcelFile(const TCHAR* filename)
{
  if (!KTLoadTemplateExcelFileProc) return FALSE;
  
  return KTLoadTemplateExcelFileProc(filename);
}

void KTAPI KTSetCellValue(int row, int col, const char* type, const TCHAR* data)
{
  if (!KTSetCellValueProc) return;
  
  KTSetCellValueProc(row, col, type, data);
}

BOOL KTAPI KTSaveExcelFile(const TCHAR* filename)
{
  if (!KTSaveExcelFileProc) return FALSE;
  
  return KTSaveExcelFileProc(filename);
}

void KTAPI KTCloseTemplateExcelFile()
{
  if (!KTCloseTemplateExcelFileProc) return;
  
  KTCloseTemplateExcelFileProc();
}


BOOL KTAPI KTGetCellValue(int row, int col, const char* type, TCHAR* data, int dlc)
{
  if (!KTGetCellValueProc) return FALSE;
  
  return KTGetCellValueProc(row, col, type, data, dlc);
}