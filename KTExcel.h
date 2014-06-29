#ifndef __KT_EXCEL_H__
#define __KT_EXCEL_H__

#ifdef __cplusplus
extern "C" {
#endif

#ifndef KTAPI
#define KTAPI __stdcall
#endif

BOOL KTAPI KTInitExcel(const TCHAR*);
void KTAPI KTUnInitExcel();

BOOL KTAPI KTLoadTemplateExcelFile(const TCHAR* filename);
void KTAPI KTSetCellValue(int row, int col, const char* type, const TCHAR* data);
BOOL KTAPI KTGetCellValue(int row, int col, const char* type, TCHAR* data, int dlc);
BOOL KTAPI KTSaveExcelFile(const TCHAR* filename);
void KTAPI KTCloseTemplateExcelFile();
BOOL KTAPI KTExcelStatus();


#ifdef __cplusplus
}
#endif


#endif