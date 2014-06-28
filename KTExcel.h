#ifndef __KT_EXCEL_H__
#define __KT_EXCEL_H__

#ifdef __cplusplus
extern "C" {
#endif


BOOL KTLoadTemplateExcelFile(const TCHAR* filename);
void KTSetCellValue(int row, int col, const char* type, const TCHAR* data);
BOOL KTSaveExcelFile(const TCHAR* filename);
void KTCloseTemplateExcelFile();



#ifdef __cplusplus
}
#endif


#endif