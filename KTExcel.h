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

int KTAPI KTLoadTemplateExcelFile(const TCHAR* filename);
int  KTAPI KTGetSheetIndex(int handle);
void KTAPI KTSetSheetIndex(int handle, int sheet);
void KTAPI KTSetCellValue(int handle, int row, int col, const char* type, const TCHAR* data);
BOOL KTAPI KTGetCellValue(int handle, int row, int col, const char* type, TCHAR* data, int dlc);
BOOL KTAPI KTSaveExcelFile(int handle, const TCHAR* filename);
void KTAPI KTCloseTemplateExcelFile(int handle);
BOOL KTAPI KTExcelStatus();


#ifdef __cplusplus
}
#endif


#endif