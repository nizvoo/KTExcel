#include <stdio.h>
#include <stdlib.h>

#include <windows.h>
#include <tchar.h>

#include <atlbase.h>
#include <atlstr.h>

#include <shellapi.h>

#include <atlwin.h>

#include <atlcoll.h>  // ATL collections


#include "KTExcel.h"

#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE14\\MSO.DLL" \
	rename( "RGB", "MSORGB" )

using namespace Office;

#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"

using namespace VBIDE;

#import "C:\\Program Files (x86)\\Microsoft Office\\OFFICE14\\EXCEL.EXE" \
	rename( "DialogBox", "ExcelDialogBox" ) \
	rename( "RGB", "ExcelRGB" ) \
	rename( "CopyFile", "ExcelCopyFile" ) \
	rename( "ReplaceText", "ExcelReplaceText" )

static void ErrorfV1(const TCHAR* text)
{
  _ftprintf(stdout, _T("%s\n"), text);
}

static void Errorf( LPCTSTR pszFormat, ... )
{
	CString		str;

	va_list	argList;

	va_start( argList, pszFormat );

	str.FormatV( pszFormat, argList );

	::MessageBox(NULL, (LPCTSTR) str, _T("WTLExcel Error"), MB_ICONHAND | MB_OK );

	return;
}

struct user_excel_st
{
  Excel::_ApplicationPtr app;
  Excel::_WorkbookPtr book;
  Excel::_WorksheetPtr sheet;
  Excel::RangePtr range;
  
  user_excel_st()
  {
  }
};

struct user_excel_st* user_excel = NULL;

extern "C" BOOL KTAPI KTLoadTemplateExcelFile(const TCHAR* filename)
{
  // Load the Excel application in the background.
  user_excel = new user_excel_st;

  if ( FAILED(user_excel->app.CreateInstance( _T("Excel.Application")))) {
    Errorf( _T("Failed to initialize Excel::_Application!") );
    return FALSE;
  }

  _variant_t	varOption( (long)DISP_E_PARAMNOTFOUND, VT_ERROR );

  if (!user_excel->app || !user_excel->app->Workbooks) {
    Errorf(_T("Workbooks is empty\n")); 
    return FALSE;
  }

  user_excel->book = user_excel->app->Workbooks->Open(filename);//, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption );
  if (user_excel->book == NULL) {
    Errorf( _T("Failed to open Excel file!") );
    return FALSE;
  }
  user_excel->app->PutVisible(0, FALSE); 

  user_excel->sheet = user_excel->book->Sheets->Item[1];

  user_excel->range = user_excel->sheet->Cells;     

  return TRUE;
}

extern "C" void KTAPI KTSetCellValue(int row, int col, const char* type, const TCHAR* data)
{
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
extern "C" BOOL KTAPI KTGetCellValue(int row, int col, const char* type, TCHAR* data, int dlc)
{
  BOOL res = FALSE;
  if (stricmp(type, "float") == 0) {
    float v = user_excel->range->Item[row][col];
    _sntprintf(data, dlc, _T("%0.3f"), v);    
    res = TRUE;
  } else if (stricmp(type, "int") == 0) {
    float v = user_excel->range->Item[row][col];
    _sntprintf(data, dlc, _T("%d"), v); 
    res = TRUE;
  } else if (stricmp(type, "string") == 0) {
		_variant_t item = user_excel->range->Item[row][col];
		_bstr_t bstrText(item);

    _sntprintf(data, dlc, _T("%s"), bstrText.GetBSTR()) ; 
    res = TRUE;
  }
  return res;
}

extern "C" BOOL KTAPI KTSaveExcelFile(const TCHAR* filename)
{
  // Switch off alert prompting to save as   
  user_excel->app->PutDisplayAlerts( LOCALE_USER_DEFAULT, VARIANT_FALSE );  

  // Save the values in book.xml  
  user_excel->sheet->SaveAs(filename);  
  user_excel->app->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_TRUE );
  // Don't save any inadvertant changes to the .xls file.
	return TRUE;
}

extern "C" void KTAPI KTCloseTemplateExcelFile()
{
  if (user_excel) {
    user_excel->book->Close( VARIANT_FALSE );
    // And switch back on again...  

    // Need to quit, otherwise Excel remains active and locks the .xls file.
    user_excel->app->Quit( );

    delete user_excel;
    user_excel = NULL;
  }
}

extern "C" BOOL LoadExcelFile(const TCHAR* filename)
{

	// Load the Excel application in the background.
	Excel::_ApplicationPtr pApplication;

	if ( FAILED( pApplication.CreateInstance( _T("Excel.Application") ) ) )
	{
		Errorf( _T("Failed to initialize Excel::_Application!") );
		return FALSE;
	}

	_variant_t	varOption( (long) DISP_E_PARAMNOTFOUND, VT_ERROR );
  _ftprintf(stdout, _T("%s\n"), filename);
  if (!pApplication || !pApplication->Workbooks) {
    Errorf(_T("Workbooks is empty\n")); 
    return FALSE;
  }
  
	Excel::_WorkbookPtr pBook = pApplication->Workbooks->Open(filename);//, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption );
	if ( pBook == NULL )
	{
		Errorf( _T("Failed to open Excel file!") );
		return FALSE;
	}
  pApplication->PutVisible(0, FALSE); 

	Excel::_WorksheetPtr pSheet = pBook->Sheets->Item[1];

	if ( pSheet == NULL )
	{
		Errorf( _T("Failed to get first Worksheet!") );
		return FALSE;
	}
  
  Excel::RangePtr pRange = pSheet->Cells;     
  
  for (int i = 24; i < 29; ++i) {
    for (int j = 3; j < 27; ++j) {
      pRange->Item[i][j] = j / 10 + i / 100.f;
    }
  }
#if 0
  fprintf(stdout, "2\n");
	// Load the column headers.
	Excel::RangePtr pRange = pSheet->GetRange( _bstr_t( _T("A1") ), _bstr_t( _T("Z1" ) ) );

	if ( pRange == NULL )
	{
		Errorf( _T("Failed to get header cell range( A1:Z1 )!") );
		return FALSE;
	}

	int			iColumns = 0;

	for ( int iColumn = 1; iColumn < 26; ++iColumn )
	{
		_variant_t	vItem = pRange->Item[ 1 ][ iColumn ];
		_bstr_t		bstrText( vItem );

		if ( bstrText.length( ) == 0 )
			break;

		//m_list.AddColumn( bstrText, iColumns++ );
    //_ftprintf(stdout, _T("%s\n"), bstrText);
    Errorf(bstrText);
	}

	// Load the rows (up to the first blank one).
	pRange = pSheet->GetRange( _bstr_t( _T("A2") ), _bstr_t( _T("Z16384" ) ) );

	for ( int iRow = 1; ; ++iRow )
	{
    int iColumn;
		for (  iColumn = 1; iColumn <= iColumns; ++iColumn )
		{
			_variant_t	vItem = pRange->Item[ iRow ][ iColumn ];
			_bstr_t		bstrText( vItem );

			if ( bstrText.length( ) == 0 )
				break;

			if ( iColumn == 1 )
				;//m_list.AddItem( iRow - 1, 0, bstrText );
			else
				;//m_list.SetItemText( iRow - 1, iColumn - 1, bstrText );
		}

		if ( iColumn == 1 )
			break;
	}
  
  pRange->Item[ 1 ][ 1 ] = 1234;
#endif
	// Make it all look pretty.
	//for ( int iColumn = 1; iColumn <= iColumns; ++iColumn )
		//m_list.SetColumnWidth( iColumn, LVSCW_AUTOSIZE_USEHEADER );


  
  
    // Switch off alert prompting to save as   
    pApplication->PutDisplayAlerts( LOCALE_USER_DEFAULT, VARIANT_FALSE );  
  
    // Save the values in book.xml  
    pSheet->SaveAs(_T("d:\\tesstbook.xls"));  
        pApplication->PutDisplayAlerts( LOCALE_USER_DEFAULT, VARIANT_TRUE );
  	// Don't save any inadvertant changes to the .xls file.
	pBook->Close( VARIANT_FALSE );
    // And switch back on again...  
   

	// Need to quit, otherwise Excel remains active and locks the .xls file.
	pApplication->Quit( );

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

	if ( FAILED(app.CreateInstance( _T("Excel.Application")))) {
    return FALSE;
  }

	app->Quit( );
  return TRUE;
}