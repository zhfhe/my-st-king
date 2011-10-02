// MyExcelReadDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "..\stkui.h"
#include "MyExcelReadDlg.h"
#include "afxdialogex.h"
#include "memory.h"
#include <time.h>
#include <atlbase.h>
#include "strptime.h"

// CMyExcelReadDlg 对话框

IMPLEMENT_DYNAMIC(CMyExcelReadDlg, CDialogEx)

CMyExcelReadDlg::CMyExcelReadDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMyExcelReadDlg::IDD, pParent)
{
	m_excelFileName = _T("");
}

CMyExcelReadDlg::~CMyExcelReadDlg()
{
}

void CMyExcelReadDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CMyExcelReadDlg, CDialogEx)
	ON_BN_CLICKED(IDC_EXCEL_EXPLORER, OnExcelExplorer)
END_MESSAGE_MAP()


// CMyExcelReadDlg 消息处理程序

void CMyExcelReadDlg::OnExcelExplorer() 
{
	// TODO: Add your control notification handler code here
	UpdateData( TRUE );

	CKSFileDialog dlg (TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_FILEMUSTEXIST | OFN_EXPLORER | OFN_ENABLESIZING ,
		"Excel files (*.xls)|*.xls|All files (*.*)|*.*||", NULL);

	int nResult = dlg.DoModal();
	if( IDOK == nResult )
	{
		m_excelFileName = dlg.GetPathName();

	}

	
}


KDATA * CMyExcelReadDlg::GetPKDATA() 
{
	PKDATA pKData = NULL;
    int i;
	CMyExcel ExcelMain;
	CSPTime	sptime;
	struct tm* pTemptime = NULL;
	
	//very important to set struct tm to 0.
	pTemptime = (struct tm*)malloc(sizeof(struct tm));	
	memset( pTemptime, 0, sizeof(struct tm) );
	
	ExcelMain.Open(m_excelFileName);
	LPDISPATCH  lpDisp = NULL;
	//get 1st sheet
	lpDisp = ExcelMain.MySheets.GetItem((_variant_t)(long)1);
	ExcelMain.MySheet.AttachDispatch(lpDisp,TRUE);
	CString sheetName = ExcelMain.MySheet.GetName();
	ExcelMain.OpenSheet( sheetName  ); 
	ExcelMain.AutoRange();

    //calculate the data row total number 
	i = TITLE_ROW_COUNT + 1;
	while( !ExcelMain.GetItemText( i, 1 ).IsEmpty( ) )
	{
		i++;
	}		 
	m_dataSize = i - TITLE_ROW_COUNT - 1;
		
				
	//init data Array
	//need add code for malloc failure
	pKData = (KDATA *)malloc(m_dataSize * sizeof(KDATA));
	memset( pKData, 0, m_dataSize * sizeof(KDATA) );	

	//set KData using Excel data	 
    for(i = 0;i < m_dataSize;i++)
    {
  		strptime(CT2CA(ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, DATE_INDEX )), "%Y-%m-%d", pTemptime);
		//strptime((LPCTSTR)ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, DATE_INDEX ), "%Y-%m-%d", pTemptime);
		sptime = mktime(pTemptime);
		pKData[i].m_date = sptime.ToStockTimeDay();
			
		pKData[i].m_fOpen = atof( ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, OPENPRICE_INDEX ));
		pKData[i].m_fHigh = atof( ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, TOPPRICE_INDEX ));
		pKData[i].m_fLow = atof( ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, LOWPRICE_INDEX ));
		pKData[i].m_fClose = atof( ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, CLOSEPRICE_INDEX ));
		pKData[i].m_fVolume = atof( ExcelMain.GetItemText( i + TITLE_ROW_COUNT + 1, VOLUME_INDEX ));
	}


	ExcelMain.Exit();
	free(pTemptime);

	return pKData;
	
}

int CMyExcelReadDlg::GetDataSize() 
{
   return m_dataSize;
}

