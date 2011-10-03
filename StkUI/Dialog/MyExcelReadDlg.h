#pragma once

#include "MyExcel.h"
#include "afxdialogex.h"
// CMyExcelReadDlg 对话框

#define TITLE_ROW_COUNT              1                //标题占几行
#define WEEK_COUNT                   5
enum EXCEL_INDEX
{
CONTRACT_INDEX=1,
DATE_INDEX,
OPENPRICE_INDEX,
TOPPRICE_INDEX,
LOWPRICE_INDEX,
CLOSEPRICE_INDEX,
VOLUME_INDEX
} ;

class CMyExcelReadDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CMyExcelReadDlg)

public:
	CMyExcelReadDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CMyExcelReadDlg();

// 对话框数据
	enum { IDD = IDD_READ_EXCEL };
	CButton	m_btnCancel;
	CButton	m_btnOK;
	CString	m_excelFileName;
	CButton m_btnExcelExplorer;
	CEdit m_editExcelName;
	int m_dataSize;

	int GetDataSize();
	KDATA * GetPKDATA();

	
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
//	virtual void OnOK();
	void OnExcelExplorer();

	DECLARE_MESSAGE_MAP()
};
