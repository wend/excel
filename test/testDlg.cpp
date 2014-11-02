
// testDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "test.h"
#include "testDlg.h"
#include "afxdialogex.h"
#include <string>
#include <fstream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CtestDlg �Ի���




CtestDlg::CtestDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CtestDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CtestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CtestDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CtestDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CtestDlg ��Ϣ�������

BOOL CtestDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	std::string sql = "asdf\n\r\na��sd\n";
	std::string::iterator it = sql.begin();
	while (it != sql.end())
	{
		if (*it == '\r' || *it == '\n')
		{
			//std::string::iterator tmpit = it;
			//it ++;
			it =sql.erase(it);

		}else {


			it++;
		}
	}

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CtestDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CtestDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CtestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

std::string trim(std::string& s)

{
	const std::string drop = " ";
	// trim right
	s.erase(s.find_last_not_of(drop)+1);
	// trim left
	return s.erase(0,s.find_first_not_of(drop));
}

void CtestDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();

	try
	{
		Excel::_ApplicationPtr pExcelApp;
		Excel::WorkbooksPtr pWorkbooks;
		Excel::_WorkbookPtr pWorkbook;
		LPDISPATCH lpDisp = NULL;


		HRESULT hr = pExcelApp.CreateInstance(L"Excel.Application");
		ATLASSERT(SUCCEEDED(hr));
		pExcelApp->Visible = false;   // make Excel��s main window visible
		CString version = pExcelApp->Version;
		pWorkbooks = pExcelApp->Workbooks;
		int value = 0;
		//for (int i = 0; i < 10; i++) {
		try {
			pWorkbook = pWorkbooks->Open("d:\\word.xls");  // open excel file
		} catch (...) {
			pWorkbook = pWorkbooks->Add();
		}


		Excel::_WorksheetPtr pWorksheet = pWorkbook->ActiveSheet;
		pWorksheet = pWorkbook->Sheets->Item[L"Sheet1"];

		Excel::RangePtr pRange = pWorksheet->Cells;

		//const int nplot = 100;
		//const double xlow = 0.0, xhigh = 20.0;
		//double h = (xhigh-xlow)/(double)nplot;
		//value += (int)(pRange->Item[1][1]);

		std::ofstream file("d:\\word.sql");

		_variant_t var_t;   
		_bstr_t bstr_t;
		int row = 0;
		while (true)
		{
			row++;
			var_t =pRange->Item[row][1];   
			bstr_t=var_t;  
			std::string wordId = (const char*)bstr_t;
			wordId = trim(wordId);
			if (wordId.find("List") != std::string::npos || wordId.find("list") != std::string::npos)
			{
				continue;
			}
			if (wordId == "")
				break;
			var_t =pRange->Item[row][2];   
			bstr_t=var_t;
			std::string word = (const char*)bstr_t;
			word = trim(word);
			var_t =pRange->Item[row][3];   
			bstr_t=var_t;
			std::string translation = (const char*)bstr_t;
			translation = trim(translation);
			std::string sql;
			if (translation.find("'")!= std::string::npos)
			{
				sql = "insert into WordSerial(WordId, Word, Translation) values(\"" + wordId + "\", \"" + word + "\", \"" + translation+"\");";

			}
			else 
			{
				sql = "insert into WordSerial(WordId, Word, Translation) values('" + wordId + "', '" + word + "', '" + translation+"');";
			}
			std::string::iterator it = sql.begin();
			while (it != sql.end())
			{
				if (*it == '\r' || *it == '\n')
				{
					it =sql.erase(it);

				}else {


					it++;
				}
			}
			file << sql<<std::endl;
		}
		file.close();
		pWorkbook->Close(VARIANT_TRUE);  // save changes
		//		}


		/*try {
		pWorkbook = pWorkbooks->Open("d:\\result.xls");  // open excel file
		} catch (...) {
		pWorkbook = pWorkbooks->Add();
		}

		Excel::_WorksheetPtr pWorksheet;
		try {
		pWorksheet = pWorkbook->Sheets->Item[L"Sheet1"];
		} catch(...) {
		pWorksheet = pWorkbook->Sheets->Add();
		pWorksheet->Name = L"result";
		}

		Excel::RangePtr pRange = pWorksheet->Cells;
		pRange->Item[1][1] = value;

		CString strSaveAsName = _T("d:\\new2.xls");
		// CString strSuffix = strSaveAsName.Mid(strSaveAsName.ReverseFind(_T('.')));
		//XlFileFormat NewFileFormat = xlOpenXMLWorkbook;
		//if (0 == strSuffix.CompareNoCase(_T(".xls")))
		//{
		//NewFileFormat = xlExcel8;
		//}
		//COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

		pWorkbook->SaveAs(_variant_t(strSaveAsName), _variant_t(Excel::xlText), vtMissing, vtMissing, vtMissing, 
		vtMissing, Excel::XlSaveAsAccessMode::xlNoChange, vtMissing, vtMissing, vtMissing, 
		vtMissing, vtMissing);
		pWorkbook->Close(VARIANT_TRUE);  // save changes

		pExcelApp->Quit();*/
	}
	catch (_com_error& error)
	{
		//ATLASSERT(FALSE);
		//ATLTRACE2(error.ErrorMessage());
	}
}
