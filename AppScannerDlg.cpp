/*
Copyright 2015 Inferapp

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// AppScannerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "AppScanner.h"
#include "AppScannerDlg.h"
#include "afxdialogex.h"
#include "file_ver.h"
#include <CkZip.h>
#include <CkZipEntry.h>
#include <CkXml.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
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


// CAppScannerDlg dialog




CAppScannerDlg::CAppScannerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CAppScannerDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CAppScannerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_RADIO_APPFOLDER, buttonAppFolder);
	DDX_Control(pDX, IDC_EDIT_APPFOLDER, editAppFolder);
	DDX_Control(pDX, IDC_EDIT_SCANPATH, editScanPath);
	DDX_Control(pDX, IDC_EDIT_PUBLISHER, editPublisher);
	DDX_Control(pDX, IDC_EDIT_PRODUCT, editProduct);
	DDX_Control(pDX, IDC_EDIT_VERSION, editVersion);
	DDX_Control(pDX, IDC_EDIT_PKGFOLDER, editPkgFolder);
}

BEGIN_MESSAGE_MAP(CAppScannerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_RADIO_MACHINE, &CAppScannerDlg::OnBnClickedRadioMachine)
	ON_BN_CLICKED(IDC_RADIO_APPFOLDER, &CAppScannerDlg::OnBnClickedRadioAppfolder)
	ON_BN_CLICKED(IDC_BUTTON_BROWSEAPPPATH, &CAppScannerDlg::OnBnClickedButtonBrowseapppath)
	ON_BN_CLICKED(IDC_BUTTON_BROWSESCANPATH, &CAppScannerDlg::OnBnClickedButtonBrowsescanpath)
	ON_BN_CLICKED(IDC_BUTTON_SCAN, &CAppScannerDlg::OnBnClickedButtonScan)
	ON_BN_CLICKED(IDC_RADIO_PKGFOLDER, &CAppScannerDlg::OnBnClickedRadioPkgfolder)
	ON_BN_CLICKED(IDC_BUTTON_BROWSEPKGPATH, &CAppScannerDlg::OnBnClickedButtonBrowsepkgpath)
END_MESSAGE_MAP()


// CAppScannerDlg message handlers

BOOL CAppScannerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
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

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	buttonAppFolder.SetCheck(BST_CHECKED);
	editPkgFolder.SetReadOnly(1);
	editPublisher.SetWindowText("NA");
	editProduct.SetWindowText("NA");
	editVersion.SetWindowText("NA");
	bAppFolder = true;
	bMachine = false;
	bPkgFolder = false;

	CkZip zip;
	bool success;

	//  Any string unlocks the component for the 1st 30-days.
	success = zip.UnlockComponent("Chilkat license key");
	if (success != true) {
		AfxMessageBox(zip.lastErrorText());
	}

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CAppScannerDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CAppScannerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CAppScannerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CAppScannerDlg::OnBnClickedRadioMachine()
{
	editAppFolder.SetReadOnly(1);
	editPkgFolder.SetReadOnly(1);
	editPublisher.SetReadOnly(1);
	editProduct.SetReadOnly(1);
	editVersion.SetReadOnly(1);
	bAppFolder = false;
	bMachine = true;
	bPkgFolder = false;
	UpdateScanPath();
}

void CAppScannerDlg::OnBnClickedRadioAppfolder()
{
	editAppFolder.SetReadOnly(0);
	editPkgFolder.SetReadOnly(1);
	editPublisher.SetReadOnly(0);
	editProduct.SetReadOnly(0);
	editVersion.SetReadOnly(0);
	bAppFolder = true;
	bMachine = false;
	bPkgFolder = false;
	UpdateScanPath();
}

void CAppScannerDlg::OnBnClickedRadioPkgfolder()
{
	editAppFolder.SetReadOnly(1);
	editPkgFolder.SetReadOnly(0);
	editPublisher.SetReadOnly(1);
	editProduct.SetReadOnly(1);
	editVersion.SetReadOnly(1);
	bAppFolder = false;
	bMachine = false;
	bPkgFolder = true;
	UpdateScanPath();
}

void CAppScannerDlg::OnBnClickedButtonBrowseapppath()
{
	BROWSEINFO bi = {0};
    bi.lpszTitle = _T("Select Folder");
    LPITEMIDLIST pidl = SHBrowseForFolder (&bi);
    if (pidl != 0)
    {
        // get the name of the folder
        TCHAR path[MAX_PATH];
        SHGetPathFromIDList (pidl, path);

        // free memory used
        IMalloc * imalloc = 0;
        if (SUCCEEDED(SHGetMalloc (&imalloc)))
        {
            imalloc->Free (pidl);
            imalloc->Release ();
        }

		CString sAppPath = path;
		if (!sAppPath.IsEmpty() && sAppPath.Right(1) != '\\')
			sAppPath += '\\';
		editAppFolder.SetWindowText(sAppPath);
		sScanPath = sAppPath;
		UpdateScanPath();
    }
}

void CAppScannerDlg::OnBnClickedButtonBrowsepkgpath()
{
	BROWSEINFO bi = {0};
    bi.lpszTitle = _T("Select Folder");
    LPITEMIDLIST pidl = SHBrowseForFolder (&bi);
    if (pidl != 0)
    {
        // get the name of the folder
        TCHAR path[MAX_PATH];
        SHGetPathFromIDList (pidl, path);

        // free memory used
        IMalloc * imalloc = 0;
        if (SUCCEEDED(SHGetMalloc (&imalloc)))
        {
            imalloc->Free (pidl);
            imalloc->Release ();
        }

		CString sPkgPath = path;
		if (!sPkgPath.IsEmpty() && sPkgPath.Right(1) != '\\')
			sPkgPath += '\\';
		editPkgFolder.SetWindowText(sPkgPath);
		sScanPath = sPkgPath;
		UpdateScanPath();
    }
}

void CAppScannerDlg::OnBnClickedButtonBrowsescanpath()
{
	BROWSEINFO bi = {0};
    bi.lpszTitle = _T("Select Folder");
    LPITEMIDLIST pidl = SHBrowseForFolder (&bi);
    if (pidl != 0)
    {
        // get the name of the folder
        TCHAR path[MAX_PATH];
        SHGetPathFromIDList (pidl, path);

        // free memory used
        IMalloc * imalloc = 0;
        if (SUCCEEDED(SHGetMalloc (&imalloc)))
        {
            imalloc->Free (pidl);
            imalloc->Release ();
        }

		sScanPath = path;
		if (!sScanPath.IsEmpty() && sScanPath.Right(1) != '\\')
			sScanPath += '\\';
		UpdateScanPath();
    }
}


void CAppScannerDlg::UpdateScanPath(void)
{
	if (sScanPath.IsEmpty())
		return;

	if (bAppFolder)
	{
		CString sPublisher, sProduct, sVersion;

		editPublisher.GetWindowText(sPublisher);
		editProduct.GetWindowText(sProduct);
		editVersion.GetWindowText(sVersion);
		editScanPath.SetWindowText(sScanPath + sPublisher + "_" + sProduct + "_" + sVersion + ".scan");
	}
	if (bMachine)
	{
		char buf[1024];
		DWORD dwCompNameLen = 1024;
		CString sMachineName;

		if (0 != GetComputerName(buf, &dwCompNameLen))
		{
			sMachineName.Format("%s", buf);
			sMachineName.MakeLower();
		}
		else
			sMachineName = "machine";
		editScanPath.SetWindowText(sScanPath + sMachineName + ".scan");
	}
	if (bPkgFolder)
		editScanPath.SetWindowText(sScanPath);
}


void CAppScannerDlg::OnBnClickedButtonScan()
{
	CString sAppFolder;
	editAppFolder.GetWindowText(sAppFolder);

	if (bAppFolder && sAppFolder.IsEmpty())
	{
		AfxMessageBox("Please select the applicaton folder to be scanned.");
		return;
	}

	CString sPkgFolder;
	editPkgFolder.GetWindowText(sPkgFolder);

	if (bPkgFolder && sPkgFolder.IsEmpty())
	{
		AfxMessageBox("Please select the package folder to be scanned.");
		return;
	}

	CString sScanPath;
	editScanPath.GetWindowText(sScanPath);

	if (sScanPath.IsEmpty())
	{
		AfxMessageBox("Please select the path for the scan(s).");
		return;
	}

	CWaitCursor wait;
	
	if (bAppFolder || bMachine)
	{
		CStdioFile scan (sScanPath, CFile::modeCreate | CFile::modeWrite);

		if (bMachine || (bAppFolder && GetDriveType(sAppFolder.Left(3)) == DRIVE_FIXED))
		{
			set<CMachineAddRemove, CMachineAddRemoveCompare> machine_addremoves;

			EnumerateHKEY_LOCAL_MACHINESubKeys("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", machine_addremoves);
			EnumerateHKEY_LOCAL_MACHINESubKeys("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall", machine_addremoves);

			EnumerateHKEY_CURRENT_USERSubKeys("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", machine_addremoves);
			EnumerateHKEY_CURRENT_USERSubKeys("SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall", machine_addremoves);			

			scan.WriteString ("<SourceName=AddRemoves><Fields=DisplayName		DisplayVersion	Publisher	InstallLocation	UninstallString		SystemComponent>\n");

			for (auto it = machine_addremoves.begin(); it != machine_addremoves.end(); it++)
				scan.WriteString (it->sDisplayName + "\t" + it->sDisplayVersion + "\t" + it->sPublisher + "\t" + it->sInstallLocation + "\t" + it->sUninstallString + "\t" + it->sSystemComponent + "\n");
		}

		scan.WriteString ("<SourceName=Files><Fields=FilePath	FileName	ProductVersion	CompanyName	ProductName	FileDescription	FileVersion	FileSize>\n");

		if (bAppFolder)
			ScanPathRecursively(sAppFolder, scan);
		if (bMachine)
		{
			// initial value
			TCHAR szDrive[ ] = "A:\\";
			DWORD uDriveMask = GetLogicalDrives();
			//printf("The bitmask of the logical drives in hex: %0X\n", uDriveMask);
			//printf("The bitmask of the logical drives in decimal: %d\n", uDriveMask);
			if(uDriveMask == 0)
			{
				AfxMessageBox("GetLogicalDrives() failed with failure code: " + GetLastError());
				return;
			}
			else
			{
				while(uDriveMask)
				{// use the bitwise AND, 1–available, 0-not available
					if(uDriveMask & 1 && GetDriveType(szDrive) == DRIVE_FIXED)
						ScanPathRecursively((CString)szDrive, scan);
					// increment... 
					++szDrive[0];
					// shift the bitmask binary right
					uDriveMask >>= 1;
				}
			}
		}
	}
	if (bPkgFolder)
	{
		CString sTmpFolder = sScanPath + "appscanner_tmp";

		// a test if the tmp folder is creatable and removable
		try
		{
			filesystem::create_directory((string)sTmpFolder);
			filesystem::remove_all((string)sTmpFolder);
		}
		catch(boost::filesystem::filesystem_error &ex)
		{
			AfxMessageBox(ex.what());
			return;
		}		

		try {
		for (filesystem::recursive_directory_iterator it((string)sPkgFolder); it != filesystem::recursive_directory_iterator(); it++)
			if (is_regular_file(*it) && to_lower_copy(it->path().extension().string()) == ".msi")
			{
				filesystem::create_directory((string)sTmpFolder);

				DWORD result = MsiInstallProduct(it->path().string().c_str(), "ACTION=ADMIN TARGETDIR=" + sTmpFolder);
				if (result != ERROR_SUCCESS)
					continue;

				CStdioFile scan (sScanPath + it->path().filename().string().c_str() + ".scan", CFile::modeCreate | CFile::modeWrite);

				scan.WriteString ("<SourceName=AddRemoves><Fields=DisplayName		DisplayVersion	Publisher	InstallLocation	UninstallString		SystemComponent>\n");

				CMachineAddRemove ma;

				ma.sDisplayName = RetrievePropertyFromMSI(it->path().string().c_str(), "ProductName");
				ma.sDisplayVersion = RetrievePropertyFromMSI(it->path().string().c_str(), "ProductVersion");
				ma.sPublisher = RetrievePropertyFromMSI(it->path().string().c_str(), "Manufacturer");

				scan.WriteString (ma.sDisplayName + "\t" + ma.sDisplayVersion + "\t" + ma.sPublisher + "\t" + "\t" + "\t" + "\n");

				scan.WriteString ("<SourceName=Files><Fields=FilePath	FileName	ProductVersion	CompanyName	ProductName	FileDescription	FileVersion	FileSize>\n");

				ScanPathRecursively(sTmpFolder, scan);

				filesystem::remove_all((string)sTmpFolder);
			}
		}catch(boost::filesystem::filesystem_error &ex){AfxMessageBox(ex.what());}		
	}

	wait.Restore();
}

CString CAppScannerDlg::RetrievePropertyFromMSI(CString sMSIPath, CString sPropertyName)
{
	MSIHANDLE hInstall = NULL;

	MsiOpenPackage(sMSIPath, &hInstall);
	
	TCHAR* szValueBuf = NULL;
    DWORD cchValueBuf = 0;
    UINT uiStat =  MsiGetProperty(hInstall, sPropertyName, TEXT(""), &cchValueBuf);
    //cchValueBuf now contains the size of the property's string, without null termination
    if (ERROR_MORE_DATA == uiStat)
    {
        ++cchValueBuf; // add 1 for null termination
        szValueBuf = new TCHAR[cchValueBuf];
        if (szValueBuf)
        {
            uiStat = MsiGetProperty(hInstall, sPropertyName, szValueBuf, &cchValueBuf);
        }
    }
    if (ERROR_SUCCESS != uiStat)
    {
        if (szValueBuf != NULL) 
           delete[] szValueBuf;
        return "";
    }

	CString sPropertyValue;
	sPropertyValue.Format("%s", szValueBuf);

	delete[] szValueBuf;
	MsiCloseHandle(hInstall);

	return sPropertyValue;
}

CString CAppScannerDlg::GetMessageFromMsiErrCode(DWORD error)
{
	CString message = _T("Unknown error");

	HMODULE module = LoadLibraryEx(_T("msimsg.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);

	if (module == NULL) return message;

	TCHAR buffer[8192] = _T("");

	LoadString(module, error, buffer, sizeof(buffer) / sizeof(TCHAR));

	FreeLibrary(module);

	message = buffer;

	return message;
}



void CAppScannerDlg::ScanPathRecursively(CString sPath, CStdioFile& scan)
{
	try {
	for (filesystem::recursive_directory_iterator it((string)sPath); it != filesystem::recursive_directory_iterator(); it++)
		if (is_regular_file(*it))
		{
			if (to_lower_copy(it->path().extension().string()) != ".exe" && to_lower_copy(it->path().extension().string()) != ".ocx" && to_lower_copy(it->path().extension().string()) != ".dll" &&
				to_lower_copy(it->path().extension().string()) != ".jar" && to_lower_copy(it->path().extension().string()) != ".war" && to_lower_copy(it->path().extension().string()) != ".ear" &&
				to_lower_copy(it->path().extension().string()) != ".swidtag" && to_lower_copy(it->path().extension().string()) != ".swtag")
				continue;

			CMachineFile mf;

			mf.sFilePath = it->path().parent_path().c_str();
			if (mf.sFilePath.Right(1) != '\\')
				mf.sFilePath += '\\';

			int index = mf.sFilePath.Find("appscanner_tmp");

			if (index != -1)
			{
				mf.sFilePath = mf.sFilePath.Mid(index);
				mf.sFilePath.Replace("appscanner_tmp", "C:");
			}

			mf.sFileName = it->path().filename().c_str();

			mf.sFileSize.Format("%d", filesystem::file_size(it->path()));
			
			if (to_lower_copy(it->path().extension().string()) == ".exe" || to_lower_copy(it->path().extension().string()) == ".ocx" || to_lower_copy(it->path().extension().string()) == ".dll")
			{
				CFileVersionInfo fvi;
				fvi.ReadVersionInfo((CString)it->path().c_str());

				if(fvi.IsValid())
				{
					mf.sProductVersion = fvi.GetVersionInfo(SFI_PRODUCTVERSION);
					mf.sCompanyName = fvi.GetVersionInfo(SFI_COMPANYNAME);
					mf.sProductName = fvi.GetVersionInfo(SFI_PRODUCTNAME);
					mf.sFileDescription = fvi.GetVersionInfo(SFI_FILEDESCRIPTION);
					mf.sFileVersion = fvi.GetVersionInfo(SFI_FILEVERSION);
				}
			}

			if (to_lower_copy(it->path().extension().string()) == ".jar" || to_lower_copy(it->path().extension().string()) == ".war" || to_lower_copy(it->path().extension().string()) == ".ear")
			{
				string manifest = RetrieveManifest(it->path().string());
				if (!manifest.empty())
				{
					smatch match;

					// company name

					if (regex_search(manifest, match, regex("Specification-Vendor: (.*)")))
						mf.sCompanyName = match[1].str().c_str();

					if (mf.sCompanyName.IsEmpty())
						if(regex_search(manifest, match, regex("Implementation-Vendor: (.*)")))
							mf.sCompanyName = match[1].str().c_str();
					
					if (mf.sCompanyName.IsEmpty())
						if(regex_search(manifest, match, regex("Bundle-Vendor: (.*)")))
							mf.sCompanyName = match[1].str().c_str();

					// product name

					if (regex_search(manifest, match, regex("Specification-Title: (.*)")))
						mf.sProductName = match[1].str().c_str();

					if (mf.sProductName.IsEmpty())
						if (regex_search(manifest, match, regex("Implementation-Title: (.*)")))
							mf.sProductName = match[1].str().c_str();

					if (mf.sProductName.IsEmpty())
						if (regex_search(manifest, match, regex("Bundle-Name: (.*)")))
							mf.sProductName = match[1].str().c_str();
					
					if (mf.sProductName.IsEmpty())
						if (regex_search(manifest, match, regex("Extension-Name: (.*)")))
							mf.sProductName = match[1].str().c_str();

					// product version

					if (regex_search(manifest, match, regex("Implementation-Version: (.*)")))
						mf.sProductVersion = match[1].str().c_str();

					if (mf.sProductVersion.IsEmpty())
						if (regex_search(manifest, match, regex("Specification-Version: (.*)")))
							mf.sProductVersion = match[1].str().c_str();

					if (mf.sProductVersion.IsEmpty())
						if (regex_search(manifest, match, regex("Bundle-Version: (.*)")))
							mf.sProductVersion = match[1].str().c_str();
				}
			}
			if (to_lower_copy(it->path().extension().string()) == ".swidtag" || to_lower_copy(it->path().extension().string()) == ".swtag")
			{
				CkString strOut;
				CkXml xml;

				if (!xml.LoadXmlFile(it->path().string().c_str()))
					return;


				if (xml.FindChild2("swid:product_title"))
					mf.sProductName = xml.content();
				
				xml.GetRoot2();
				if (xml.FindChild2("swid:product_version"))
					if (xml.FindChild2("swid:name"))
						mf.sProductVersion = xml.content();

				xml.GetRoot2();
				if (xml.FindChild2("swid:software_creator"))
					if(xml.FindChild2("swid:name"))
						mf.sCompanyName = xml.content();
			}

			scan.WriteString (mf.sFilePath + "\t" + mf.sFileName + "\t" + mf.sProductVersion + "\t" + mf.sCompanyName + "\t" + 
					mf.sProductName + "\t" + mf.sFileDescription + "\t" + mf.sFileVersion + "\t" + mf.sFileSize + "\n");
		}
	}catch(boost::filesystem::filesystem_error &ex){ cout << ex.what();}
}

void CAppScannerDlg::EnumerateHKEY_LOCAL_MACHINESubKeys(CString sRegKey, set<CMachineAddRemove, CMachineAddRemoveCompare>& machine_addremoves)
{
	CRegKey regKey, regSubKey;      

    if(regKey.Open(HKEY_LOCAL_MACHINE, sRegKey)==ERROR_SUCCESS)
    {
        DWORD dwSize = MAX_PATH, dwIndex = 0;
        TCHAR tcsStringName[MAX_PATH];

        while(RegEnumKey(regKey.m_hKey, dwIndex, tcsStringName, dwSize) == ERROR_SUCCESS)
        {
			CMachineAddRemove ma;

			regSubKey.Open(HKEY_LOCAL_MACHINE, sRegKey + "\\" + tcsStringName);
			ULONG len = 255;

			if(regSubKey.QueryStringValue("DisplayName", ma.sDisplayName.GetBufferSetLength(len), &len) != ERROR_SUCCESS)
			{
				regSubKey.Close();
				dwIndex++;
				continue;
			}
			ma.sDisplayName.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("DisplayVersion", ma.sDisplayVersion.GetBufferSetLength(len), &len);
			ma.sDisplayVersion.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("Publisher", ma.sPublisher.GetBufferSetLength(len), &len);
			ma.sPublisher.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("InstallLocation", ma.sInstallLocation.GetBufferSetLength(len), &len);
			ma.sInstallLocation.ReleaseBuffer();
				
			len = 255;
			regSubKey.QueryStringValue("UninstallString", ma.sUninstallString.GetBufferSetLength(len), &len);
			ma.sUninstallString.ReleaseBuffer();

			DWORD buf;

			if (regSubKey.QueryDWORDValue("SystemComponent", buf) == ERROR_SUCCESS)
				ma.sSystemComponent.Format("%d", buf);
			else
				ma.sSystemComponent = "0";

			machine_addremoves.insert(ma);

			regSubKey.Close();
            dwIndex++;
        }
    }
}

void CAppScannerDlg::EnumerateHKEY_CURRENT_USERSubKeys(CString sRegKey, set<CMachineAddRemove, CMachineAddRemoveCompare>& machine_addremoves)
{
	CRegKey regKey, regSubKey;      

    if(regKey.Open(HKEY_CURRENT_USER, sRegKey)==ERROR_SUCCESS)
    {
        DWORD dwSize = MAX_PATH, dwIndex = 0;
        TCHAR tcsStringName[MAX_PATH];

        while(RegEnumKey(regKey.m_hKey, dwIndex, tcsStringName, dwSize) == ERROR_SUCCESS)
        {
			CMachineAddRemove ma;

			regSubKey.Open(HKEY_CURRENT_USER, sRegKey + "\\" + tcsStringName);
			ULONG len = 255;

			if(regSubKey.QueryStringValue("DisplayName", ma.sDisplayName.GetBufferSetLength(len), &len) != ERROR_SUCCESS)
			{
				regSubKey.Close();
				dwIndex++;
				continue;
			}
			ma.sDisplayName.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("DisplayVersion", ma.sDisplayVersion.GetBufferSetLength(len), &len);
			ma.sDisplayVersion.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("Publisher", ma.sPublisher.GetBufferSetLength(len), &len);
			ma.sPublisher.ReleaseBuffer();

			len = 255;
			regSubKey.QueryStringValue("InstallLocation", ma.sInstallLocation.GetBufferSetLength(len), &len);
			ma.sInstallLocation.ReleaseBuffer();
				
			len = 255;
			regSubKey.QueryStringValue("UninstallString", ma.sUninstallString.GetBufferSetLength(len), &len);
			ma.sUninstallString.ReleaseBuffer();

			DWORD buf;

			if (regSubKey.QueryDWORDValue("SystemComponent", buf) == ERROR_SUCCESS)
				ma.sSystemComponent.Format("%d", buf);
			else
				ma.sSystemComponent = "0";

			machine_addremoves.insert(ma);

			regSubKey.Close();
            dwIndex++;
        }
    }
}

string CAppScannerDlg::RetrieveManifest(string sFilePath)
{
	bool success;
	CkString strOut;
	CkZip zip;

	success = zip.OpenZip(sFilePath.c_str());
    if (success != true) {
        AfxMessageBox(zip.lastErrorText());
        return "";
    }

	CkZipEntry *entry = 0;
	entry = zip.FirstMatchingEntry("*MANIFEST.MF");
	if (!(entry == 0 )) {

		//  lineEndingBehavior:
		//  0 = leave unchanged.
		//  1 = convert all to bare LF's
		//  2 = convert all to CRLF's
		int lineEndingBehavior;
		lineEndingBehavior = 0;
		const char * srcCharset;
		srcCharset = "utf-8";

		const char * strCsv;
		strCsv = entry->unzipToString(lineEndingBehavior,srcCharset);
		strOut.append(strCsv);
		strOut.append("\r\n");
		delete entry;
	}
	
	return strOut.getString();
}
