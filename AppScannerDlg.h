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

// AppScannerDlg.h : header file
//

#pragma once

struct CMachineAddRemove
{
	CString sDisplayName;
	CString sDisplayVersion;
	CString sPublisher;
	CString sInstallLocation;
	CString sUninstallString;
	CString sSystemComponent;
};

struct CMachineAddRemoveCompare {
	bool operator()(const CMachineAddRemove& lhs, const CMachineAddRemove& rhs) const {
		if (lhs.sDisplayName != rhs.sDisplayName)
			return lhs.sDisplayName < rhs.sDisplayName;
		if (lhs.sDisplayVersion != rhs.sDisplayVersion)
			return lhs.sDisplayVersion < rhs.sDisplayVersion;
		return lhs.sPublisher < rhs.sPublisher;
	}
};

struct CMachineFile
{
	CString sFilePath;
	CString sFileName;
	CString sFileSize;
	CString sFileVersion;
	CString sProductVersion;
	CString sFileDescription;
	CString sCompanyName;
	CString sProductName;
};

#include "afxwin.h"


// CAppScannerDlg dialog
class CAppScannerDlg : public CDialogEx
{
// Construction
public:
	CAppScannerDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_APPSCANNER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	bool bAppFolder;
	bool bMachine;
	bool bPkgFolder;
	afx_msg void OnBnClickedRadioMachine();
	CButton buttonAppFolder;
	CEdit editAppFolder;
	afx_msg void OnBnClickedRadioAppfolder();
	afx_msg void OnBnClickedButtonBrowseapppath();
	afx_msg void OnBnClickedButtonBrowsescanpath();
	CEdit editScanPath;
	void UpdateScanPath(void);
	CEdit editPublisher;
	CEdit editProduct;
	CEdit editVersion;
	CString sScanPath;
	afx_msg void OnBnClickedButtonScan();
	void ScanPathRecursively(CString sPath, CStdioFile& scan);
	void EnumerateHKEY_LOCAL_MACHINESubKeys(CString sRegKey, set<CMachineAddRemove, CMachineAddRemoveCompare>& machine_addremoves);
	void EnumerateHKEY_CURRENT_USERSubKeys(CString sRegKey, set<CMachineAddRemove, CMachineAddRemoveCompare>& machine_addremoves);
	string RetrieveManifest(string sFilePath);
	string ParseSWIDTag(string sFilePath, CMachineFile& mf);
	CEdit editPkgFolder;
	afx_msg void OnBnClickedRadioPkgfolder();
	afx_msg void OnBnClickedButtonBrowsepkgpath();
	CString GetMessageFromMsiErrCode(DWORD error);
	CString RetrievePropertyFromMSI(CString sMSIPath, CString sPropertyName);
};
