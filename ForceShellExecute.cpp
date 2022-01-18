#include "framework.h"
#include "ForceShellExecute.h"

LPCTSTR szForceShellExecute = TEXT("ForceShellExecute");
LPCTSTR szHklmShellExec = TEXT("SOFTWARE\\Microsoft\\Office\\9.0\\Common\\Internet");
LPCTSTR szHklmShellExecWOW6432 = TEXT("SOFTWARE\\WOW6432Node\\Microsoft\\Office\\9.0\\Common\\Internet");
const DWORD dwShellExecValue = 0x00000001;

BOOL CreateRegKey(HKEY hKeyParent, LPCTSTR subKey)
{
	HKEY  hKey;
	DWORD dwDisposition;
	if (ERROR_SUCCESS != RegCreateKeyEx(hKeyParent, subKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dwDisposition))
		return FALSE;
	RegCloseKey(hKey);
	return TRUE;
}

BOOL WriteRegDword(HKEY hKeyParent, LPCTSTR subKey, LPCTSTR valueName, DWORD data)
{
	HKEY hKey;
	if (ERROR_SUCCESS != RegOpenKeyEx(hKeyParent, subKey, 0, KEY_WRITE, &hKey))
		return FALSE;
	BOOL bRet = (ERROR_SUCCESS != RegSetValueEx(hKey, valueName, 0, REG_DWORD, reinterpret_cast<BYTE*>(&data), sizeof(data)));
	RegCloseKey(hKey);
	return bRet;
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
	if (CreateRegKey(HKEY_LOCAL_MACHINE, szHklmShellExec))
		WriteRegDword(HKEY_LOCAL_MACHINE, szHklmShellExec, szForceShellExecute, dwShellExecValue);
	else
		MessageBox(NULL, TEXT("Error - Faled to create ForceShellExecute 64-bit entries"), TEXT("ForceShellExecute"), MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
	if (CreateRegKey(HKEY_LOCAL_MACHINE, szHklmShellExecWOW6432))
		WriteRegDword(HKEY_LOCAL_MACHINE, szHklmShellExecWOW6432, szForceShellExecute, dwShellExecValue);
	else
		MessageBox(NULL, TEXT("Error - Faled to create ForceShellExecute 32-bit entries"), TEXT("ForceShellExecute"), MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);

	MessageBox(NULL, TEXT("Success - Registry Entries for ForceShellExecute created successfully"), TEXT("ForceShellExecute"), MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL);
	return 0;
}
