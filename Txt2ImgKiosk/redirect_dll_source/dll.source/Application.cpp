
// Application.cpp : Implementation of CApplication
#include "stdafx.h"
#include "Redirect.h"
#include "Application.h"

/////////////////////////////////////////////////////////////////////////////
// CApplication

DWORD WINAPI ProcessThread(LPVOID lpApp);

STDMETHODIMP CApplication::get_Name(BSTR *pVal)
{
	// TODO: Add your implementation code here
	_bstr_t bstrVal(*pVal, false);
	bstrVal = m_bstrName;
	return S_OK;
}

STDMETHODIMP CApplication::put_Name(BSTR newVal)
{
	// TODO: Add your implementation code here
	m_bstrName = newVal;
	return S_OK;
}

STDMETHODIMP CApplication::Start(eStartResult *result)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;

	m_dwLastError = 0;

	if ( m_dwProcessId )
	{
		*result = laAlreadyRunning;
		return S_OK;
	}

	PROCESS_INFORMATION pi;
	SECURITY_ATTRIBUTES sa;
	STARTUPINFO si;
	DWORD dwThreadId;
	HANDLE hPipeStdoutReadTmp;
	HANDLE hPipeStdinWriteTmp;
	HANDLE hPipeStdoutWrite;
	HANDLE hPipeStdinRead;

	HANDLE hSaveStdout;
	HANDLE hSaveStdin;

	sa.nLength = sizeof(SECURITY_ATTRIBUTES);
	sa.bInheritHandle = TRUE;
	sa.lpSecurityDescriptor = NULL;

	// Redirect Stdout

	hSaveStdout = GetStdHandle(STD_OUTPUT_HANDLE);
	if ( ! CreatePipe(&hPipeStdoutReadTmp, &hPipeStdoutWrite, &sa, 0) )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}
	SetStdHandle(STD_OUTPUT_HANDLE, hPipeStdoutWrite);

	if ( ! DuplicateHandle(GetCurrentProcess(), hPipeStdoutReadTmp,
						   GetCurrentProcess(), &m_hPipeStdoutRead, 
						   0, FALSE, DUPLICATE_SAME_ACCESS) )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}

	CloseHandle(hPipeStdoutReadTmp);

	// Redirect Stdin
	hSaveStdin = GetStdHandle(STD_INPUT_HANDLE);
	if ( ! CreatePipe(&hPipeStdinRead, &hPipeStdinWriteTmp, &sa, 0) )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}
	SetStdHandle(STD_INPUT_HANDLE, hPipeStdinRead);

	if ( ! DuplicateHandle(GetCurrentProcess(), hPipeStdinWriteTmp,
						   GetCurrentProcess(), &m_hPipeStdinWrite,
						   0, FALSE, DUPLICATE_SAME_ACCESS) )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}

	CloseHandle(hPipeStdinWriteTmp);	

	ZeroMemory(&si, sizeof(STARTUPINFO));
	ZeroMemory(&pi, sizeof(PROCESS_INFORMATION));

	si.cb = sizeof(STARTUPINFO);
	si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE;
	si.hStdOutput = hPipeStdoutWrite;
	si.hStdError = hPipeStdoutWrite;
	si.hStdInput = hPipeStdinRead;

	if ( ! CreateProcess(NULL, OLE2T(m_bstrName), NULL, NULL, TRUE,
						 CREATE_NEW_CONSOLE, NULL, NULL, &si, &pi) )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}

	m_dwProcessId = pi.dwProcessId;

	SetStdHandle(STD_OUTPUT_HANDLE, hSaveStdout);
	SetStdHandle(STD_INPUT_HANDLE, hSaveStdin);

	CloseHandle(hPipeStdoutWrite);
	CloseHandle(hPipeStdinRead);

	if ( (m_hThread = CreateThread(NULL, 0, ProcessThread, (LPVOID) this, 0, &dwThreadId)) == NULL )
	{
		*result = laWindowsError;
		m_dwLastError = GetLastError();
		return S_OK;
	}

	*result = laOk;
	return S_OK;
}

DWORD WINAPI ProcessThread(LPVOID lpApp)
{
	CApplication *pApp = (CApplication *) lpApp;
	HANDLE hProcess = OpenProcess(SYNCHRONIZE, FALSE, pApp->m_dwProcessId);
	while(1)
	{
		pApp->read();
		if ( WaitForSingleObject(hProcess, pApp->m_lWait) == WAIT_OBJECT_0 )
		{
			while (pApp->read());
			pApp->PostMessage(WM_PROCESS_ENDED, 0, 0);
			CloseHandle(hProcess);
			pApp->m_dwProcessId = 0;
			break;
		}
	}
	if ( pApp->m_hPipeStdinWrite )
	{
		CloseHandle(pApp->m_hPipeStdinWrite);
		pApp->m_hPipeStdinWrite = NULL;
	}

	if ( pApp->m_hPipeStdoutRead )
	{
		CloseHandle(pApp->m_hPipeStdoutRead);
		pApp->m_hPipeStdoutRead = NULL;
	}
	if ( pApp->m_hThread )
	{
		CloseHandle(pApp->m_hThread);
		pApp->m_hThread = NULL;
	}
	return 0;
}

STDMETHODIMP CApplication::Stop()
{
	// TODO: Add your implementation code here

	if ( m_dwProcessId )
	{
		HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, m_dwProcessId);
		TerminateProcess(hProcess, 0);
		CloseHandle(hProcess);
		m_dwProcessId = 0;
	}
	return S_OK;
}

bool CApplication::read(void)
{
	DWORD dwBytesLeft = 0;
	DWORD dwBytesRead = 0;
	DWORD dwBytesTotalAvailable = 0;
	DWORD dwBytesWritten = 0;

	LPTSTR pszData = (LPTSTR) VirtualAlloc(NULL, m_nBufferSize, MEM_COMMIT, PAGE_READWRITE);

	if ( ! ReadFile(m_hPipeStdoutRead, pszData, m_nBufferSize - 1,
				  &dwBytesRead, NULL) )
		return false;

	pszData[dwBytesRead] = '\0';

	for(DWORD i = 0; i < dwBytesRead; i++)
	{
		if ( pszData[i] == _T('\b') )
			 pszData[i] = ' ';
	}

	PostMessage(WM_DATA_RECEIVED, (WPARAM) pszData, 0); // We post this, because it is called from a thread !!!

	return true;
}

LRESULT CApplication::OnDataReceived(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	Fire_DataReceived(_bstr_t((LPTSTR) wParam));
	VirtualFree((LPVOID) wParam, 0, MEM_RELEASE);
	return 0;
}

STDMETHODIMP CApplication::get_Running(VARIANT_BOOL *pVal)
{
	// TODO: Add your implementation code here
	*pVal = (m_dwProcessId != 0) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CApplication::get_BufferSize(short *pVal)
{
	// TODO: Add your implementation code here
	*pVal = m_nBufferSize;	
	return S_OK;
}

STDMETHODIMP CApplication::put_BufferSize(short newVal)
{
	// TODO: Add your implementation code here
	m_nBufferSize = newVal;
	return S_OK;
}

STDMETHODIMP CApplication::get_Wait(long *pVal)
{
	// TODO: Add your implementation code here
	*pVal = m_lWait;
	return S_OK;
}

STDMETHODIMP CApplication::put_Wait(long newVal)
{
	// TODO: Add your implementation code here
	m_lWait = newVal;
	return S_OK;
}

STDMETHODIMP CApplication::Write(BSTR sCommandString, VARIANT_BOOL *pVal)
{
	// TODO: Add your implementation code here
	USES_CONVERSION;
	DWORD dwWritten;

	if ( ! WriteFile(m_hPipeStdinWrite, OLE2T(sCommandString), SysStringLen(sCommandString), &dwWritten, NULL) )
	{
		*pVal = VARIANT_FALSE;
		m_dwLastError = GetLastError();
		return S_OK;
	}
	*pVal = VARIANT_TRUE;
	return S_OK;
}

STDMETHODIMP CApplication::get_LastErrorNumber(long *pVal)
{
	// TODO: Add your implementation code here
	*pVal = m_dwLastError;
	return S_OK;
}
