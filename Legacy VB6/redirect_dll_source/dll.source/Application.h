// Application.h : Declaration of the CApplication

#ifndef __APPLICATION_H_
#define __APPLICATION_H_

#include <atlwin.h>
#include "comdef.h"
#include "resource.h"       // main symbols
#include "RedirectCP.h"

#define WM_DATA_RECEIVED (WM_USER + 101)
#define WM_PROCESS_ENDED (WM_USER + 102)

/////////////////////////////////////////////////////////////////////////////
// CApplication
class ATL_NO_VTABLE CApplication : 
	public CWindowImpl<CApplication>,
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CApplication, &CLSID_Application>,
	public IConnectionPointContainerImpl<CApplication>,
	public IDispatchImpl<IApplication, &IID_IApplication, &LIBID_RedirectLib>,
	public CProxy_IApplicationEvents< CApplication >,
    public ISupportErrorInfoImpl< &IID_IApplication >

{
public:
	CApplication()
	{
		m_hThread = NULL;
		m_hPipeStdoutRead = NULL;
		m_hPipeStdinWrite = NULL;
		m_nBufferSize = 8192;
		m_lWait = 0;
		m_dwProcessId = 0;
	}
	_bstr_t m_bstrName;
	HANDLE m_hThread;
	DWORD m_dwProcessId;
	HANDLE m_hPipeStdoutRead;
	HANDLE m_hPipeStdinWrite;

	short m_nBufferSize;
	long m_lWait;
	DWORD m_dwLastError;
DECLARE_WND_CLASS(TEXT("Application"))
BEGIN_MSG_MAP(CApplication)
	MESSAGE_HANDLER(WM_DATA_RECEIVED, OnDataReceived)
	MESSAGE_HANDLER(WM_PROCESS_ENDED, OnProcessEnded)
END_MSG_MAP()

	LRESULT OnDataReceived(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
	LRESULT OnProcessEnded(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		Fire_ProcessEnded();
		return 0;
	}
	HRESULT FinalConstruct()
	{
		RECT rect;
		rect.left = 0;
		rect.right = 100;
		rect.top = 0;
		rect.bottom = 100;
		HWND hwnd = Create(NULL, rect, TEXT("WatchWindow"), WS_POPUP);
		if ( hwnd ) 
			return S_OK;
		else
			return HRESULT_FROM_WIN32(GetLastError());
	}
	void FinalRelease()
	{
		if ( m_dwProcessId )
			Stop();
		if ( m_hWnd != NULL )
			DestroyWindow();
	}

DECLARE_REGISTRY_RESOURCEID(IDR_APPLICATION)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CApplication)
	COM_INTERFACE_ENTRY(IApplication)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
END_COM_MAP()
BEGIN_CONNECTION_POINT_MAP(CApplication)
CONNECTION_POINT_ENTRY(DIID__IApplicationEvents)
END_CONNECTION_POINT_MAP()


// IApplication
public:
	STDMETHOD(get_LastErrorNumber)(/*[out, retval]*/ long *pVal);
	STDMETHOD(Write)(/*[in]*/ BSTR sCommandString, /*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_Wait)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_Wait)(/*[in]*/ long newVal);
	STDMETHOD(get_BufferSize)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_BufferSize)(/*[in]*/ short newVal);
	STDMETHOD(get_Running)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	bool read(void);
	STDMETHOD(Stop)();
	STDMETHOD(Start)(/*[out, retval]*/ eStartResult *result);
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Name)(/*[in]*/ BSTR newVal);
	};

#endif //__APPLICATION_H_
