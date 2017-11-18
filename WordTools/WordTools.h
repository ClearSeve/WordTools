// WordTools.h : CWordTools 的声明

#pragma once
#include "resource.h"       // 主符号



#include "WordTools_i.h"
#include "ContentWindow.h"
#include <ShlGuid.h>
#include <Shobjidl.h>
#include <Shlobj.h>
#include "ContentWindow.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


// CWordTools

class ATL_NO_VTABLE CWordTools :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CWordTools, &CLSID_WordTools>,
	public IObjectWithSiteImpl<CWordTools>,
	public IDeskBand2,
	public IPersistStreamInitImpl<CWordTools>,
	public IContextMenu
	
{
	typedef IPersistStreamInitImpl<CWordTools> IPersistStreamImpl;
public:
	CWordTools()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_WORDTOOLS1)

DECLARE_NOT_AGGREGATABLE(CWordTools)

BEGIN_COM_MAP(CWordTools)
	COM_INTERFACE_ENTRY(IContextMenu)
	COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY(IOleWindow)
	COM_INTERFACE_ENTRY(IDockingWindow)
	COM_INTERFACE_ENTRY(IDeskBand)
	COM_INTERFACE_ENTRY(IDeskBand2)
	COM_INTERFACE_ENTRY_IID(IID_IPersist, IPersistStreamImpl)
	COM_INTERFACE_ENTRY_IID(IID_IPersistStream, IPersistStreamImpl)
	COM_INTERFACE_ENTRY_IID(IID_IPersistStreamInit, IPersistStreamImpl)
END_COM_MAP()

BEGIN_CATEGORY_MAP(CWordTools)
	IMPLEMENTED_CATEGORY(CATID_DeskBand)
END_CATEGORY_MAP()

BEGIN_PROP_MAP(CWordTools)
	PROP_DATA_ENTRY("TestData", m_DataSettingDialog.nData, VT_UI4)
END_PROP_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	// IObjectWithSite
	//
	STDMETHOD(SetSite)(IUnknown* pUnkSite);

	// IDeskBand
	//
	STDMETHOD(GetBandInfo)(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO *pdbi);

	// IDeskBand2
	//
	STDMETHOD(CanRenderComposited)(BOOL *pfCanRenderComposited);
	STDMETHOD(SetCompositionState)(BOOL fCompositionEnabled);
	STDMETHOD(GetCompositionState)(BOOL *pfCompositionEnabled);

	// IDockingWindow
	//
	STDMETHOD(ShowDW)(BOOL fShow);
	STDMETHOD(CloseDW)(DWORD dwReserved);
	STDMETHOD(ResizeBorderDW)(LPCRECT prcBorder, IUnknown *punkToolbarSite, BOOL fReserved);

	// IOleWindow
	//
	STDMETHOD(GetWindow)(HWND *phwnd);
	STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode);

	// IContextMenu
	//
	STDMETHOD(QueryContextMenu)(HMENU hmenu,UINT indexMenu,UINT idCmdFirst,UINT idCmdLast,UINT uFlags);
	STDMETHOD(InvokeCommand)(CMINVOKECOMMANDINFO *pici);
	STDMETHOD(GetCommandString)(UINT_PTR idCmd,UINT uType,UINT *pReserved,LPSTR pszName,UINT cchMax);


	// IPersistStreamImpl
	//
	HRESULT IPersistStreamInit_Load(
		/* [in] */ LPSTREAM pStm,
		/* [in] */ const ATL_PROPMAP_ENTRY* pMap);

	HRESULT IPersistStreamInit_Save(
		/* [in] */ LPSTREAM pStm,
		/* [in] */ BOOL fClearDirty,
		/* [in] */ const ATL_PROPMAP_ENTRY* pMap);



	CContentWindow m_wndContentWindow;

	DWORD m_dwBandID;
	DWORD m_dwViewMode;
	BOOL m_bCompositionEnabled;
	DATASETTINGDIALOG m_DataSettingDialog;

public:
	bool m_bRequiresSave; // used by IPersistStreamInitImpl

};

OBJECT_ENTRY_AUTO(__uuidof(WordTools), CWordTools)
