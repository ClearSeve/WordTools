// dllmain.h : 模块类的声明。

class CWordToolsModule : public ATL::CAtlDllModuleT< CWordToolsModule >
{
public :
	DECLARE_LIBID(LIBID_WordToolsLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WORDTOOLS, "{579F4AB8-ADE6-4E7F-87C2-F4B666E8A54B}")
};

extern class CWordToolsModule _AtlModule;
