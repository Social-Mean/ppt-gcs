// Office hosts a native ATL ActiveX control. Managed code creates its WinForms
// view normally and attaches the resulting HWND; WinForms is not registered or
// activated as an unsupported ActiveX control.
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

using namespace ATL;

extern const CLSID CLSID_ConstraintPaneHost =
{ 0x7fa35af7, 0x2cf8, 0x4b2a, { 0x8f, 0x5d, 0x2a, 0x4b, 0xcc, 0x1e, 0x21, 0xa0 } };

class ATL_NO_VTABLE ConstraintPaneHost :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ConstraintPaneHost, &CLSID_ConstraintPaneHost>,
    public CComControl<ConstraintPaneHost>,
    public IDispatch,
    public IPersistStreamInitImpl<ConstraintPaneHost>,
    public IOleControlImpl<ConstraintPaneHost>,
    public IOleObjectImpl<ConstraintPaneHost>,
    public IOleInPlaceActiveObjectImpl<ConstraintPaneHost>,
    public IViewObjectExImpl<ConstraintPaneHost>,
    public IOleInPlaceObjectWindowlessImpl<ConstraintPaneHost>
{
public:
    ConstraintPaneHost() { m_bWindowOnly = TRUE; }
    DECLARE_NO_REGISTRY()
    DECLARE_NOT_AGGREGATABLE(ConstraintPaneHost)
    DECLARE_PROTECT_FINAL_CONSTRUCT()
    DECLARE_WND_CLASS_EX(L"PptGcsConstraintPaneHost", CS_DBLCLKS, COLOR_WINDOW)
    DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE | OLEMISC_CANTLINKINSIDE |
        OLEMISC_INSIDEOUT | OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_SETCLIENTSITEFIRST)
    DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

    BEGIN_COM_MAP(ConstraintPaneHost)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IPersistStreamInit)
        COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
        COM_INTERFACE_ENTRY(IOleControl)
        COM_INTERFACE_ENTRY(IOleObject)
        COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
        COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
        COM_INTERFACE_ENTRY2(IOleInPlaceObject, IOleInPlaceObjectWindowless)
        COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
        COM_INTERFACE_ENTRY(IViewObjectEx)
        COM_INTERFACE_ENTRY2(IViewObject2, IViewObjectEx)
        COM_INTERFACE_ENTRY2(IViewObject, IViewObjectEx)
    END_COM_MAP()

    BEGIN_PROP_MAP(ConstraintPaneHost)
    END_PROP_MAP()

    BEGIN_MSG_MAP(ConstraintPaneHost)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        MESSAGE_HANDLER(WM_SETFOCUS, OnFocus)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
        CHAIN_MSG_MAP(CComControl<ConstraintPaneHost>)
        DEFAULT_REFLECTION_HANDLER()
    END_MSG_MAP()

    HRESULT OnDraw(ATL_DRAWINFO& info)
    {
        RECT bounds = { info.prcBounds->left, info.prcBounds->top,
            info.prcBounds->right, info.prcBounds->bottom };
        ::FillRect(info.hdcDraw, &bounds, ::GetSysColorBrush(COLOR_WINDOW));
        return S_OK;
    }

    STDMETHOD(GetTypeInfoCount)(UINT* count) override
    { if (!count) return E_POINTER; *count = 0; return S_OK; }
    STDMETHOD(GetTypeInfo)(UINT, LCID, ITypeInfo**) override
    { return DISP_E_BADINDEX; }
    STDMETHOD(GetIDsOfNames)(REFIID, LPOLESTR* names, UINT count, LCID, DISPID* ids) override
    {
        if (!names || !ids) return E_POINTER;
        for (UINT index = 0; index < count; index++)
        {
            if (_wcsicmp(names[index], L"AttachWindow") == 0) ids[index] = 1;
            else if (_wcsicmp(names[index], L"DetachWindow") == 0) ids[index] = 2;
            else if (_wcsicmp(names[index], L"HostHandle") == 0) ids[index] = 3;
            else return DISP_E_UNKNOWNNAME;
        }
        return S_OK;
    }
    STDMETHOD(Invoke)(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS* parameters,
        VARIANT* result, EXCEPINFO*, UINT*) override
    {
        if (id == 1 && (flags & DISPATCH_METHOD))
        {
            if (!parameters || parameters->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            CComVariant value;
            HRESULT hr = value.ChangeType(VT_I8, &parameters->rgvarg[0]);
            if (FAILED(hr)) return DISP_E_TYPEMISMATCH;
            HWND child = reinterpret_cast<HWND>(static_cast<INT_PTR>(value.llVal));
            if (!::IsWindow(child) || ::GetWindowThreadProcessId(child, nullptr) != ::GetCurrentThreadId())
                return E_INVALIDARG;
            Detach();
            child_ = child;
            oldParent_ = ::GetParent(child);
            Attach();
            return S_OK;
        }
        if (id == 2 && (flags & DISPATCH_METHOD)) { Detach(); return S_OK; }
        if (id == 3 && (flags & DISPATCH_PROPERTYGET))
        {
            if (!result) return E_POINTER;
            ::VariantInit(result); result->vt = VT_I8;
            result->llVal = reinterpret_cast<INT_PTR>(m_hWnd);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }

    STDMETHOD(TranslateAccelerator)(LPMSG message) override
    {
        if (child_ && ::IsWindow(child_) && message &&
            message->message >= WM_KEYFIRST && message->message <= WM_KEYLAST &&
            (::GetFocus() == child_ || ::IsChild(child_, ::GetFocus())) &&
            ::IsDialogMessage(child_, message)) return S_OK;
        return S_FALSE;
    }

    void FinalRelease() { Detach(); }

private:
    HWND child_ = nullptr;
    HWND oldParent_ = nullptr;

    void Attach()
    {
        if (!m_hWnd || !child_ || !::IsWindow(child_)) return;
        ::SetParent(child_, m_hWnd);
        ResizeChild();
        ::ShowWindow(child_, SW_SHOWNA);
    }
    void Detach()
    {
        if (child_ && ::IsWindow(child_))
        {
            ::ShowWindow(child_, SW_HIDE);
            ::SetParent(child_, ::IsWindow(oldParent_) ? oldParent_ : HWND_MESSAGE);
        }
        child_ = nullptr; oldParent_ = nullptr;
    }
    void ResizeChild()
    {
        if (!child_ || !::IsWindow(child_) || !m_hWnd) return;
        RECT bounds; ::GetClientRect(m_hWnd, &bounds);
        ::MoveWindow(child_, 0, 0, bounds.right - bounds.left, bounds.bottom - bounds.top, TRUE);
    }
    LRESULT OnCreate(UINT, WPARAM, LPARAM, BOOL& handled)
    { Attach(); handled = FALSE; return 0; }
    LRESULT OnSize(UINT, WPARAM, LPARAM, BOOL& handled)
    { ResizeChild(); handled = FALSE; return 0; }
    LRESULT OnFocus(UINT, WPARAM, LPARAM, BOOL& handled)
    { if (child_ && ::IsWindow(child_)) ::SetFocus(child_); handled = FALSE; return 0; }
    LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL& handled)
    { Detach(); handled = FALSE; return 0; }
};

OBJECT_ENTRY_AUTO(CLSID_ConstraintPaneHost, ConstraintPaneHost)

class PaneModule : public CAtlDllModuleT<PaneModule> { };
PaneModule _AtlModule;

extern "C" BOOL WINAPI DllMain(HINSTANCE, DWORD reason, LPVOID reserved)
{ return _AtlModule.DllMain(reason, reserved); }
STDAPI DllCanUnloadNow() { return _AtlModule.DllCanUnloadNow(); }
STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, LPVOID* object)
{ return _AtlModule.DllGetClassObject(clsid, iid, object); }
