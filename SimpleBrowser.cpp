// SimpleBrowser.cpp  –  Win32 + C++03  単一ウィンドウWebブラウザ (IWebBrowser2)
//
// 構成:
//   - URLエディットボックス + 「移動」ボタン (ツールバー)
//   - IWebBrowser2 を COM in-place activate でクライアント領域に埋め込む
//
// ビルド (MSVC コマンドライン):
//   cl SimpleBrowser.cpp /DUNICODE /D_UNICODE /link ole32.lib oleaut32.lib uuid.lib user32.lib gdi32.lib
//
// ビルド (Visual Studio):
//   空のWin32プロジェクトにこのファイルを追加してビルド

#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <objbase.h>     // CoCreateInstance (basetyps.h 経由で interface マクロを定義)
#include <exdisp.h>      // IWebBrowser2, CLSID_WebBrowser
#include <oleidl.h>      // IOleObject, IOleClientSite, IOleInPlaceSite
#include <ocidl.h>       // IOleInPlaceFrame, IOleInPlaceUIWindow
#include <oaidl.h>       // VARIANT, BSTR
#include <shlguid.h>     // IID_IWebBrowser2

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

//----------------------------------------------------------------------
// 定数
//----------------------------------------------------------------------
static const int    TOOLBAR_H  = 32;   // ツールバー (URLバー) の高さ
static const int    MARGIN     = 4;
static const int    EDIT_H     = 24;
static const int    BTN_W      = 60;
static const WCHAR  CLASS_NAME[] = L"SimpleBrowserWnd";

enum { ID_EDIT = 101, ID_GO = 102 };

//----------------------------------------------------------------------
// IOleInPlaceFrame 実装
// (IOleInPlaceSite とは IOleWindow の継承が重複するため別クラスに分離)
//----------------------------------------------------------------------
class FrameSite : public IOleInPlaceFrame
{
    LONG m_ref;
    HWND m_hwnd;
public:
    explicit FrameSite(HWND hwnd) : m_ref(1), m_hwnd(hwnd) {}

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID r, void** pp)
    {
        if (r == IID_IUnknown || r == IID_IOleWindow ||
            r == IID_IOleInPlaceUIWindow || r == IID_IOleInPlaceFrame)
        { *pp = this; AddRef(); return S_OK; }
        *pp = NULL; return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release()
    {
        ULONG r = InterlockedDecrement(&m_ref);
        if (!r) delete this;
        return r;
    }

    // IOleWindow
    STDMETHODIMP GetWindow(HWND* p)                              { *p = m_hwnd; return S_OK; }
    STDMETHODIMP ContextSensitiveHelp(BOOL)                      { return E_NOTIMPL; }

    // IOleInPlaceUIWindow
    STDMETHODIMP GetBorder(LPRECT)                               { return E_NOTIMPL; }
    STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS)             { return INPLACE_E_NOTOOLSPACE; }
    STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS)                 { return E_NOTIMPL; }
    STDMETHODIMP SetActiveObject(IOleInPlaceActiveObject*, LPCOLESTR) { return S_OK; }

    // IOleInPlaceFrame
    STDMETHODIMP InsertMenus(HMENU, LPOLEMENUGROUPWIDTHS)        { return E_NOTIMPL; }
    STDMETHODIMP SetMenu(HMENU, HOLEMENU, HWND)                  { return S_OK; }
    STDMETHODIMP RemoveMenus(HMENU)                              { return E_NOTIMPL; }
    STDMETHODIMP SetStatusText(LPCOLESTR)                        { return S_OK; }
    STDMETHODIMP EnableModeless(BOOL)                            { return S_OK; }
    STDMETHODIMP TranslateAccelerator(LPMSG, WORD)               { return E_NOTIMPL; }
};

//----------------------------------------------------------------------
// BrowserSite: IOleClientSite + IOleInPlaceSite
//----------------------------------------------------------------------
class BrowserSite : public IOleClientSite, public IOleInPlaceSite
{
    LONG       m_ref;
    HWND       m_hwnd;
    FrameSite* m_frame;
public:
    explicit BrowserSite(HWND hwnd) : m_ref(1), m_hwnd(hwnd)
    {
        m_frame = new FrameSite(hwnd);
    }
    ~BrowserSite() { if (m_frame) m_frame->Release(); }

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID r, void** pp)
    {
        if (r == IID_IUnknown || r == IID_IOleClientSite)
        { *pp = static_cast<IOleClientSite*>(this); AddRef(); return S_OK; }
        if (r == IID_IOleInPlaceSite || r == IID_IOleWindow)
        { *pp = static_cast<IOleInPlaceSite*>(this); AddRef(); return S_OK; }
        *pp = NULL; return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release()
    {
        ULONG r = InterlockedDecrement(&m_ref);
        if (!r) delete this;
        return r;
    }

    // IOleWindow (IOleInPlaceSite の基底)
    STDMETHODIMP GetWindow(HWND* p)                              { *p = m_hwnd; return S_OK; }
    STDMETHODIMP ContextSensitiveHelp(BOOL)                      { return E_NOTIMPL; }

    // IOleClientSite
    STDMETHODIMP SaveObject()                                    { return E_NOTIMPL; }
    STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** pp)         { *pp = NULL; return E_NOTIMPL; }
    STDMETHODIMP GetContainer(IOleContainer** pp)                { *pp = NULL; return E_NOINTERFACE; }
    STDMETHODIMP ShowObject()                                    { return S_OK; }
    STDMETHODIMP OnShowWindow(BOOL)                              { return S_OK; }
    STDMETHODIMP RequestNewObjectLayout()                        { return E_NOTIMPL; }

    // IOleInPlaceSite
    STDMETHODIMP CanInPlaceActivate()                            { return S_OK; }
    STDMETHODIMP OnInPlaceActivate()                             { return S_OK; }
    STDMETHODIMP OnUIActivate()                                  { return S_OK; }
    STDMETHODIMP GetWindowContext(
        IOleInPlaceFrame**    ppFrame,
        IOleInPlaceUIWindow** ppDoc,
        LPRECT                pPos,
        LPRECT                pClip,
        LPOLEINPLACEFRAMEINFO pInfo)
    {
        *ppFrame = m_frame; m_frame->AddRef();
        *ppDoc   = NULL;
        RECT rc; GetClientRect(m_hwnd, &rc);
        rc.top += TOOLBAR_H;
        *pPos = *pClip = rc;
        pInfo->fMDIApp       = FALSE;
        pInfo->hwndFrame     = m_hwnd;
        pInfo->haccel        = NULL;
        pInfo->cAccelEntries = 0;
        return S_OK;
    }
    STDMETHODIMP Scroll(SIZE)                                    { return E_NOTIMPL; }
    STDMETHODIMP OnUIDeactivate(BOOL)                            { return S_OK; }
    STDMETHODIMP OnInPlaceDeactivate()                           { return S_OK; }
    STDMETHODIMP DiscardUndoState()                              { return E_NOTIMPL; }
    STDMETHODIMP DeactivateAndUndo()                             { return E_NOTIMPL; }
    STDMETHODIMP OnPosRectChange(LPCRECT)                        { return S_OK; }
};

//----------------------------------------------------------------------
// グローバル (シングルウィンドウ)
//----------------------------------------------------------------------
static IWebBrowser2* g_browser  = NULL;
static IOleObject*   g_oleObj   = NULL;
static BrowserSite*  g_site     = NULL;
static HWND          g_hwndEdit = NULL;

//----------------------------------------------------------------------
// WebBrowser の生成・埋め込み
//----------------------------------------------------------------------
static bool EmbedBrowser(HWND hwnd)
{
    g_site = new BrowserSite(hwnd);

    RECT rc; GetClientRect(hwnd, &rc);
    rc.top += TOOLBAR_H;

    HRESULT hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER,
                                  IID_IOleObject,
                                  reinterpret_cast<void**>(&g_oleObj));
    if (FAILED(hr)) return false;

    g_oleObj->SetClientSite(static_cast<IOleClientSite*>(g_site));

    hr = g_oleObj->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL,
                          static_cast<IOleClientSite*>(g_site), 0, hwnd, &rc);
    if (FAILED(hr)) return false;

    hr = g_oleObj->QueryInterface(IID_IWebBrowser2,
                                  reinterpret_cast<void**>(&g_browser));
    return SUCCEEDED(hr) && g_browser;
}

//----------------------------------------------------------------------
// ブラウザのサイズ更新
//----------------------------------------------------------------------
static void ResizeBrowser(HWND hwnd)
{
    if (!g_oleObj) return;
    RECT rc; GetClientRect(hwnd, &rc);
    rc.top += TOOLBAR_H;

    IOleInPlaceObject* ipo = NULL;
    if (SUCCEEDED(g_oleObj->QueryInterface(IID_IOleInPlaceObject,
                                           reinterpret_cast<void**>(&ipo))))
    {
        ipo->SetObjectRects(&rc, &rc);
        ipo->Release();
    }
}

//----------------------------------------------------------------------
// URL へ移動
//----------------------------------------------------------------------
static void Navigate(HWND hwndEdit)
{
    if (!g_browser) return;
    WCHAR buf[2048] = {};
    GetWindowTextW(hwndEdit, buf, 2048);

    BSTR url = SysAllocString(buf);
    VARIANT v0, v1, v2, v3;
    VariantInit(&v0); VariantInit(&v1); VariantInit(&v2); VariantInit(&v3);
    g_browser->Navigate(url, &v0, &v1, &v2, &v3);
    SysFreeString(url);
}

//----------------------------------------------------------------------
// ウィンドウプロシージャ
//----------------------------------------------------------------------
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    static HWND s_edit = NULL;

    switch (msg)
    {
    case WM_CREATE:
    {
        RECT rc; GetClientRect(hwnd, &rc);
        int editW = rc.right - BTN_W - MARGIN * 3;

        s_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"https://www.google.com",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            MARGIN, MARGIN, editW, EDIT_H,
            hwnd, (HMENU)(INT_PTR)ID_EDIT, NULL, NULL);
        g_hwndEdit = s_edit;

        CreateWindowW(L"BUTTON", L"移動",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            MARGIN + editW + MARGIN, MARGIN, BTN_W, EDIT_H,
            hwnd, (HMENU)(INT_PTR)ID_GO, NULL, NULL);

        if (!EmbedBrowser(hwnd)) {
            MessageBoxW(hwnd, L"WebBrowserの初期化に失敗しました", L"エラー",
                        MB_OK | MB_ICONERROR);
            return -1;
        }
        Navigate(s_edit);
        return 0;
    }

    case WM_SIZE:
    {
        int w = LOWORD(lp);
        if (s_edit) {
            int editW = w - BTN_W - MARGIN * 3;
            SetWindowPos(s_edit, NULL, MARGIN, MARGIN, editW, EDIT_H, SWP_NOZORDER);
            HWND hBtn = GetDlgItem(hwnd, ID_GO);
            if (hBtn)
                SetWindowPos(hBtn, NULL,
                             MARGIN + editW + MARGIN, MARGIN, BTN_W, EDIT_H,
                             SWP_NOZORDER);
        }
        ResizeBrowser(hwnd);
        return 0;
    }

    case WM_COMMAND:
        if (LOWORD(wp) == ID_GO)
            Navigate(s_edit);
        return 0;

    case WM_DESTROY:
        if (g_browser) { g_browser->Release(); g_browser = NULL; }
        if (g_oleObj)  { g_oleObj->Close(OLECLOSE_NOSAVE); g_oleObj->Release(); g_oleObj = NULL; }
        if (g_site)    { g_site->Release(); g_site = NULL; }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

//----------------------------------------------------------------------
// エントリポイント
//----------------------------------------------------------------------
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int nShow)
{
    OleInitialize(NULL);

    WNDCLASSEXW wc = {};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = CLASS_NAME;
    RegisterClassExW(&wc);

    HWND hwnd = CreateWindowExW(0, CLASS_NAME, L"Simple Browser",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1024, 768,
        NULL, NULL, hInst, NULL);
    if (!hwnd) { OleUninitialize(); return 1; }

    ShowWindow(hwnd, nShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessageW(&msg, NULL, 0, 0))
    {
        // Enterキーでもナビゲーション実行
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_RETURN
            && msg.hwnd == g_hwndEdit)
        {
            Navigate(g_hwndEdit);
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    OleUninitialize();
    return (int)msg.wParam;
}
