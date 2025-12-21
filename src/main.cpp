#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <windowsx.h>
#include <shellapi.h>
#include <commctrl.h>
#include <commdlg.h>
#include <objidl.h>
#include <propidl.h>
#include <gdiplus.h>
#include <string>
#include <vector>
#include <algorithm>
#include "Bookmark.h"
#include "resource.h"

// Forward declaration (defined later in this file)
extern BookmarkManager* g_bookmarkManager;

// Forward declarations for helpers
std::string WideToUtf8(const std::wstring& wstr);
std::wstring Utf8ToWide(const std::string& str);

static const std::string base64_chars = 
             "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
             "abcdefghijklmnopqrstuvwxyz"
             "0123456789+/";

static std::string base64_encode(unsigned char const* bytes_to_encode, unsigned int in_len) {
  std::string ret;
  int i = 0;
  int j = 0;
  unsigned char char_array_3[3];
  unsigned char char_array_4[4];

  while (in_len--) {
    char_array_3[i++] = *(bytes_to_encode++);
    if (i == 3) {
      char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
      char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
      char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
      char_array_4[3] = char_array_3[2] & 0x3f;

      for(i = 0; (i <4) ; i++)
        ret += base64_chars[char_array_4[i]];
      i = 0;
    }
  }

  if (i) {
    for(j = i; j < 3; j++)
      char_array_3[j] = '\0';

    char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
    char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
    char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
    char_array_4[3] = char_array_3[2] & 0x3f;

    for (j = 0; (j < i + 1); j++)
      ret += base64_chars[char_array_4[j]];

    while((i++ < 3))
      ret += '=';
  }

  return ret;
}

static std::vector<unsigned char> base64_decode(std::string const& encoded_string) {
  int in_len = encoded_string.size();
  int i = 0;
  int j = 0;
  int in_ = 0;
  unsigned char char_array_4[4], char_array_3[3];
  std::vector<unsigned char> ret;

  while (in_len-- && ( encoded_string[in_] != '=') && (isalnum(encoded_string[in_]) || (encoded_string[in_] == '+') || (encoded_string[in_] == '/'))) {
    char_array_4[i++] = encoded_string[in_]; in_++;
    if (i ==4) {
      for (i = 0; i <4; i++)
        char_array_4[i] = base64_chars.find(char_array_4[i]);

      char_array_3[0] = (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
      char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
      char_array_3[2] = ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

      for (i = 0; (i < 3); i++)
        ret.push_back(char_array_3[i]);
      i = 0;
    }
  }

  if (i) {
    for (j = i; j <4; j++)
      char_array_4[j] = base64_chars.find(char_array_4[j]);

    char_array_3[0] = (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
    char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
    char_array_3[2] = ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

    for (j = 0; (j < i - 1); j++) ret.push_back(char_array_3[j]);
  }

  return ret;
}

static std::string Trim(const std::string& s) {
    size_t start = s.find_first_not_of(" \t\r\n");
    if (start == std::string::npos) return "";
    size_t end = s.find_last_not_of(" \t\r\n");
    return s.substr(start, end - start + 1);
}

static std::string GetCurrentDeviceId() {
    wchar_t buffer[MAX_PATH];
    DWORD size = MAX_PATH;
    if (GetComputerNameW(buffer, &size)) {
        return WideToUtf8(buffer);
    }
    return "Unknown";
}

static std::vector<std::string> ParseTags(const std::string& s) {
    // Comma-separated tags. Also accepts whitespace around commas.
    std::vector<std::string> tags;
    std::string cur;

    auto push = [&](std::string v) {
        v = Trim(v);
        if (v.empty()) return;
        // normalize leading '#'
        if (!v.empty() && v[0] == '#') v = Trim(v.substr(1));
        if (v.empty()) return;
        // de-dupe (case-insensitive-ish for ASCII)
        std::string lower = v;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        for (const auto& existing : tags) {
            std::string el = existing;
            std::transform(el.begin(), el.end(), el.begin(), ::tolower);
            if (el == lower) return;
        }
        tags.push_back(v);
    };

    for (char c : s) {
        if (c == ',') {
            push(cur);
            cur.clear();
        } else {
            cur.push_back(c);
        }
    }
    push(cur);
    return tags;
}

static std::string FormatTags(const std::vector<std::string>& tags) {
    std::string out;
    for (size_t i = 0; i < tags.size(); ++i) {
        if (i) out += ", ";
        out += tags[i];
    }
    return out;
}

static std::wstring FormatLocalTime(std::time_t t) {
    if (t <= 0) return L"";
    std::tm tmLocal{};
    localtime_s(&tmLocal, &t);
    wchar_t buf[64] = {};
    // yyyy-mm-dd hh:mm:ss
    wcsftime(buf, _countof(buf), L"%Y-%m-%d %H:%M:%S", &tmLocal);
    return buf;
}

static std::string NormalizeContentForDedup(const std::string& s) {
    std::string v = Trim(s);
    std::transform(v.begin(), v.end(), v.begin(), ::tolower);
    return v;
}

static int FindDuplicateIndex(const std::string& content, int excludeIndex = -1) {
    if (!g_bookmarkManager) return -1;
    std::string key = NormalizeContentForDedup(content);
    if (key.empty()) return -1;
    for (size_t i = 0; i < g_bookmarkManager->bookmarks.size(); ++i) {
        if ((int)i == excludeIndex) continue;
        if (NormalizeContentForDedup(g_bookmarkManager->bookmarks[i].content) == key) return (int)i;
    }
    return -1;
}

static LRESULT CALLBACK EditDialogChildSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR /*uIdSubclass*/, DWORD_PTR /*dwRefData*/) {
    if (msg == WM_KEYDOWN) {
        if (wParam == VK_RETURN) {
            HWND parent = GetParent(hWnd);
            if (parent) PostMessageW(parent, WM_COMMAND, MAKEWPARAM(1, BN_CLICKED), 0);
            return 0;
        }
        if (wParam == VK_ESCAPE) {
            HWND parent = GetParent(hWnd);
            if (parent) PostMessageW(parent, WM_COMMAND, MAKEWPARAM(2, BN_CLICKED), 0);
            return 0;
        }
    }
    if (msg == WM_CHAR) {
        if (wParam == VK_RETURN || wParam == VK_ESCAPE) return 0;
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

// Helper functions for Unicode conversion
std::string WideToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

std::wstring Utf8ToWide(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

// Global variables
BookmarkManager* g_bookmarkManager = nullptr;
HINSTANCE g_hInstance = nullptr;
HWND g_hSearchWnd = NULL;
HWND g_hEdit = NULL;
HWND g_hList = NULL;
HWND g_hCreateButton = NULL;
HWND g_hLogo = NULL;
HINSTANCE g_hInst = NULL;
bool g_isSearchVisible = false;
HWND g_lastForegroundWnd = NULL;
HFONT g_hUiFont = NULL;
HFONT g_hEmojiFont = NULL;

ULONG_PTR g_gdiplusToken = 0;
Gdiplus::Image* g_logoImage = nullptr;

// Constants
#define HOTKEY_ID_OPEN 1
#define HOTKEY_ID_CAPTURE 2
#define IDC_SEARCH_EDIT 101
#define IDC_RESULT_LIST 102
#define IDC_CREATE_BUTTON 103
#define IDC_LOGO 104
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_OPTIONS 1002
#define ID_TRAY_OPEN 1003

// Forward declarations
LRESULT CALLBACK MainWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK SearchWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK ListBoxProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK EditProc(HWND, UINT, WPARAM, LPARAM);
void CaptureAndBookmark();
void CreateSearchWindow();
void UpdateSearchList(const std::string& query);
void ToggleSearchWindow();
void HideSearchWindowNoRestore();
void DeleteSelectedBookmark();
void EditSelectedBookmark();
void EditBookmarkAtIndex(size_t index, const Bookmark* pDuplicateSource = nullptr);
void DuplicateSelectedBookmark();
void ExecuteSelectedBookmark();
void WaitForTargetFocus();
void PasteText(const std::string& text);
bool TryRunPowershell(const std::wstring& cmd);
void TypeText(const std::string& text);
void RestoreFocusToLastWindow();
void ShowOptionsDialog();
void InitTrayIcon(HWND hWnd);
void RemoveTrayIcon(HWND hWnd);
void AddContextMenuRegistry();
HICON GetAppIcon();
static void LoadLogoPngIfPresent();

WNDPROC g_OriginalListBoxProc;
WNDPROC g_OriginalEditProc;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    g_hInstance = hInstance;
    SetProcessDPIAware();
    g_hInst = hInstance;

    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    // GDI+ for loading/drawing PNG logo
    {
        Gdiplus::GdiplusStartupInput gdiplusStartupInput;
        Gdiplus::GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, nullptr);
    }
    
    LoadLogoPngIfPresent();
    
    // Initialize Bookmark Manager
    char path[MAX_PATH];
    GetModuleFileNameA(NULL, path, MAX_PATH);
    std::string exePath(path);
    std::string jsonPath = exePath.substr(0, exePath.find_last_of("\\/")) + "\\bookmarks.json";
    g_bookmarkManager = new BookmarkManager(jsonPath);

    // Check for command line arguments (Context Menu Mode)
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argc > 1) {
        bool binaryMode = false;
        std::string contentToBookmark;

        for (int i = 1; i < argc; ++i) {
            std::wstring arg = argv[i];
            if (arg == L"--binary" || arg == L"-b") {
                binaryMode = true;
            } else {
                contentToBookmark = WideToUtf8(arg);
            }
        }

        if (!contentToBookmark.empty()) {
            // Create a hidden window just to be the parent for the dialog
            WNDCLASSEXW wc = {0};
            wc.cbSize = sizeof(WNDCLASSEXW);
            wc.lpfnWndProc = DefWindowProc;
            wc.hInstance = hInstance;
            wc.lpszClassName = L"CheckmegTemp";
            RegisterClassExW(&wc);
            g_hSearchWnd = CreateWindowExW(0, L"CheckmegTemp", L"", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

            if (binaryMode) {
                // Read file content
                std::ifstream file(contentToBookmark, std::ios::binary);
                if (file) {
                    std::ostringstream ss;
                    ss << file.rdbuf();
                    std::string data = ss.str();
                    std::string base64 = base64_encode((unsigned char*)data.data(), data.size());
                    
                    std::string filename = contentToBookmark.substr(contentToBookmark.find_last_of("\\/") + 1);
                    
                    // Store filename in content, base64 in binaryData
                    g_bookmarkManager->addBinary(filename, base64, "application/octet-stream", GetCurrentDeviceId(), true);
                    if (!g_bookmarkManager->bookmarks.empty()) {
                        size_t idx = g_bookmarkManager->bookmarks.size() - 1;
                        // g_bookmarkManager->bookmarks[idx].tags.push_back("filename:" + filename); // No longer needed as tag if content is filename
                        g_bookmarkManager->save();
                        EditBookmarkAtIndex(idx);
                    }
                } else {
                    MessageBoxW(NULL, L"Failed to open file for binary bookmarking.", L"Error", MB_OK);
                }
            } else {
                // Add bookmark and show edit dialog
                g_bookmarkManager->add(contentToBookmark, GetCurrentDeviceId(), true);
                
                // Show edit dialog for the new bookmark
                if (!g_bookmarkManager->bookmarks.empty()) {
                    EditBookmarkAtIndex(g_bookmarkManager->bookmarks.size() - 1);
                }
            }
        }

        LocalFree(argv);
        delete g_bookmarkManager;
        if (g_gdiplusToken) Gdiplus::GdiplusShutdown(g_gdiplusToken);
        return 0;
    }
    LocalFree(argv);

    // Register Window Class for the main hidden window
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CheckmegMain";
    wc.hIcon = GetAppIcon();
    wc.hIconSm = GetAppIcon();
    RegisterClassExW(&wc);

    // Create a hidden window to handle hotkeys
    HWND hWnd = CreateWindowExW(0, L"CheckmegMain", L"Checkmeg", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    // Init Tray Icon
    InitTrayIcon(hWnd);

    // Register Hotkeys
    // WIN + ALT + X to Open Search
    if (!RegisterHotKey(hWnd, HOTKEY_ID_OPEN, MOD_WIN | MOD_ALT, 'X')) {
        MessageBoxA(NULL, "Failed to register Open Hotkey (Win+Alt+X)!", "Error", MB_OK);
    }
    
    // WIN + ALT + C to Capture
    if (!RegisterHotKey(hWnd, HOTKEY_ID_CAPTURE, MOD_WIN | MOD_ALT, 'C')) {
        MessageBoxA(NULL, "Failed to register Capture Hotkey (Win+Alt+C)!", "Error", MB_OK);
    }

    // Main message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    RemoveTrayIcon(hWnd);
    delete g_bookmarkManager;

    if (g_logoImage) {
        delete g_logoImage;
        g_logoImage = nullptr;
    }
    if (g_gdiplusToken) {
        Gdiplus::GdiplusShutdown(g_gdiplusToken);
        g_gdiplusToken = 0;
    }
    return (int) msg.wParam;
}

static std::wstring GetExeDirW() {
    wchar_t path[MAX_PATH] = {};
    GetModuleFileNameW(NULL, path, MAX_PATH);
    std::wstring exePath(path);
    size_t pos = exePath.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return L".";
    return exePath.substr(0, pos);
}

static Gdiplus::Image* LoadImageFromResource(HINSTANCE hInstance, INT resId, const WCHAR* resType) {
    HRSRC hResource = FindResourceW(hInstance, MAKEINTRESOURCEW(resId), resType);
    if (!hResource) return nullptr;

    DWORD imageSize = SizeofResource(hInstance, hResource);
    if (imageSize == 0) return nullptr;

    HGLOBAL hGlobal = LoadResource(hInstance, hResource);
    if (!hGlobal) return nullptr;

    void* pResourceData = LockResource(hGlobal);
    if (!pResourceData) return nullptr;

    HGLOBAL hBuffer = GlobalAlloc(GMEM_MOVEABLE, imageSize);
    if (!hBuffer) return nullptr;

    void* pBuffer = GlobalLock(hBuffer);
    if (!pBuffer) {
        GlobalFree(hBuffer);
        return nullptr;
    }

    CopyMemory(pBuffer, pResourceData, imageSize);
    GlobalUnlock(hBuffer);

    IStream* pStream = nullptr;
    if (CreateStreamOnHGlobal(hBuffer, TRUE, &pStream) != S_OK) {
        GlobalFree(hBuffer);
        return nullptr;
    }

    Gdiplus::Image* pImage = Gdiplus::Image::FromStream(pStream);
    pStream->Release();
    
    return pImage;
}

static void LoadLogoPngIfPresent() {
    if (!g_gdiplusToken) return;
    if (g_logoImage) {
        delete g_logoImage;
        g_logoImage = nullptr;
    }
    
    Gdiplus::Image* img = LoadImageFromResource(g_hInstance, IDB_PNG_LOGO, RT_RCDATA);
    if (!img) return;
    if (img->GetLastStatus() != Gdiplus::Ok) {
        delete img;
        return;
    }
    g_logoImage = img;
}

LRESULT CALLBACK LogoWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_LBUTTONUP:
            // Click opens options dialog
            ShowOptionsDialog();
            return 0;
        case WM_ERASEBKGND:
            return 1;
        case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            RECT rc;
            GetClientRect(hWnd, &rc);

            HBRUSH bg = GetSysColorBrush(COLOR_WINDOW);
            FillRect(hdc, &rc, bg);

            if (g_logoImage) {
                Gdiplus::Graphics g(hdc);
                g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);

                int cw = rc.right - rc.left;
                int ch = rc.bottom - rc.top;
                int pad = 2;
                int availW = std::max(1, cw - 2 * pad);
                int availH = std::max(1, ch - 2 * pad);

                int iw = (int)g_logoImage->GetWidth();
                int ih = (int)g_logoImage->GetHeight();
                if (iw > 0 && ih > 0) {
                    double scale = std::min((double)availW / (double)iw, (double)availH / (double)ih);
                    int dw = (int)(iw * scale);
                    int dh = (int)(ih * scale);
                    int dx = (cw - dw) / 2;
                    int dy = (ch - dh) / 2;
                    g.DrawImage(g_logoImage, dx, dy, dw, dh);
                }
            }

            EndPaint(hWnd, &ps);
            return 0;
        }
    }
    return DefWindowProcW(hWnd, message, wParam, lParam);
}

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_HOTKEY:
            if (wParam == HOTKEY_ID_OPEN) {
                ToggleSearchWindow();
            } else if (wParam == HOTKEY_ID_CAPTURE) {
                CaptureAndBookmark();
            }
            break;
        case WM_TRAYICON:
            if (lParam == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                SetForegroundWindow(hWnd); // Needed for track popup menu to close properly
                HMENU hMenu = CreatePopupMenu();
                AppendMenuW(hMenu, MF_STRING, ID_TRAY_OPEN, L"Open Search");
                AppendMenuW(hMenu, MF_STRING, ID_TRAY_OPTIONS, L"Options");
                AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                AppendMenuW(hMenu, MF_STRING, ID_TRAY_EXIT, L"Quit");
                int selection = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
                DestroyMenu(hMenu);
                if (selection == ID_TRAY_EXIT) {
                    PostQuitMessage(0);
                } else if (selection == ID_TRAY_OPTIONS) {
                    ShowOptionsDialog();
                } else if (selection == ID_TRAY_OPEN) {
                    ToggleSearchWindow();
                }
            } else if (lParam == WM_LBUTTONUP) {
                ToggleSearchWindow();
            }
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

HICON GetAppIcon() {
    if (g_logoImage) {
        HICON hIcon = NULL;
        Gdiplus::Bitmap* bmp = static_cast<Gdiplus::Bitmap*>(g_logoImage);
        bmp->GetHICON(&hIcon);
        return hIcon;
    }
    return LoadIcon(NULL, IDI_APPLICATION);
}

void InitTrayIcon(HWND hWnd) {
    NOTIFYICONDATAW nid = { sizeof(nid) };
    nid.hWnd = hWnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = GetAppIcon();
    wcscpy_s(nid.szTip, L"Checkmeg");
    Shell_NotifyIconW(NIM_ADD, &nid);
}

void RemoveTrayIcon(HWND hWnd) {
    NOTIFYICONDATAW nid = { sizeof(nid) };
    nid.hWnd = hWnd;
    nid.uID = 1;
    Shell_NotifyIconW(NIM_DELETE, &nid);
}

void AddContextMenuRegistry() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    
    std::wstring cmdPath = L"\"";
    cmdPath += exePath;
    cmdPath += L"\" \"%1\"";

    std::wstring cmdBinary = L"\"";
    cmdBinary += exePath;
    cmdBinary += L"\" --binary \"%1\"";

    // 1. Main Menu Item
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\*\\shell\\Checkmeg", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        RegSetValueExW(hKey, L"MUIVerb", 0, REG_SZ, (const BYTE*)L"Checkmeg", (DWORD)(wcslen(L"Checkmeg") + 1) * sizeof(wchar_t));
        RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE*)exePath, (DWORD)(wcslen(exePath) + 1) * sizeof(wchar_t));
        RegSetValueExW(hKey, L"ExtendedSubCommandsKey", 0, REG_SZ, (const BYTE*)L"Checkmeg.ContextMenu", (DWORD)(wcslen(L"Checkmeg.ContextMenu") + 1) * sizeof(wchar_t));
        
        // Remove old keys if they exist
        RegDeleteValueW(hKey, L"SubCommands");
        RegDeleteKeyW(hKey, L"command");
        RegCloseKey(hKey);
    }

    // 2. Submenu Items (using ExtendedSubCommandsKey structure in HKCU)
    // Checkmeg.ContextMenu\shell\BookmarkPath
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\Checkmeg.ContextMenu\\shell\\BookmarkPath", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        RegSetValueExW(hKey, L"MUIVerb", 0, REG_SZ, (const BYTE*)L"Bookmark Path", (DWORD)(wcslen(L"Bookmark Path") + 1) * sizeof(wchar_t));
        RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE*)exePath, (DWORD)(wcslen(exePath) + 1) * sizeof(wchar_t));
        
        HKEY hCmd;
        if (RegCreateKeyExW(hKey, L"command", 0, NULL, 0, KEY_WRITE, NULL, &hCmd, NULL) == ERROR_SUCCESS) {
            RegSetValueExW(hCmd, NULL, 0, REG_SZ, (const BYTE*)cmdPath.c_str(), (DWORD)(cmdPath.length() + 1) * sizeof(wchar_t));
            RegCloseKey(hCmd);
        }
        RegCloseKey(hKey);
    }

    // Checkmeg.ContextMenu\shell\BookmarkBinary
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\Checkmeg.ContextMenu\\shell\\BookmarkBinary", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        RegSetValueExW(hKey, L"MUIVerb", 0, REG_SZ, (const BYTE*)L"Bookmark File", (DWORD)(wcslen(L"Bookmark File") + 1) * sizeof(wchar_t));
        RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE*)exePath, (DWORD)(wcslen(exePath) + 1) * sizeof(wchar_t));
        
        HKEY hCmd;
        if (RegCreateKeyExW(hKey, L"command", 0, NULL, 0, KEY_WRITE, NULL, &hCmd, NULL) == ERROR_SUCCESS) {
            RegSetValueExW(hCmd, NULL, 0, REG_SZ, (const BYTE*)cmdBinary.c_str(), (DWORD)(cmdBinary.length() + 1) * sizeof(wchar_t));
            RegCloseKey(hCmd);
        }
        RegCloseKey(hKey);
    }

    MessageBoxW(NULL, L"Context menu items updated successfully!", L"Success", MB_OK);
}

static int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
   UINT  num = 0;          // number of image encoders
   UINT  size = 0;         // size of the image encoder array in bytes

   Gdiplus::GetImageEncodersSize(&num, &size);
   if(size == 0)
      return -1;  // Failure

   Gdiplus::ImageCodecInfo* pImageCodecInfo = (Gdiplus::ImageCodecInfo*)(malloc(size));
   if(pImageCodecInfo == NULL)
      return -1;  // Failure

   Gdiplus::GetImageEncoders(num, size, pImageCodecInfo);

   for(UINT j = 0; j < num; ++j)
   {
      if( wcscmp(pImageCodecInfo[j].MimeType, format) == 0 )
      {
         *pClsid = pImageCodecInfo[j].Clsid;
         free(pImageCodecInfo);
         return j;  // Success
      }    
   }

   free(pImageCodecInfo);
   return -1;  // Failure
}

void CaptureAndBookmark() {
    // 1. Release modifiers (Win + Alt) so they don't interfere with Ctrl+C
    // We simulate KeyUp for them.
    INPUT release[4] = {};
    int r = 0;
    
    if (GetAsyncKeyState(VK_LWIN) & 0x8000) {
        release[r].type = INPUT_KEYBOARD;
        release[r].ki.wVk = VK_LWIN;
        release[r].ki.dwFlags = KEYEVENTF_KEYUP;
        r++;
    }
    if (GetAsyncKeyState(VK_RWIN) & 0x8000) {
        release[r].type = INPUT_KEYBOARD;
        release[r].ki.wVk = VK_RWIN;
        release[r].ki.dwFlags = KEYEVENTF_KEYUP;
        r++;
    }
    if (GetAsyncKeyState(VK_MENU) & 0x8000) {
        release[r].type = INPUT_KEYBOARD;
        release[r].ki.wVk = VK_MENU;
        release[r].ki.dwFlags = KEYEVENTF_KEYUP;
        r++;
    }

    if (r > 0) {
        SendInput(r, release, sizeof(INPUT));
        Sleep(50);
    }

    // 2. Clear Clipboard to ensure we get new data
    // Retry opening clipboard as it might be locked by another app
    for (int i = 0; i < 5; ++i) {
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            CloseClipboard();
            break;
        }
        Sleep(20);
    }

    // 3. Send Ctrl+C
    INPUT inputs[4] = {};
    ZeroMemory(inputs, sizeof(inputs));

    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_CONTROL;
    
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 'C';

    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = 'C';
    inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = VK_CONTROL;
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    SendInput(4, inputs, sizeof(INPUT));

    // 4. Wait for clipboard to update with retries
    bool gotData = false;
    bool isText = false;
    bool isDib = false;
    bool isDrop = false;

    for (int i = 0; i < 20; ++i) { // Wait up to 1 second
        Sleep(50); 
        // Check formats in priority order: File -> Image -> Text
        if (IsClipboardFormatAvailable(CF_HDROP)) {
            gotData = true;
            isDrop = true;
            break;
        }
        if (IsClipboardFormatAvailable(CF_DIB)) {
            gotData = true;
            isDib = true;
            break;
        }
        if (IsClipboardFormatAvailable(CF_UNICODETEXT)) {
            gotData = true;
            isText = true;
            break;
        }
    }

    if (!gotData) {
        // Optional: Beep or visual feedback that nothing was captured
        return;
    }

    // 5. Get Data from Clipboard
    if (!OpenClipboard(NULL)) {
        MessageBoxW(NULL, L"Failed to open clipboard.", L"Error", MB_OK);
        return;
    }

    if (isText) {
        HANDLE hData = GetClipboardData(CF_UNICODETEXT);
        if (hData != NULL) {
            wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
            if (pszText != NULL) {
                std::wstring wtext(pszText);
                GlobalUnlock(hData);
                
                // 6. Save Bookmark
                std::string text = WideToUtf8(wtext);
                if (!text.empty()) {
                    // Prevent duplicates
                    g_bookmarkManager->loadIfChanged(true);
                    int dup = FindDuplicateIndex(text);
                    if (dup >= 0) {
                        int res = MessageBoxW(NULL,
                            L"Duplicate bookmark detected.\n\nOpen the existing one for editing?",
                            L"Duplicate",
                            MB_YESNO | MB_TOPMOST | MB_SETFOREGROUND);
                        if (res == IDYES) EditBookmarkAtIndex((size_t)dup);
                        CloseClipboard();
                        return;
                    }

                    g_bookmarkManager->add(text, GetCurrentDeviceId(), true);

                    // Immediately show edit dialog for the newly added bookmark
                    if (!g_bookmarkManager->bookmarks.empty()) {
                        EditBookmarkAtIndex(g_bookmarkManager->bookmarks.size() - 1);
                    }
                }
            }
        }
    } else if (isDib) {
        HANDLE hData = GetClipboardData(CF_DIB);
        if (hData != NULL) {
            void* ptr = GlobalLock(hData);
            if (ptr) {
                BITMAPINFO* bmi = (BITMAPINFO*)ptr;
                // Calculate pointer to bits.
                // Assuming packed DIB.
                BYTE* pBits = (BYTE*)ptr + bmi->bmiHeader.biSize;
                int nColors = bmi->bmiHeader.biClrUsed;
                if (nColors == 0 && bmi->bmiHeader.biBitCount <= 8) nColors = 1 << bmi->bmiHeader.biBitCount;
                if (bmi->bmiHeader.biCompression == BI_BITFIELDS) pBits += 12;
                pBits += nColors * 4;

                Gdiplus::Bitmap* bmp = new Gdiplus::Bitmap(bmi, pBits);
                if (bmp->GetLastStatus() == Gdiplus::Ok) {
                    IStream* pStream = NULL;
                    if (CreateStreamOnHGlobal(NULL, TRUE, &pStream) == S_OK) {
                        CLSID pngClsid;
                        if (GetEncoderClsid(L"image/png", &pngClsid) != -1) {
                            bmp->Save(pStream, &pngClsid, NULL);
                            
                            LARGE_INTEGER liZero = {};
                            pStream->Seek(liZero, STREAM_SEEK_SET, NULL);
                            STATSTG stg;
                            pStream->Stat(&stg, STATFLAG_NONAME);
                            DWORD size = stg.cbSize.LowPart;
                            std::vector<char> buffer(size);
                            ULONG bytesRead;
                            pStream->Read(buffer.data(), size, &bytesRead);
                            
                            std::string pngData(buffer.begin(), buffer.end());
                            std::string base64 = base64_encode((unsigned char*)pngData.data(), pngData.size());
                            
                            // Use a generic name for clipboard images
                            std::string name = "Clipboard Image.png";
                            g_bookmarkManager->addBinary(name, base64, "image/png", GetCurrentDeviceId(), true);
                            if (!g_bookmarkManager->bookmarks.empty()) {
                                EditBookmarkAtIndex(g_bookmarkManager->bookmarks.size() - 1);
                            }
                        }
                        pStream->Release();
                    }
                }
                delete bmp;
                GlobalUnlock(hData);
            }
        }
    } else if (isDrop) {
        HDROP hDrop = (HDROP)GetClipboardData(CF_HDROP);
        if (hDrop != NULL) {
            wchar_t path[MAX_PATH];
            if (DragQueryFileW(hDrop, 0, path, MAX_PATH)) {
                // Read file
                std::ifstream file(path, std::ios::binary);
                if (file) {
                    std::ostringstream ss;
                    ss << file.rdbuf();
                    std::string data = ss.str();
                    std::string base64 = base64_encode((unsigned char*)data.data(), data.size());
                    
                    std::wstring wpath(path);
                    std::string filename = WideToUtf8(wpath.substr(wpath.find_last_of(L"\\/") + 1));
                    
                    g_bookmarkManager->addBinary(filename, base64, "application/octet-stream", GetCurrentDeviceId(), true);
                    if (!g_bookmarkManager->bookmarks.empty()) {
                        size_t idx = g_bookmarkManager->bookmarks.size() - 1;
                        // g_bookmarkManager->bookmarks[idx].tags.push_back("filename:" + filename); // No longer needed
                        g_bookmarkManager->save();
                        EditBookmarkAtIndex(idx);
                    }
                }
            }
        }
    }

    CloseClipboard();
}

void ToggleSearchWindow() {
    if (g_hSearchWnd == NULL) {
        CreateSearchWindow();
    }

    if (g_isSearchVisible) {
        ShowWindow(g_hSearchWnd, SW_HIDE);
        g_isSearchVisible = false;

        // Actively restore focus to the window that was active before Checkmeg was shown.
        RestoreFocusToLastWindow();
    } else {
        // Remember what was focused before we steal focus.
        g_lastForegroundWnd = GetForegroundWindow();
        // Clear search criteria when opening.
        SetWindowTextW(g_hEdit, L"");
        UpdateSearchList(""); // Show all initially and resize/center
        ShowWindow(g_hSearchWnd, SW_SHOW);
        SetForegroundWindow(g_hSearchWnd);
        SetFocus(g_hEdit);
        g_isSearchVisible = true;
    }
}

void HideSearchWindowNoRestore() {
    if (g_isSearchVisible) {
        ShowWindow(g_hSearchWnd, SW_HIDE);
        g_isSearchVisible = false;
    }
}

LRESULT CALLBACK SearchWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_CTLCOLORSTATIC:
             // Make static controls transparent if needed, or set background color
             return (LRESULT)GetStockObject(WHITE_BRUSH);
        case WM_COMMAND:
            if (LOWORD(wParam) == IDC_SEARCH_EDIT && HIWORD(wParam) == EN_CHANGE) {
                wchar_t buffer[256];
                GetWindowTextW(g_hEdit, buffer, 256);
                UpdateSearchList(WideToUtf8(buffer));
            } else if (LOWORD(wParam) == IDC_RESULT_LIST && HIWORD(wParam) == LBN_DBLCLK) {
                ExecuteSelectedBookmark();
            } else if (LOWORD(wParam) == IDC_CREATE_BUTTON && HIWORD(wParam) == BN_CLICKED) {
                // Create a bookmark from the current query
                wchar_t buffer[2048];
                GetWindowTextW(g_hEdit, buffer, 2048);
                std::string text = WideToUtf8(buffer);
                if (!text.empty()) {
                    g_bookmarkManager->loadIfChanged(true);
                    int dup = FindDuplicateIndex(text);
                    if (dup >= 0) {
                        int res = MessageBoxW(g_hSearchWnd,
                            L"Duplicate bookmark detected.\n\nOpen the existing one for editing?",
                            L"Duplicate",
                            MB_YESNO | MB_TOPMOST | MB_SETFOREGROUND);
                        if (res == IDYES) EditBookmarkAtIndex((size_t)dup);
                        break;
                    }
                    g_bookmarkManager->add(text, GetCurrentDeviceId(), true);
                    // Immediately show edit dialog for the newly added bookmark
                    if (!g_bookmarkManager->bookmarks.empty()) {
                        EditBookmarkAtIndex(g_bookmarkManager->bookmarks.size() - 1);
                    }
                    // Refresh list after editing (or cancel)
                    UpdateSearchList(text);
                }
            } else if (LOWORD(wParam) == IDC_LOGO) {
                ShowOptionsDialog();
            }
            break;
        case WM_DRAWITEM:
            {
                const DRAWITEMSTRUCT* dis = (const DRAWITEMSTRUCT*)lParam;
                if (dis && dis->CtlID == IDC_RESULT_LIST) {
                    if (dis->itemID == (UINT)-1) break;

                    // Background
                    HBRUSH bg = (dis->itemState & ODS_SELECTED) ? GetSysColorBrush(COLOR_HIGHLIGHT) : GetSysColorBrush(COLOR_WINDOW);
                    FillRect(dis->hDC, &dis->rcItem, bg);

                    // Get bookmark index
                    int originalIdx = (int)SendMessage(dis->hwndItem, LB_GETITEMDATA, dis->itemID, 0);
                    BookmarkType type = BookmarkType::Text;
                    std::wstring text;
                    std::vector<std::wstring> tags;
                    bool isValid = true;

                    if (originalIdx >= 0 && originalIdx < (int)g_bookmarkManager->bookmarks.size()) {
                        const auto& b = g_bookmarkManager->bookmarks[originalIdx];
                        type = b.type;
                        text = Utf8ToWide(b.content);
                        tags.reserve(b.tags.size());
                        for (const auto& t : b.tags) {
                            std::wstring wt = Utf8ToWide(t);
                            if (!wt.empty()) tags.push_back(wt);
                        }
                        isValid = b.validOnAnyDevice || (b.deviceId == GetCurrentDeviceId());
                    } else {
                        // Fallback: fetch listbox string
                        wchar_t buf[2048] = {};
                        SendMessageW(dis->hwndItem, LB_GETTEXT, dis->itemID, (LPARAM)buf);
                        text = buf;
                    }

                    // Icon (emoji) + text
                    std::wstring icon;
                    if (type == BookmarkType::URL) icon = L"\U0001F310";       // 🌐
                    else if (type == BookmarkType::File) icon = L"\U0001F4C1"; // 📁
                    else if (type == BookmarkType::Command) icon = L"\U0001F5A5"; // 🖥️
                    else if (type == BookmarkType::Binary) icon = L"\U0001F4BE"; // 💾
                    else icon = L"\U0001F4C4";                                  // 📄

                    SetBkMode(dis->hDC, TRANSPARENT);
                    
                    if (!isValid) {
                        SetTextColor(dis->hDC, GetSysColor(COLOR_GRAYTEXT));
                    } else {
                        SetTextColor(dis->hDC, (dis->itemState & ODS_SELECTED) ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT));
                    }

                    int padding = 8;
                    RECT rc = dis->rcItem;
                    rc.left += padding;

                    HFONT oldFont = (HFONT)SelectObject(dis->hDC, g_hEmojiFont ? g_hEmojiFont : g_hUiFont);
                    RECT rcIcon = rc;
                    rcIcon.right = rcIcon.left + 28;
                    DrawTextW(dis->hDC, icon.c_str(), (int)icon.size(), &rcIcon, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

                    SelectObject(dis->hDC, g_hUiFont ? g_hUiFont : oldFont);
                    RECT rcText = rc;
                    rcText.left += 32;

                    // Reserve space for tags (up to half the item width)
                    int itemW = dis->rcItem.right - dis->rcItem.left;
                    int maxReserve = itemW / 2;
                    int reserved = 0;

                    auto MeasureTextWidth = [&](const std::wstring& s) -> int {
                        RECT r = {0, 0, 0, 0};
                        DrawTextW(dis->hDC, s.c_str(), (int)s.size(), &r, DT_SINGLELINE | DT_NOPREFIX | DT_CALCRECT);
                        return r.right - r.left;
                    };

                    int pillPadX = 10;
                    int pillPadY = 2;
                    int pillGap = 6;
                    int pillH = (dis->rcItem.bottom - dis->rcItem.top) - 2 * (pillPadY + 2);
                    if (pillH < 16) pillH = 16;

                    std::vector<std::wstring> drawTags;
                    int remaining = 0;
                    if (!tags.empty()) {
                        for (size_t ti = 0; ti < tags.size(); ++ti) {
                            int w = MeasureTextWidth(tags[ti]) + 2 * pillPadX;
                            int next = (drawTags.empty() ? 0 : pillGap) + w;
                            if (reserved + next > maxReserve) {
                                remaining = (int)(tags.size() - ti);
                                break;
                            }
                            reserved += next;
                            drawTags.push_back(tags[ti]);
                        }
                        if (remaining > 0) {
                            std::wstring more = L"+" + std::to_wstring(remaining);
                            int w = MeasureTextWidth(more) + 2 * pillPadX;
                            int next = (drawTags.empty() ? 0 : pillGap) + w;
                            if (reserved + next <= maxReserve) {
                                reserved += next;
                                drawTags.push_back(more);
                            }
                        }
                    }

                    RECT rcTextAvail = rcText;
                    if (reserved > 0) {
                        rcTextAvail.right = std::max(rcTextAvail.left, rcTextAvail.right - reserved - pillGap);
                    }
                    DrawTextW(dis->hDC, text.c_str(), (int)text.size(), &rcTextAvail, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

                    // Draw tag pills to the right
                    if (!drawTags.empty()) {
                        int x = rcTextAvail.right + pillGap;
                        int yMid = (dis->rcItem.top + dis->rcItem.bottom) / 2;
                        int yTop = yMid - pillH / 2;
                        int yBot = yTop + pillH;

                        COLORREF fill = GetSysColor((dis->itemState & ODS_SELECTED) ? COLOR_HIGHLIGHT : COLOR_BTNFACE);
                        COLORREF border = GetSysColor((dis->itemState & ODS_SELECTED) ? COLOR_HIGHLIGHTTEXT : COLOR_3DSHADOW);
                        COLORREF txt = GetSysColor((dis->itemState & ODS_SELECTED) ? COLOR_HIGHLIGHTTEXT : COLOR_BTNTEXT);
                        
                        if (!isValid) {
                            // Dim pills if invalid
                            fill = GetSysColor(COLOR_BTNFACE); // Keep gray
                            border = GetSysColor(COLOR_GRAYTEXT);
                            txt = GetSysColor(COLOR_GRAYTEXT);
                        }

                        HPEN hPen = CreatePen(PS_SOLID, 1, border);
                        HBRUSH hBrush = CreateSolidBrush(fill);
                        HPEN oldPen = (HPEN)SelectObject(dis->hDC, hPen);
                        HBRUSH oldBrush = (HBRUSH)SelectObject(dis->hDC, hBrush);
                        COLORREF oldTxt = SetTextColor(dis->hDC, txt);

                        for (const auto& tag : drawTags) {
                            int w = MeasureTextWidth(tag) + 2 * pillPadX;
                            RECT prc = { x, yTop, x + w, yBot };
                            int radius = pillH / 2;
                            RoundRect(dis->hDC, prc.left, prc.top, prc.right, prc.bottom, radius, radius);

                            RECT trc = prc;
                            trc.left += pillPadX;
                            trc.right -= pillPadX;
                            DrawTextW(dis->hDC, tag.c_str(), (int)tag.size(), &trc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

                            x += w + pillGap;
                            if (x >= dis->rcItem.right - padding) break;
                        }

                        SetTextColor(dis->hDC, oldTxt);
                        SelectObject(dis->hDC, oldBrush);
                        SelectObject(dis->hDC, oldPen);
                        DeleteObject(hBrush);
                        DeleteObject(hPen);
                    }

                    SelectObject(dis->hDC, oldFont);

                    // Focus rect
                    if (dis->itemState & ODS_FOCUS) {
                        DrawFocusRect(dis->hDC, &dis->rcItem);
                    }
                    return TRUE;
                }
            }
            break;
        case WM_MEASUREITEM:
            {
                MEASUREITEMSTRUCT* mis = (MEASUREITEMSTRUCT*)lParam;
                if (mis && mis->CtlID == IDC_RESULT_LIST) {
                    // Fixed height based on UI font metrics
                    HDC hdc = GetDC(g_hSearchWnd);
                    HFONT old = (HFONT)SelectObject(hdc, g_hUiFont);
                    TEXTMETRIC tm;
                    GetTextMetrics(hdc, &tm);
                    SelectObject(hdc, old);
                    ReleaseDC(g_hSearchWnd, hdc);

                    mis->itemHeight = (UINT)(tm.tmHeight + 8);
                    return TRUE;
                }
            }
            break;
        case WM_KEYDOWN:
            if (wParam == VK_ESCAPE) {
                HideSearchWindowNoRestore();
            }
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void CreateSearchWindow() {
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = SearchWndProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"CheckmegSearch";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = GetAppIcon();
    wc.hIconSm = GetAppIcon();
    RegisterClassExW(&wc);

    g_hSearchWnd = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST, // Toolwindow to hide from taskbar
        L"CheckmegSearch", 
        L"", 
        WS_POPUP | WS_BORDER, // Popup for frameless
        0, 0, 600, 400, 
        NULL, NULL, g_hInst, NULL
    );

    // Register & create logo control
    {
        WNDCLASSEXW lwc = {0};
        lwc.cbSize = sizeof(WNDCLASSEXW);
        lwc.lpfnWndProc = LogoWndProc;
        lwc.hInstance = g_hInst;
        lwc.lpszClassName = L"CheckmegLogo";
        lwc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        RegisterClassExW(&lwc);

        g_hLogo = CreateWindowExW(0, L"CheckmegLogo", L"",
            WS_CHILD | WS_VISIBLE,
            10, 10, 25, 25,
            g_hSearchWnd, (HMENU)IDC_LOGO, g_hInst, NULL);
    }

    // Create Edit Control
    g_hEdit = CreateWindowExW(
        WS_EX_CLIENTEDGE, L"EDIT", L"", 
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 
        10, 10, 580, 25, 
        g_hSearchWnd, (HMENU)IDC_SEARCH_EDIT, g_hInst, NULL
    );
    
    // Set font for edit
    HDC hdc = GetDC(g_hSearchWnd);
    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(g_hSearchWnd, hdc);
    int fontSize = -MulDiv(11, dpiY, 72); // 11pt font

    if (g_hUiFont == NULL) {
        // Prefer Fira Code; fall back to Segoe UI if not installed.
        g_hUiFont = CreateFontW(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Fira Code");
        if (g_hUiFont == NULL) {
            g_hUiFont = CreateFontW(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        }
    }
    if (g_hEmojiFont == NULL) {
        // Emoji font for reliable rendering of emoji glyphs.
        g_hEmojiFont = CreateFontW(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI Emoji");
    }
    SendMessage(g_hEdit, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

    // Subclass Edit Control
    g_OriginalEditProc = (WNDPROC)SetWindowLongPtr(g_hEdit, GWLP_WNDPROC, (LONG_PTR)EditProc);

    // Create ListBox
    g_hList = CreateWindowExW(
        WS_EX_CLIENTEDGE, L"LISTBOX", L"", 
        WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS, 
        10, 45, 580, 345, 
        g_hSearchWnd, (HMENU)IDC_RESULT_LIST, g_hInst, NULL
    );
    SendMessage(g_hList, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

    // Create "Create bookmark" button (hidden by default)
    g_hCreateButton = CreateWindowExW(
        0, L"BUTTON", L"CREATE",
        WS_CHILD | BS_DEFPUSHBUTTON,
        10, 45, 580, 40,
        g_hSearchWnd, (HMENU)IDC_CREATE_BUTTON, g_hInst, NULL
    );
    SendMessage(g_hCreateButton, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
    ShowWindow(g_hCreateButton, SW_HIDE);
    
    // Subclass ListBox to handle Delete key and Context Menu
    g_OriginalListBoxProc = (WNDPROC)SetWindowLongPtr(g_hList, GWLP_WNDPROC, (LONG_PTR)ListBoxProc);
}

LRESULT CALLBACK EditProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_CHAR:
            // Swallow Enter/Esc to avoid system chime from single-line edit.
            if (wParam == VK_RETURN || wParam == VK_ESCAPE) return 0;
            // Swallow Ctrl+Backspace char (0x7F)
            if (wParam == 0x7F) return 0;
            // Ctrl+A (Select All)
            if (wParam == 1) {
                SendMessage(hWnd, EM_SETSEL, 0, -1);
                return 0;
            }
            break;
        case WM_SYSKEYDOWN:
            if (wParam == 'D' || wParam == 'd') {
                SetFocus(g_hEdit);
                SendMessageW(g_hEdit, EM_SETSEL, 0, -1);
                return 0;
            }
            break;
        case WM_KEYDOWN:
            if (wParam == VK_DOWN) {
                // Move focus to list
                SetFocus(g_hList);
                int cur = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
                if (cur == LB_ERR) {
                    SendMessage(g_hList, LB_SETCURSEL, 0, 0);
                }
                return 0;
            } else if (wParam == VK_UP) {
                // Keep focus in edit, but maybe do nothing or move caret?
                // Default edit behavior is fine.
                return 0;
            } else if (wParam == VK_BACK && (GetKeyState(VK_CONTROL) & 0x8000)) {
                // Ctrl+Backspace: Delete word
                DWORD start, end;
                SendMessage(hWnd, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
                if (start != end) {
                    SendMessage(hWnd, WM_CLEAR, 0, 0);
                } else if (start > 0) {
                    // Find start of previous word
                    wchar_t buffer[1024]; // Should be enough for context
                    // Get text before cursor
                    // This is complex to do perfectly without getting full text.
                    // Let's just get full text.
                    int len = GetWindowTextLengthW(hWnd);
                    std::vector<wchar_t> buf(len + 1);
                    GetWindowTextW(hWnd, &buf[0], len + 1);
                    
                    int i = start - 1;
                    // Skip whitespace
                    while (i >= 0 && iswspace(buf[i])) i--;
                    // Skip non-whitespace
                    while (i >= 0 && !iswspace(buf[i])) i--;
                    
                    SendMessage(hWnd, EM_SETSEL, i + 1, start);
                    SendMessage(hWnd, WM_CLEAR, 0, 0);
                }
                return 0;
            } else if ((wParam == 'E' || wParam == 'e') && (GetKeyState(VK_CONTROL) & 0x8000)) {
                // Allow editing even when focus stays in the search box.
                EditSelectedBookmark();
                return 0;
            } else if ((wParam == 'D' || wParam == 'd') && (GetKeyState(VK_CONTROL) & 0x8000)) {
                DuplicateSelectedBookmark();
                return 0;
            } else if (wParam == VK_RETURN) {
                if (g_hCreateButton && IsWindowVisible(g_hCreateButton)) {
                    SendMessageW(g_hSearchWnd, WM_COMMAND, MAKEWPARAM(IDC_CREATE_BUTTON, BN_CLICKED), (LPARAM)g_hCreateButton);
                } else {
                    ExecuteSelectedBookmark();
                }
                return 0;
            } else if (wParam == VK_ESCAPE) {
                HideSearchWindowNoRestore();
                return 0;
            }
            break;
    }
    return CallWindowProc(g_OriginalEditProc, hWnd, message, wParam, lParam);
}

LRESULT CALLBACK ListBoxProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_CHAR:
            if (wParam == VK_RETURN || wParam == VK_ESCAPE) return 0;
            // Forward typing to Edit
            SetFocus(g_hEdit);
            SendMessage(g_hEdit, WM_CHAR, wParam, lParam);
            {
                int len = GetWindowTextLength(g_hEdit);
                SendMessage(g_hEdit, EM_SETSEL, len, len);
            }
            return 0;
        case WM_KEYDOWN:
            if (wParam == VK_BACK) {
                // Forward Backspace to Edit
                SetFocus(g_hEdit);
                SendMessage(g_hEdit, WM_KEYDOWN, VK_BACK, lParam);
                {
                    int len = GetWindowTextLength(g_hEdit);
                    SendMessage(g_hEdit, EM_SETSEL, len, len);
                }
                return 0;
            }
            if (wParam == VK_DELETE) {
                int idx = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
                if (idx != LB_ERR) {
                    int originalIdx = (int)SendMessage(g_hList, LB_GETITEMDATA, idx, 0);
                    std::wstring preview;
                    if (originalIdx >= 0 && originalIdx < (int)g_bookmarkManager->bookmarks.size()) {
                        preview = Utf8ToWide(g_bookmarkManager->bookmarks[originalIdx].content);
                    }
                    std::wstring msg = L"Delete this bookmark?";
                    if (!preview.empty()) {
                        if (preview.size() > 140) preview = preview.substr(0, 140) + L"…";
                        msg += L"\n\n" + preview;
                    }
                    int res = MessageBoxW(g_hSearchWnd, msg.c_str(), L"Confirm Delete", MB_YESNO | MB_TOPMOST | MB_SETFOREGROUND);
                    if (res == IDYES) {
                        DeleteSelectedBookmark();
                    }
                }
                return 0;
            } else if ((wParam == 'E' || wParam == 'e') && (GetKeyState(VK_CONTROL) & 0x8000)) {
                EditSelectedBookmark();
                return 0;
            } else if ((wParam == 'D' || wParam == 'd') && (GetKeyState(VK_CONTROL) & 0x8000)) {
                DuplicateSelectedBookmark();
                return 0;
            } else if (wParam == VK_RETURN) {
                ExecuteSelectedBookmark();
                return 0;
            } else if (wParam == VK_ESCAPE) {
                HideSearchWindowNoRestore();
                return 0;
            }
            break;
        case WM_CONTEXTMENU:
            {
                // Ensure right-click also selects the item under cursor.
                POINT screenPt;
                screenPt.x = GET_X_LPARAM(lParam);
                screenPt.y = GET_Y_LPARAM(lParam);

                if (screenPt.x == -1 && screenPt.y == -1) {
                    // Invoked via keyboard/context key; use current cursor position.
                    GetCursorPos(&screenPt);
                }

                POINT clientPt = screenPt;
                ScreenToClient(hWnd, &clientPt);
                LRESULT hit = SendMessageW(hWnd, LB_ITEMFROMPOINT, 0, MAKELPARAM(clientPt.x, clientPt.y));
                int hitIndex = LOWORD(hit);
                BOOL isOutside = HIWORD(hit);
                if (!isOutside && hitIndex != LB_ERR) {
                    SendMessageW(hWnd, LB_SETCURSEL, hitIndex, 0);
                }

                SetFocus(hWnd);

                HMENU hMenu = CreatePopupMenu();
                AppendMenuW(hMenu, MF_STRING, 1, L"Edit");
                AppendMenuW(hMenu, MF_STRING, 3, L"Duplicate");
                AppendMenuW(hMenu, MF_STRING, 2, L"Delete");
                
                POINT pt;
                GetCursorPos(&pt);
                int selection = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_hSearchWnd, NULL);
                DestroyMenu(hMenu);

                if (selection == 1) EditSelectedBookmark();
                else if (selection == 2) DeleteSelectedBookmark();
                else if (selection == 3) DuplicateSelectedBookmark();
            }
            return 0;
    }
    return CallWindowProc(g_OriginalListBoxProc, hWnd, message, wParam, lParam);
}

void PasteText(const std::string& text) {
    auto IsCheckmegWindow = [](HWND hWnd) -> bool {
        return (hWnd == NULL) || (hWnd == g_hSearchWnd) || (hWnd == g_hEdit) || (hWnd == g_hList) || (hWnd == g_hCreateButton);
    };

    auto RestoreFocusToLast = [&]() {
        if (g_lastForegroundWnd == NULL || IsCheckmegWindow(g_lastForegroundWnd) || !IsWindow(g_lastForegroundWnd)) {
            return;
        }

        DWORD targetThreadId = GetWindowThreadProcessId(g_lastForegroundWnd, NULL);
        DWORD myThreadId = GetCurrentThreadId();

        // Attach so SetForegroundWindow/SetFocus have a better chance of working.
        if (targetThreadId != 0 && targetThreadId != myThreadId) {
            AttachThreadInput(myThreadId, targetThreadId, TRUE);
        }

        SetForegroundWindow(g_lastForegroundWnd);
        SetActiveWindow(g_lastForegroundWnd);
        BringWindowToTop(g_lastForegroundWnd);

        if (targetThreadId != 0 && targetThreadId != myThreadId) {
            AttachThreadInput(myThreadId, targetThreadId, FALSE);
        }
    };

    auto ReleaseModifiers = [&]() {
        // Ensure Win/Alt/Shift/Ctrl aren't stuck down when we synthesize paste.
        const WORD keys[] = { VK_LWIN, VK_RWIN, VK_MENU, VK_SHIFT, VK_CONTROL };
        INPUT ups[sizeof(keys) / sizeof(keys[0])] = {};
        int n = 0;
        for (WORD vk : keys) {
            if (GetAsyncKeyState(vk) & 0x8000) {
                ups[n].type = INPUT_KEYBOARD;
                ups[n].ki.wVk = vk;
                ups[n].ki.dwFlags = KEYEVENTF_KEYUP;
                n++;
            }
        }
        if (n > 0) {
            SendInput(n, ups, sizeof(INPUT));
            Sleep(10);
        }
    };

    RestoreFocusToLast();

    // 1. Set Clipboard
    // Retry opening clipboard
    for (int i = 0; i < 5; ++i) {
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            std::wstring wText = Utf8ToWide(text);
            HGLOBAL hGlob = GlobalAlloc(GMEM_MOVEABLE, (wText.length() + 1) * sizeof(wchar_t));
            if (hGlob) {
                void* locked = GlobalLock(hGlob);
                if (locked) {
                    memcpy(locked, wText.c_str(), (wText.length() + 1) * sizeof(wchar_t));
                    GlobalUnlock(hGlob);
                    SetClipboardData(CF_UNICODETEXT, hGlob);
                } else {
                    GlobalFree(hGlob);
                }
            }
            CloseClipboard();
            break;
        }
        Sleep(20);
    }

    // Prefer direct typing because it works reliably in browsers (Firefox address bar)
    // and simple apps (Notepad), where focus/foreground timing can make Ctrl+V flaky.
    RestoreFocusToLast();
    ReleaseModifiers();
    TypeText(text);
}

bool TryRunPowershell(const std::wstring& cmd) {
    // Construct command: powershell -NoProfile -WindowStyle Hidden -Command "& { try { <cmd> } catch { exit 1 } }"
    std::wstring psCmd = L"powershell.exe -NoProfile -WindowStyle Hidden -Command \"& { try { " + cmd + L" } catch { exit 1 } }\"";
    
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    
    std::vector<wchar_t> buf(psCmd.begin(), psCmd.end());
    buf.push_back(0);
    
    if (CreateProcessW(NULL, &buf[0], NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        // Wait up to 500ms
        DWORD waitResult = WaitForSingleObject(pi.hProcess, 500);
        DWORD exitCode = 0;
        bool success = true;
        
        if (waitResult == WAIT_OBJECT_0) {
            GetExitCodeProcess(pi.hProcess, &exitCode);
            if (exitCode != 0) success = false;
        } else {
            // Timeout - assume it's running successfully (long running or GUI)
            success = true;
        }
        
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        return success;
    }
    return false;
}

void ExecuteSelectedBookmark() {
    int idx = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
    if (idx != LB_ERR) {
        int originalIdx = (int)SendMessage(g_hList, LB_GETITEMDATA, idx, 0);
        if (originalIdx >= 0 && originalIdx < g_bookmarkManager->bookmarks.size()) {
            const auto& b = g_bookmarkManager->bookmarks[originalIdx];
            
            // Check validity
            bool isValid = b.validOnAnyDevice || (b.deviceId == GetCurrentDeviceId());
            if (!isValid) return;

            // Update last used time
            g_bookmarkManager->updateLastUsed(originalIdx);

            std::wstring wContent = Utf8ToWide(b.content);
            
            ToggleSearchWindow(); // Close first to restore focus
            WaitForTargetFocus();

            // Behavior is type-driven:
            // - Binary: prompt to save
            // - URL/File: open
            // - Command: run PowerShell
            // - Text: type into target (never auto-exec)
            if (b.type == BookmarkType::Binary) {
                g_bookmarkManager->ensureBinaryDataLoaded(originalIdx);
                OPENFILENAMEW ofn = { sizeof(ofn) };
                wchar_t szFile[MAX_PATH] = { 0 };
                
                std::string ext = ".bin";
                if (b.mimeType == "image/png") ext = ".png";
                else if (b.mimeType == "image/bmp") ext = ".bmp";
                else if (b.mimeType == "image/jpeg") ext = ".jpg";
                
                // Use content as filename if it looks like one, otherwise fallback
                if (!b.content.empty() && b.content.length() < 250 && b.content.find("base64") == std::string::npos) {
                     wcscpy_s(szFile, Utf8ToWide(b.content).c_str());
                } else {
                    wcscpy_s(szFile, L"bookmark");
                    wcscat_s(szFile, Utf8ToWide(ext).c_str());
                }

                ofn.lpstrFile = szFile;
                ofn.nMaxFile = MAX_PATH;
                ofn.lpstrTitle = L"Save Binary Bookmark";
                ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
                
                std::wstring filter = L"All Files (*.*)\0*.*\0";
                if (ext == ".png") filter = L"PNG Image (*.png)\0*.png\0All Files (*.*)\0*.*\0";
                ofn.lpstrFilter = filter.c_str();

                if (GetSaveFileNameW(&ofn)) {
                    // Use binaryData if available, otherwise fallback to content (legacy)
                    std::string dataStr = b.binaryData;
                    if (dataStr.empty()) dataStr = b.content;

                    std::vector<unsigned char> data = base64_decode(dataStr);
                    std::ofstream outfile(ofn.lpstrFile, std::ios::binary);
                    outfile.write((const char*)data.data(), data.size());
                    outfile.close();
                }
            } else if (b.type == BookmarkType::URL || b.type == BookmarkType::File) {
                HINSTANCE hInst = ShellExecuteW(NULL, L"open", wContent.c_str(), NULL, NULL, SW_SHOWNORMAL);
                if ((INT_PTR)hInst <= 32) {
                    WaitForTargetFocus();
                    PasteText(b.content);
                }
            } else if (b.type == BookmarkType::Command) {
                if (!TryRunPowershell(wContent)) {
                    // If it fails, fall back to typing it as text.
                    WaitForTargetFocus();
                    PasteText(b.content);
                }
            } else {
                WaitForTargetFocus();
                PasteText(b.content);
            }
        }
    }
}

void RestoreFocusToLastWindow() {
    if (g_lastForegroundWnd == NULL || g_lastForegroundWnd == g_hSearchWnd || !IsWindow(g_lastForegroundWnd)) {
        return;
    }

    DWORD targetThreadId = GetWindowThreadProcessId(g_lastForegroundWnd, NULL);
    DWORD myThreadId = GetCurrentThreadId();

    if (targetThreadId != 0 && targetThreadId != myThreadId) {
        AttachThreadInput(myThreadId, targetThreadId, TRUE);
    }

    ShowWindowAsync(g_lastForegroundWnd, SW_SHOW);
    SetForegroundWindow(g_lastForegroundWnd);
    SetActiveWindow(g_lastForegroundWnd);
    BringWindowToTop(g_lastForegroundWnd);

    if (targetThreadId != 0 && targetThreadId != myThreadId) {
        AttachThreadInput(myThreadId, targetThreadId, FALSE);
    }
}

static bool IsContextMenuRegistered() {
    HKEY hKey;
    LONG lRes = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\*\\shell\\Checkmeg", 0, KEY_READ, &hKey);
    if (lRes == ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return true;
    }
    return false;
}

static void RemoveContextMenuRegistry() {
    RegDeleteTreeW(HKEY_CURRENT_USER, L"Software\\Classes\\*\\shell\\Checkmeg");
    RegDeleteTreeW(HKEY_CURRENT_USER, L"Software\\Classes\\Checkmeg.ContextMenu");
}

static bool IsRunAtStartup() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        wchar_t path[MAX_PATH];
        DWORD size = sizeof(path);
        DWORD type = 0;
        if (RegQueryValueExW(hKey, L"Checkmeg", NULL, &type, (LPBYTE)path, &size) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            return true;
        }
        RegCloseKey(hKey);
    }
    return false;
}

static void SetRunAtStartup(bool enable) {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            // Quote path to be safe
            std::wstring cmd = L"\"";
            cmd += exePath;
            cmd += L"\"";
            RegSetValueExW(hKey, L"Checkmeg", 0, REG_SZ, (const BYTE*)cmd.c_str(), (DWORD)(cmd.length() + 1) * sizeof(wchar_t));
        } else {
            RegDeleteValueW(hKey, L"Checkmeg");
        }
        RegCloseKey(hKey);
    }
}

void ShowOptionsDialog() {
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"CheckmegOptions";
    wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
    wc.hIcon = GetAppIcon();
    wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT {
        static HFONT hTitleFont = NULL;
        static HFONT hVersionFont = NULL;
        static HFONT hSectionFont = NULL;
        static HWND hBtnCtx = NULL;
        static HWND hBtnRun = NULL;
        static HWND hBtnOpenJson = NULL;
        static HWND hBtnOk = NULL;
        static bool isRegistered = false;
        static bool isRunAtStartup = false;

        switch (msg) {
            case WM_CREATE: {
                HDC hdc = GetDC(hWnd);
                int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                ReleaseDC(hWnd, hdc);
                float scale = dpiY / 96.0f;

                // Fonts
                hTitleFont = CreateFontW((int)(-24 * scale), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
                hVersionFont = CreateFontW((int)(-14 * scale), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
                hSectionFont = CreateFontW((int)(-16 * scale), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

                isRegistered = IsContextMenuRegistered();
                isRunAtStartup = IsRunAtStartup();

                // Run at Startup Checkbox
                hBtnRun = CreateWindowExW(0, L"BUTTON", L"Run at Startup", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    (int)(40 * scale), (int)(130 * scale), (int)(300 * scale), (int)(25 * scale), hWnd, (HMENU)3, g_hInst, NULL);
                SendMessageW(hBtnRun, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SendMessageW(hBtnRun, BM_SETCHECK, isRunAtStartup ? BST_CHECKED : BST_UNCHECKED, 0);

                // Context Menu Button
                const wchar_t* btnText = isRegistered ? L"Remove from File Context Menu" : L"Add to File Context Menu";
                hBtnCtx = CreateWindowExW(0, L"BUTTON", btnText, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(40 * scale), (int)(165 * scale), (int)(300 * scale), (int)(30 * scale), hWnd, (HMENU)2, g_hInst, NULL);
                SendMessageW(hBtnCtx, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                // Export Bookmarks File Button
                hBtnOpenJson = CreateWindowExW(0, L"BUTTON", L"Export Bookmarks File", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(40 * scale), (int)(205 * scale), (int)(300 * scale), (int)(30 * scale), hWnd, (HMENU)4, g_hInst, NULL);
                SendMessageW(hBtnOpenJson, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                hBtnOk = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                    (int)(280 * scale), (int)(250 * scale), (int)(80 * scale), (int)(30 * scale), hWnd, (HMENU)1, g_hInst, NULL);
                SendMessageW(hBtnOk, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                SetFocus(hBtnOk);
                return 0;
            }
            case WM_PAINT: {
                PAINTSTRUCT ps;
                HDC hdc = BeginPaint(hWnd, &ps);
                
                int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                float scale = dpiY / 96.0f;

                // Draw Logo
                if (g_logoImage) {
                    Gdiplus::Graphics g(hdc);
                    g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
                    g.DrawImage(g_logoImage, (int)(20 * scale), (int)(20 * scale), (int)(48 * scale), (int)(48 * scale));
                }

                // Draw Title
                SetBkMode(hdc, TRANSPARENT);
                HFONT oldFont = (HFONT)SelectObject(hdc, hTitleFont);
                SetTextColor(hdc, RGB(33, 33, 33));
                TextOutW(hdc, (int)(80 * scale), (int)(20 * scale), L"Checkmeg", 8);
                
                // Draw Version
                SelectObject(hdc, hVersionFont);
                SetTextColor(hdc, RGB(100, 100, 100));
                TextOutW(hdc, (int)(80 * scale), (int)(55 * scale), L"v1.0", 4);

                // Separator
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(220, 220, 220));
                HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, (int)(20 * scale), (int)(90 * scale), NULL);
                LineTo(hdc, (int)(380 * scale), (int)(90 * scale));
                SelectObject(hdc, oldPen);
                DeleteObject(hPen);

                // Section Header
                SelectObject(hdc, hSectionFont);
                SetTextColor(hdc, RGB(0, 0, 0));
                TextOutW(hdc, (int)(20 * scale), (int)(105 * scale), L"System Integration", 18);

                SelectObject(hdc, oldFont);
                EndPaint(hWnd, &ps);
                return 0;
            }
            case WM_COMMAND: {
                int id = LOWORD(wParam);
                if (id == 1) { // OK
                    DestroyWindow(hWnd);
                } else if (id == 2) { // Context Menu Toggle
                    if (isRegistered) {
                        RemoveContextMenuRegistry();
                        isRegistered = false;
                        SetWindowTextW(hBtnCtx, L"Add to File Context Menu");
                        MessageBoxW(hWnd, L"Removed from File Context Menu.", L"Success", MB_OK | MB_ICONINFORMATION);
                    } else {
                        AddContextMenuRegistry();
                        isRegistered = true;
                        SetWindowTextW(hBtnCtx, L"Remove from File Context Menu");
                        MessageBoxW(hWnd, L"Added to File Context Menu.", L"Success", MB_OK | MB_ICONINFORMATION);
                    }
                } else if (id == 3) { // Run at Startup Checkbox
                    if (HIWORD(wParam) == BN_CLICKED) {
                        bool check = (SendMessageW(hBtnRun, BM_GETCHECK, 0, 0) == BST_CHECKED);
                        SetRunAtStartup(check);
                    }
                } else if (id == 4) { // Export Bookmarks File
                    if (g_bookmarkManager) {
                        wchar_t szFile[MAX_PATH] = L"bookmarks.json";
                        OPENFILENAMEW ofn = {0};
                        ofn.lStructSize = sizeof(ofn);
                        ofn.hwndOwner = hWnd;
                        ofn.lpstrFile = szFile;
                        ofn.nMaxFile = sizeof(szFile);
                        ofn.lpstrFilter = L"JSON Files\0*.json\0All Files\0*.*\0";
                        ofn.nFilterIndex = 1;
                        ofn.lpstrFileTitle = NULL;
                        ofn.nMaxFileTitle = 0;
                        ofn.lpstrInitialDir = NULL;
                        ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

                        if (GetSaveFileNameW(&ofn) == TRUE) {
                            std::wstring src = Utf8ToWide(g_bookmarkManager->filePath);
                            if (CopyFileW(src.c_str(), ofn.lpstrFile, FALSE)) {
                                MessageBoxW(hWnd, L"Bookmarks exported successfully.", L"Success", MB_OK | MB_ICONINFORMATION);
                            } else {
                                MessageBoxW(hWnd, L"Failed to export bookmarks.", L"Error", MB_OK | MB_ICONERROR);
                            }
                        }
                    }
                }
                return 0;
            }
            case WM_DESTROY:
                if (hTitleFont) DeleteObject(hTitleFont);
                if (hVersionFont) DeleteObject(hVersionFont);
                if (hSectionFont) DeleteObject(hSectionFont);
                return 0;
            case WM_CLOSE:
                DestroyWindow(hWnd);
                return 0;
        }
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    };

    RegisterClassExW(&wc);

    HDC hdc = GetDC(NULL);
    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(NULL, hdc);
    float scale = dpiY / 96.0f;
    
    int w = (int)(400 * scale);
    int h = (int)(330 * scale);
    int x = GetSystemMetrics(SM_CXSCREEN) / 2 - w / 2;
    int y = GetSystemMetrics(SM_CYSCREEN) / 2 - h / 2;

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"CheckmegOptions", L"Options",
        WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU,
        x, y, w, h,
        g_hSearchWnd, NULL, g_hInst, NULL);

    MSG msg;
    BOOL bRet;
    while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0) {
        if (bRet == -1) break;
        if (!IsWindow(hDlg)) break;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnregisterClassW(L"CheckmegOptions", g_hInst);
}

void TypeText(const std::string& text) {
    // Types text into the currently focused control using Unicode input.
    // This is generally more reliable than Ctrl+V for browsers/address bars.
    std::wstring w = Utf8ToWide(text);
    if (w.empty()) return;

    // Build inputs: key down + key up per character.
    std::vector<INPUT> inputs;
    inputs.reserve(w.size() * 2);

    for (wchar_t ch : w) {
        if (ch == L'\r') continue;
        if (ch == L'\n') {
            INPUT down = {};
            down.type = INPUT_KEYBOARD;
            down.ki.wVk = VK_RETURN;
            INPUT up = down;
            up.ki.dwFlags = KEYEVENTF_KEYUP;
            inputs.push_back(down);
            inputs.push_back(up);
            continue;
        }

        INPUT down = {};
        down.type = INPUT_KEYBOARD;
        down.ki.wScan = ch;
        down.ki.dwFlags = KEYEVENTF_UNICODE;

        INPUT up = down;
        up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

        inputs.push_back(down);
        inputs.push_back(up);
    }

    if (!inputs.empty()) {
        SendInput((UINT)inputs.size(), inputs.data(), sizeof(INPUT));
    }
}

void WaitForTargetFocus() {
    auto IsCheckmegWindow = [](HWND hWnd) -> bool {
        return (hWnd == NULL) || (hWnd == g_hSearchWnd) || (hWnd == g_hEdit) || (hWnd == g_hList) || (hWnd == g_hCreateButton);
    };

    auto RestoreFocusToLast = [&]() {
        if (g_lastForegroundWnd == NULL || IsCheckmegWindow(g_lastForegroundWnd) || !IsWindow(g_lastForegroundWnd)) {
            return;
        }

        DWORD targetThreadId = GetWindowThreadProcessId(g_lastForegroundWnd, NULL);
        DWORD myThreadId = GetCurrentThreadId();

        if (targetThreadId != 0 && targetThreadId != myThreadId) {
            AttachThreadInput(myThreadId, targetThreadId, TRUE);
        }

        SetForegroundWindow(g_lastForegroundWnd);
        SetActiveWindow(g_lastForegroundWnd);
        BringWindowToTop(g_lastForegroundWnd);

        if (targetThreadId != 0 && targetThreadId != myThreadId) {
            AttachThreadInput(myThreadId, targetThreadId, FALSE);
        }
    };

    // Actively try to restore focus to the last window and wait until it becomes foreground.
    for (int i = 0; i < 50; ++i) { // up to ~1500ms
        RestoreFocusToLastWindow();
        HWND fg = GetForegroundWindow();
        if (g_lastForegroundWnd != NULL && fg == g_lastForegroundWnd) {
            break;
        }
        // Accept any non-Checkmeg foreground if we don't know last window.
        if (g_lastForegroundWnd == NULL && !IsCheckmegWindow(fg) && fg != NULL) {
            break;
        }
        Sleep(30);
    }

    // Let the target app finish applying caret/focus.
    Sleep(120);
}

void DuplicateSelectedBookmark() {
    int idx = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
    if (idx != LB_ERR) {
        int originalIdx = (int)SendMessage(g_hList, LB_GETITEMDATA, idx, 0);
        if (originalIdx >= 0 && originalIdx < g_bookmarkManager->bookmarks.size()) {
            // Open edit dialog for the new duplicate (pass -1 as index to indicate new)
            Bookmark b = g_bookmarkManager->bookmarks[originalIdx];
            EditBookmarkAtIndex((size_t)-1, &b);
        }
    }
}

void DeleteSelectedBookmark() {
    int idx = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
    if (idx != LB_ERR) {
        int originalIdx = (int)SendMessage(g_hList, LB_GETITEMDATA, idx, 0);
        if (originalIdx >= 0 && originalIdx < g_bookmarkManager->bookmarks.size()) {
            g_bookmarkManager->remove(originalIdx);
            
            // Refresh list
            wchar_t buffer[256];
            GetWindowTextW(g_hEdit, buffer, 256);
            UpdateSearchList(WideToUtf8(buffer));

            // Restore selection to previous item
            int count = (int)SendMessage(g_hList, LB_GETCOUNT, 0, 0);
            if (count > 0) {
                int newSel = idx; 
                if (newSel >= count) newSel = count - 1; // If we deleted the last item
                
                SendMessage(g_hList, LB_SETCURSEL, newSel, 0);
            }
            SetFocus(g_hList);
        }
    }
}

void EditSelectedBookmark() {
    int idx = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
    if (idx == LB_ERR) return;
    
    int originalIdx = (int)SendMessage(g_hList, LB_GETITEMDATA, idx, 0);
    if (originalIdx < 0 || originalIdx >= g_bookmarkManager->bookmarks.size()) return;

    EditBookmarkAtIndex((size_t)originalIdx);
}

void EditBookmarkAtIndex(size_t originalIdx, const Bookmark* pDuplicateSource) {
    Bookmark b;
    bool isNew = false;
    if (pDuplicateSource) {
        b = *pDuplicateSource;
        isNew = true;
        b.deviceId = GetCurrentDeviceId();
        b.timestamp = std::time(nullptr);
    } else {
        if (originalIdx >= g_bookmarkManager->bookmarks.size()) return;
        b = g_bookmarkManager->bookmarks[originalIdx];
    }

    std::wstring currentContent = Utf8ToWide(b.content);
    std::string tagsText = FormatTags(b.tags);
    std::wstring timestampText = FormatLocalTime(b.timestamp);
    std::wstring deviceIdText = Utf8ToWide(b.deviceId);
    if (deviceIdText.empty()) deviceIdText = L"Unknown";

    // Simple Input Dialog (using a MessageBox for now is not enough, we need a real dialog)
    // Since we don't have resources, let's create a temporary modal window.
    
    static std::wstring s_editResult;
    static int s_typeSelection;
    static std::wstring s_tagsResult;
    static std::wstring s_timeResult;
    static std::wstring s_deviceIdResult;
    static bool s_validOnAnyDevice;
    static bool s_saved;
    static size_t s_originalIdx;
    s_saved = false;
    s_editResult = currentContent;
    s_tagsResult = Utf8ToWide(tagsText);
    s_timeResult = timestampText;
    s_deviceIdResult = deviceIdText;
    s_validOnAnyDevice = b.validOnAnyDevice;
    s_originalIdx = originalIdx;

    // 0=Auto, 1=Text, 2=URL, 3=File, 4=Command, 5=Binary
    if (!b.typeExplicit) {
        s_typeSelection = 0;
    } else {
        if (b.type == BookmarkType::Text) s_typeSelection = 1;
        else if (b.type == BookmarkType::URL) s_typeSelection = 2;
        else if (b.type == BookmarkType::File) s_typeSelection = 3;
        else if (b.type == BookmarkType::Command) s_typeSelection = 4;
        else if (b.type == BookmarkType::Binary) s_typeSelection = 5;
        else s_typeSelection = 0;
    }

    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT {
        static HWND hEdit;
        static HWND hCombo;
        static HWND hTags;
        static HWND hTagsLabel;
        static HWND hTimeLabel;
        static HWND hDeviceLabel;
        static HWND hValidCheck;
        static HWND hSave;
        static HWND hCancel;
        switch(msg) {
            case WM_CREATE:
            {
                // Type dropdown
                hCombo = CreateWindowExW(0, L"COMBOBOX", L"",
                    WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
                    10, 10, 360, 250, hWnd, (HMENU)3, g_hInst, NULL);

                // Emoji icons in dropdown
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F9EA  Auto");  // 🧪
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4C4  Text");  // 📄
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F310  URL");   // 🌐
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4C1  File");  // 📁
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F5A5  Command");  // 🖥️
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4BE  Binary");  // 💾
                SendMessageW(hCombo, CB_SETCURSEL, (WPARAM)s_typeSelection, 0);

                hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", ((LPCWSTR)((LPCREATESTRUCT)lParam)->lpCreateParams), 
                    WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL, 10, 40, 360, 70, hWnd, NULL, g_hInst, NULL);

                hTagsLabel = CreateWindowExW(0, L"STATIC", L"Tags (comma-separated):", WS_CHILD | WS_VISIBLE,
                    10, 115, 360, 18, hWnd, NULL, g_hInst, NULL);
                hTags = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", s_tagsResult.c_str(),
                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 10, 135, 360, 22, hWnd, NULL, g_hInst, NULL);

                std::wstring deviceLine = L"Device ID: " + s_deviceIdResult;
                hDeviceLabel = CreateWindowExW(0, L"STATIC", deviceLine.c_str(), WS_CHILD | WS_VISIBLE,
                    10, 165, 360, 18, hWnd, NULL, g_hInst, NULL);

                std::wstring timeLine = L"Saved: " + s_timeResult;
                hTimeLabel = CreateWindowExW(0, L"STATIC", timeLine.c_str(), WS_CHILD | WS_VISIBLE,
                    10, 185, 360, 18, hWnd, NULL, g_hInst, NULL);

                hValidCheck = CreateWindowExW(0, L"BUTTON", L"Valid on any device", 
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    10, 205, 360, 20, hWnd, (HMENU)4, g_hInst, NULL);
                if (s_validOnAnyDevice) {
                    SendMessageW(hValidCheck, BM_SETCHECK, BST_CHECKED, 0);
                }

                hSave = CreateWindowExW(0, L"BUTTON", L"Save", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 200, 240, 80, 30, hWnd, (HMENU)1, g_hInst, NULL);
                hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE, 290, 240, 80, 30, hWnd, (HMENU)2, g_hInst, NULL);

                if (g_hUiFont) {
                    // Use emoji font for combo so emojis render, but keep Fira Code elsewhere.
                    SendMessageW(hCombo, WM_SETFONT, (WPARAM)(g_hEmojiFont ? g_hEmojiFont : g_hUiFont), TRUE);
                    SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTagsLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTags, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hDeviceLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTimeLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hValidCheck, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hSave, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hCancel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                }

                // Enter saves, Esc cancels (from any input)
                SetWindowSubclass(hEdit, EditDialogChildSubclassProc, 1, 0);
                SetWindowSubclass(hTags, EditDialogChildSubclassProc, 2, 0);
                SetWindowSubclass(hCombo, EditDialogChildSubclassProc, 3, 0);
                SetWindowSubclass(hValidCheck, EditDialogChildSubclassProc, 4, 0);

                // Force an initial layout pass.
                RECT rc;
                GetClientRect(hWnd, &rc);
                SendMessageW(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
                SetFocus(hEdit);
                break;
            }
            case WM_SIZE:
            {
                int cw = LOWORD(lParam);
                int ch = HIWORD(lParam);
                int margin = 10;
                int w = std::max(100, cw - 2 * margin);

                int y = margin;
                int comboH = 250; // includes dropdown
                SetWindowPos(hCombo, NULL, margin, y, w, comboH, SWP_NOZORDER);
                y += 30; // actual visible height of the combobox

                int gapY = 10;
                int editH = 80;
                SetWindowPos(hEdit, NULL, margin, y + gapY, w, editH, SWP_NOZORDER);
                y = y + gapY + editH;

                SetWindowPos(hTagsLabel, NULL, margin, y + gapY, w, 18, SWP_NOZORDER);
                y = y + gapY + 18;
                SetWindowPos(hTags, NULL, margin, y + 4, w, 24, SWP_NOZORDER);
                y = y + 4 + 24;

                SetWindowPos(hDeviceLabel, NULL, margin, y + gapY, w, 18, SWP_NOZORDER);
                y = y + gapY + 18;

                SetWindowPos(hTimeLabel, NULL, margin, y + 2, w, 18, SWP_NOZORDER);
                y = y + 2 + 18;

                SetWindowPos(hValidCheck, NULL, margin, y + 4, w, 20, SWP_NOZORDER);
                y = y + 4 + 20;

                int btnW = 80;
                int btnH = 30;
                int gap = 10;
                int btnY = ch - margin - btnH;
                int cancelX = cw - margin - btnW;
                int saveX = cancelX - gap - btnW;
                SetWindowPos(hSave, NULL, saveX, btnY, btnW, btnH, SWP_NOZORDER);
                SetWindowPos(hCancel, NULL, cancelX, btnY, btnW, btnH, SWP_NOZORDER);
                return 0;
            }
            case WM_COMMAND:
                if (LOWORD(wParam) == 1) { // Save
                    int len = GetWindowTextLengthW(hEdit);
                    std::vector<wchar_t> buf(len + 1);
                    GetWindowTextW(hEdit, &buf[0], len + 1);
                    std::wstring candidateW = &buf[0];
                    int candidateTypeSel = (int)SendMessageW(hCombo, CB_GETCURSEL, 0, 0);

                    int tlen = GetWindowTextLengthW(hTags);
                    std::vector<wchar_t> tbuf(tlen + 1);
                    GetWindowTextW(hTags, &tbuf[0], tlen + 1);
                    std::wstring tagsW = &tbuf[0];

                    // Duplicate check (do not close dialog if user declines)
                    if (g_bookmarkManager) {
                        g_bookmarkManager->loadIfChanged(true);
                        // Only check duplicates if not binary (binary content is filename, which can be dup)
                        // Or check if filename is dup?
                        // For now, skip dup check for binary to avoid confusion
                        if (candidateTypeSel != 5) {
                            int dup = FindDuplicateIndex(WideToUtf8(candidateW), (int)s_originalIdx);
                            if (dup >= 0) {
                                MessageBoxW(hWnd,
                                    L"Duplicate bookmark detected.\n\nDuplicates are not allowed.",
                                    L"Duplicate",
                                    MB_OK | MB_TOPMOST | MB_SETFOREGROUND);
                                SetFocus(hEdit);
                                return 0;
                            }
                        }
                    }

                    s_editResult = candidateW;
                    s_tagsResult = tagsW;
                    s_typeSelection = candidateTypeSel;
                    s_validOnAnyDevice = (SendMessageW(hValidCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);

                    s_saved = true;
                    DestroyWindow(hWnd);
                } else if (LOWORD(wParam) == 2) { // Cancel
                    DestroyWindow(hWnd);
                }
                break;
            case WM_CLOSE:
                DestroyWindow(hWnd);
                break;
            default: return DefWindowProcW(hWnd, msg, wParam, lParam);
        }
        return 0;
    };
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"CheckmegEdit";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    RegisterClassExW(&wc);

    HWND hEditWnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"CheckmegEdit", L"Edit Bookmark", 
        WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU, 
        GetSystemMetrics(SM_CXSCREEN)/2 - 220, GetSystemMetrics(SM_CYSCREEN)/2 - 175, 440, 350, 
        g_hSearchWnd, NULL, g_hInst, (LPVOID)currentContent.c_str());

    // Message loop for the modal window
    MSG msg;
    BOOL bRet;
    while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0) {
        if (bRet == -1) break;
        if (!IsWindow(hEditWnd)) break;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    UnregisterClassW(L"CheckmegEdit", g_hInst);

    if (s_saved) {
        bool hasExplicitType = false;
        BookmarkType explicitType = BookmarkType::Text;
        if (s_typeSelection == 1) { hasExplicitType = true; explicitType = BookmarkType::Text; }
        else if (s_typeSelection == 2) { hasExplicitType = true; explicitType = BookmarkType::URL; }
        else if (s_typeSelection == 3) { hasExplicitType = true; explicitType = BookmarkType::File; }
        else if (s_typeSelection == 4) { hasExplicitType = true; explicitType = BookmarkType::Command; }
        else if (s_typeSelection == 5) { hasExplicitType = true; explicitType = BookmarkType::Binary; }
        else { hasExplicitType = false; }

        std::vector<std::string> tags = ParseTags(WideToUtf8(s_tagsResult));
        // Preserve original device ID when editing
        if (isNew) {
            g_bookmarkManager->add(WideToUtf8(s_editResult), WideToUtf8(s_deviceIdResult), s_validOnAnyDevice);
            size_t newIdx = g_bookmarkManager->bookmarks.size() - 1;
            g_bookmarkManager->update(newIdx, WideToUtf8(s_editResult), hasExplicitType, explicitType, tags, WideToUtf8(s_deviceIdResult), s_validOnAnyDevice);
        } else {
            g_bookmarkManager->update(originalIdx, WideToUtf8(s_editResult), hasExplicitType, explicitType, tags, WideToUtf8(s_deviceIdResult), s_validOnAnyDevice);
        }
        // Refresh list
        wchar_t buffer[256];
        GetWindowTextW(g_hEdit, buffer, 256);
        UpdateSearchList(WideToUtf8(buffer));
    }
}

void UpdateSearchList(const std::string& query) {
    SendMessageW(g_hList, LB_RESETCONTENT, 0, 0);
    
    std::string lowerQuery = query;
    std::transform(lowerQuery.begin(), lowerQuery.end(), lowerQuery.begin(), ::tolower);

    // Reload from file to ensure we have latest (skip binary blobs for performance)
    g_bookmarkManager->loadIfChanged(true);

    // Sort by lastUsed descending
    std::sort(g_bookmarkManager->bookmarks.begin(), g_bookmarkManager->bookmarks.end(), 
        [](const Bookmark& a, const Bookmark& b) {
            // If lastUsed is equal (e.g. both 0), sort by timestamp descending (newest first)
            if (a.lastUsed == b.lastUsed) {
                return a.timestamp > b.timestamp;
            }
            return a.lastUsed > b.lastUsed;
        });

    for (size_t i = 0; i < g_bookmarkManager->bookmarks.size(); ++i) {
        const auto& b = g_bookmarkManager->bookmarks[i];
        std::string lowerContent = b.content;
        std::transform(lowerContent.begin(), lowerContent.end(), lowerContent.begin(), ::tolower);

        std::string lowerTags;
        for (const auto& t : b.tags) {
            if (!lowerTags.empty()) lowerTags += " ";
            std::string lt = t;
            std::transform(lt.begin(), lt.end(), lt.begin(), ::tolower);
            lowerTags += lt;
        }

        if (query.empty() || lowerContent.find(lowerQuery) != std::string::npos || lowerTags.find(lowerQuery) != std::string::npos) {
            // Owner-drawn listbox: store content as string, render emoji icon in WM_DRAWITEM.
            std::wstring wDisplay = Utf8ToWide(b.content);
            int idx = (int)SendMessageW(g_hList, LB_ADDSTRING, 0, (LPARAM)wDisplay.c_str());
            SendMessage(g_hList, LB_SETITEMDATA, idx, (LPARAM)i);
        }
    }

    int itemCount = (int)SendMessage(g_hList, LB_GETCOUNT, 0, 0);
    bool showCreate = (!query.empty() && itemCount == 0);
    ShowWindow(g_hCreateButton, showCreate ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hList, showCreate ? SW_HIDE : SW_SHOW);

    // Dynamic Sizing & Centering
    HDC hdc = GetDC(g_hSearchWnd);
    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(g_hSearchWnd, hdc);
    float scale = dpiY / 96.0f;

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    
    int winW = (int)(600 * scale);
    int padding = (int)(10 * scale);
    int editH = (int)(25 * scale);
    
    int itemH = (int)SendMessage(g_hList, LB_GETITEMHEIGHT, 0, 0);
    
    // Calculate list height
    // If empty, show at least one line worth of space or just 0? 
    // Let's show a small empty space.
    int listH = 0;
    int buttonH = (int)(40 * scale);
    if (showCreate) {
        listH = buttonH;
    } else {
        listH = itemCount * itemH;
        if (listH == 0) listH = itemH;
        // Add borders/margins to list height calculation
        int border = GetSystemMetrics(SM_CYEDGE) * 2;
        listH += border;
    }
    
    int totalH = padding + editH + padding + listH + padding;
    
    // Max height constraint
    int maxH = screenH / 2;
    if (totalH > maxH) {
        totalH = maxH;
        listH = totalH - (padding + editH + padding + padding);
    }

    // Resize Main Window
    SetWindowPos(g_hSearchWnd, HWND_TOPMOST, (screenW - winW) / 2, (screenH - totalH) / 2, winW, totalH, SWP_NOZORDER | SWP_NOACTIVATE);

    // Resize Controls
    int logoW = editH;
    int gap = padding;
    if (g_hLogo) {
        SetWindowPos(g_hLogo, NULL, padding, padding, logoW, editH, SWP_NOZORDER);
    }
    int editX = padding + logoW + gap;
    int editW = winW - editX - padding;
    SetWindowPos(g_hEdit, NULL, editX, padding, editW, editH, SWP_NOZORDER);
    SetWindowPos(g_hList, NULL, padding, padding + editH + padding, winW - 2 * padding, listH, SWP_NOZORDER);
    SetWindowPos(g_hCreateButton, NULL, padding, padding + editH + padding, winW - 2 * padding, listH, SWP_NOZORDER);

    // When we show results, ensure there's a selection so Enter works naturally.
    if (!showCreate && itemCount > 0) {
        int curSel = (int)SendMessage(g_hList, LB_GETCURSEL, 0, 0);
        if (curSel == LB_ERR) {
            SendMessage(g_hList, LB_SETCURSEL, 0, 0);
        }
    }
}
