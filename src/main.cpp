#define _WIN32_WINNT 0x0600
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <windowsx.h>
#include <shellapi.h>
#include <commctrl.h>
#include <commdlg.h>
#include <objidl.h>
#include <propidl.h>
#include <gdiplus.h>
#include <winhttp.h>
#include <string>
#include <vector>
#include <algorithm>
#include <unordered_map>
#include <unordered_set>
#include <mutex>
#include <memory>
#include <thread>
#include "Bookmark.h"
#include "SupabaseAuth.h"
#include "SupabaseBookmarks.h"
#include "resource.h"
#include "Updater.h"

extern BookmarkManager* g_bookmarkManager;

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
    std::vector<std::string> tags;
    std::string cur;

    auto push = [&](std::string v) {
        v = Trim(v);
        if (v.empty()) return;
        if (!v.empty() && v[0] == '#') v = Trim(v.substr(1));
        if (v.empty()) return;
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
    UINT_PTR, DWORD_PTR) {
    if (msg == WM_KEYDOWN) {
        if (wParam == VK_TAB) {
            HWND parent = GetParent(hWnd);
            if (parent) {
                bool backwards = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                HWND next = GetNextDlgTabItem(parent, hWnd, backwards ? TRUE : FALSE);
                if (next) SetFocus(next);
            }
            return 0;
        }
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
        if (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB) return 0;
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

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
static HWND g_hMainWnd = NULL;
static HWND g_hOptionsWnd = NULL;

static constexpr UINT WM_APP_TOGGLE_SEARCH = WM_APP + 10;
static constexpr UINT WM_APP_CAPTURE_BOOKMARK = WM_APP + 11;
static constexpr UINT WM_APP_FAVICON_UPDATED = WM_APP + 12;

static HHOOK g_hKeyboardHook = NULL;
static bool g_searchSingleDown = false;
static bool g_searchSingleUsedWithOther = false;
static bool g_captureSingleDown = false;
static bool g_captureSingleUsedWithOther = false;

static std::mutex g_faviconMutex;
static std::unordered_map<std::wstring, HICON> g_faviconByHost;
static std::unordered_set<std::wstring> g_faviconInFlight;
static std::unordered_set<std::wstring> g_faviconFailed;

struct HotkeySpec {
    bool singleKey = false;
    DWORD vk = 0;
    DWORD mods = 0; // 1=Win, 2=Ctrl, 4=Shift, 8=Alt
    DWORD altSide = 0; // 0=any, 1=left, 2=right
    DWORD winSide = 0; // 0=any, 1=left, 2=right
};

static constexpr DWORD HKMOD_WIN = 1;
static constexpr DWORD HKMOD_CTRL = 2;
static constexpr DWORD HKMOD_SHIFT = 4;
static constexpr DWORD HKMOD_ALT = 8;

static HotkeySpec g_hotkeySearch;
static HotkeySpec g_hotkeyCapture;

static bool g_hotkeyCaptureMode = false;

static bool HotkeyEquals(const HotkeySpec& a, const HotkeySpec& b) {
    return a.singleKey == b.singleKey && a.vk == b.vk && a.mods == b.mods && a.altSide == b.altSide && a.winSide == b.winSide;
}

static std::wstring VkToDisplayString(DWORD vk) {
    if (vk >= 'A' && vk <= 'Z') {
        wchar_t c = (wchar_t)vk;
        return std::wstring(1, c);
    }
    if (vk >= '0' && vk <= '9') {
        wchar_t c = (wchar_t)vk;
        return std::wstring(1, c);
    }
    if (vk == VK_RMENU) return L"Right Alt";
    if (vk == VK_LMENU) return L"Left Alt";
    if (vk == VK_MENU) return L"Alt";
    if (vk == VK_LWIN) return L"Left Win";
    if (vk == VK_RWIN) return L"Right Win";
    if (vk == VK_SHIFT) return L"Shift";
    if (vk == VK_CONTROL) return L"Ctrl";
    if (vk == VK_ESCAPE) return L"Esc";
    if (vk == VK_SPACE) return L"Space";
    if (vk == VK_TAB) return L"Tab";
    if (vk == VK_RETURN) return L"Enter";

    UINT scan = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scan) {
        // Extended keys need the extended bit set in lParam.
        LONG lparam = (LONG)(scan << 16);
        wchar_t name[128] = {};
        if (GetKeyNameTextW(lparam, name, _countof(name)) > 0) {
            return name;
        }
    }
    wchar_t buf[32] = {};
    swprintf_s(buf, L"VK_%u", (unsigned)vk);
    return buf;
}

static std::wstring HotkeyToDisplayString(const HotkeySpec& hk) {
    if (hk.vk == 0) return L"(not set)";
    if (hk.singleKey || hk.mods == 0) {
        return VkToDisplayString(hk.vk);
    }
    std::wstring out;
    auto append = [&](const wchar_t* s) {
        if (!out.empty()) out += L" + ";
        out += s;
    };
    if (hk.mods & HKMOD_WIN) {
        if (hk.winSide == 1) append(L"Left Win");
        else if (hk.winSide == 2) append(L"Right Win");
        else append(L"Win");
    }
    if (hk.mods & HKMOD_CTRL) append(L"Ctrl");
    if (hk.mods & HKMOD_SHIFT) append(L"Shift");
    if (hk.mods & HKMOD_ALT) {
        if (hk.altSide == 1) append(L"Left Alt");
        else if (hk.altSide == 2) append(L"Right Alt");
        else append(L"Alt");
    }
    append(VkToDisplayString(hk.vk).c_str());
    return out;
}

static DWORD NormalizeModifierVkFromLparam(DWORD vk, LPARAM lParam) {
    // In window messages, Alt/Ctrl may come through as VK_MENU/VK_CONTROL.
    // Bit 24 in lParam indicates extended key (right-side for Alt/Ctrl).
    bool extended = ((lParam >> 24) & 1) != 0;
    if (vk == VK_MENU) return extended ? VK_RMENU : VK_LMENU;
    if (vk == VK_CONTROL) return extended ? VK_RCONTROL : VK_LCONTROL;
    return vk;
}

static DWORD NormalizeVkFromHook(const KBDLLHOOKSTRUCT* kb) {
    if (!kb) return 0;
    DWORD vk = kb->vkCode;
    // WH_KEYBOARD_LL generally provides VK_L*/VK_R*, but be defensive.
    if (vk == VK_MENU) return (kb->flags & LLKHF_EXTENDED) ? VK_RMENU : VK_LMENU;
    if (vk == VK_CONTROL) return (kb->flags & LLKHF_EXTENDED) ? VK_RCONTROL : VK_LCONTROL;
    return vk;
}

static void SetDefaultsIfUnset() {
    // Defaults: Search = Right Alt, Capture = Win + Left Alt + X
    if (g_hotkeySearch.vk == 0) {
        g_hotkeySearch.singleKey = true;
        g_hotkeySearch.vk = VK_RMENU;
        g_hotkeySearch.mods = 0;
        g_hotkeySearch.altSide = 2;
        g_hotkeySearch.winSide = 0;
    }
    if (g_hotkeyCapture.vk == 0) {
        g_hotkeyCapture.singleKey = false;
        g_hotkeyCapture.vk = 'X';
        g_hotkeyCapture.mods = HKMOD_WIN | HKMOD_ALT;
        g_hotkeyCapture.altSide = 1;
        g_hotkeyCapture.winSide = 0;
    }
}

static HotkeySpec LoadHotkeyFromRegistry(const wchar_t* namePrefix, const HotkeySpec& def) {
    HotkeySpec hk = def;
    HKEY hKey = NULL;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Checkmeg", 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
        return hk;
    }

    auto readDword = [&](const std::wstring& name, DWORD* out) {
        if (!out) return;
        DWORD type = REG_DWORD;
        DWORD size = sizeof(DWORD);
        DWORD v = 0;
        if (RegQueryValueExW(hKey, name.c_str(), NULL, &type, (LPBYTE)&v, &size) == ERROR_SUCCESS) {
            *out = v;
        }
    };

    DWORD type = hk.singleKey ? 0 : 1;
    readDword(std::wstring(namePrefix) + L"Type", &type);
    hk.singleKey = (type == 0);
    readDword(std::wstring(namePrefix) + L"Vk", &hk.vk);
    readDword(std::wstring(namePrefix) + L"Mods", &hk.mods);
    readDword(std::wstring(namePrefix) + L"AltSide", &hk.altSide);
    readDword(std::wstring(namePrefix) + L"WinSide", &hk.winSide);

    // Migration: older versions may have stored VK_MENU for Right Alt.
    if (hk.singleKey && hk.vk == VK_MENU) {
        hk.vk = VK_RMENU;
    }

    RegCloseKey(hKey);
    return hk;
}

static void SaveHotkeyToRegistry(const wchar_t* namePrefix, const HotkeySpec& hk) {
    HKEY hKey = NULL;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Checkmeg", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
        return;
    }
    auto writeDword = [&](const std::wstring& name, DWORD value) {
        RegSetValueExW(hKey, name.c_str(), 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
    };
    writeDword(std::wstring(namePrefix) + L"Type", hk.singleKey ? 0 : 1);
    writeDword(std::wstring(namePrefix) + L"Vk", hk.vk);
    writeDword(std::wstring(namePrefix) + L"Mods", hk.mods);
    writeDword(std::wstring(namePrefix) + L"AltSide", hk.altSide);
    writeDword(std::wstring(namePrefix) + L"WinSide", hk.winSide);
    RegCloseKey(hKey);
}

static void LoadHotkeySettings() {
    HotkeySpec empty;
    g_hotkeySearch = LoadHotkeyFromRegistry(L"HotkeySearch", empty);
    g_hotkeyCapture = LoadHotkeyFromRegistry(L"HotkeyCapture", empty);
    SetDefaultsIfUnset();
}

static bool IsOptionsForeground() {
    if (!g_hOptionsWnd || !IsWindow(g_hOptionsWnd)) return false;
    HWND fg = GetForegroundWindow();
    if (!fg) return false;
    return fg == g_hOptionsWnd || IsChild(g_hOptionsWnd, fg);
}

static bool HotkeyModifiersMatch(const HotkeySpec& hk) {
    if (hk.mods == 0) return true;
    bool winDownL = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0;
    bool winDownR = (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
    bool ctrlDown = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
    bool shiftDown = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
    bool altDownL = (GetAsyncKeyState(VK_LMENU) & 0x8000) != 0;
    bool altDownR = (GetAsyncKeyState(VK_RMENU) & 0x8000) != 0;

    if (hk.mods & HKMOD_WIN) {
        if (hk.winSide == 1) { if (!winDownL) return false; }
        else if (hk.winSide == 2) { if (!winDownR) return false; }
        else { if (!winDownL && !winDownR) return false; }
    }
    if (hk.mods & HKMOD_CTRL) { if (!ctrlDown) return false; }
    if (hk.mods & HKMOD_SHIFT) { if (!shiftDown) return false; }
    if (hk.mods & HKMOD_ALT) {
        if (hk.altSide == 1) { if (!altDownL) return false; }
        else if (hk.altSide == 2) { if (!altDownR) return false; }
        else { if (!altDownL && !altDownR) return false; }
    }
    return true;
}

ULONG_PTR g_gdiplusToken = 0;
Gdiplus::Image* g_logoImage = nullptr;

SupabaseAuth g_supabaseAuth;
SupabaseBookmarks g_supabaseBookmarks(&g_supabaseAuth);

static void ConfigureBookmarkSyncHooks();
static void ReloadBookmarksFromActiveBackend(HWND hWndForUi, bool showErrors);
static void RefreshSearchResultsAfterBookmarkReload();
static std::string GetBookmarksJsonPathA();
static void SyncLocalBookmarksToSupabase(HWND hWndForUi);

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
#define ID_TRAY_CHECK_UPDATE 1004

LRESULT CALLBACK MainWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK SearchWndProc(HWND, UINT, WPARAM, LPARAM);

static bool ShowEmailPasswordDialog(const wchar_t* title, std::string* outEmail, std::string* outPassword);
static void RefreshOptionsLoginButtons(HWND hBtnLogin, HWND hBtnSignup, HWND hBtnLogout);
static void TryAutoRestoreSupabaseSession();
static bool IsBinarySyncEnabled();
static void SetBinarySyncEnabled(bool enable);
LRESULT CALLBACK ListBoxProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK EditProc(HWND, UINT, WPARAM, LPARAM);
void CaptureAndBookmark();
void CreateSearchWindow();
void UpdateSearchList(const std::string& query);
void ToggleSearchWindow();
void HideSearchWindowNoRestore();
void DeleteSelectedBookmark();
void EditSelectedBookmark();
bool EditBookmarkAtIndex(size_t index, const Bookmark* pDuplicateSource = nullptr, std::string* outSavedContentUtf8 = nullptr);
void DuplicateSelectedBookmark();
void ExecuteSelectedBookmark();
void WaitForTargetFocus();
void PasteText(const std::string& text, bool clearClipboardAfter = false);
bool TryRunPowershell(const std::wstring& cmd);
void TypeText(const std::string& text);
void RestoreFocusToLastWindow();
void ShowOptionsDialog();
void InitTrayIcon(HWND hWnd);
void RemoveTrayIcon(HWND hWnd);
void AddContextMenuRegistry();
HICON GetAppIcon();
static void LoadLogoPngIfPresent();

static void CleanupFaviconCache();
static std::wstring ExtractUrlHostForFavicon(const std::string& urlUtf8);
static void EnsureFaviconFetchForHostAsync(const std::wstring& host);

static LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
static LRESULT CALLBACK OptionsHotkeyButtonSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK OptionsTabPageSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK SearchChildAltDSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR);

WNDPROC g_OriginalListBoxProc;
WNDPROC g_OriginalEditProc;

static void RunAutoUpdateCheck(bool silent) {
    std::thread([silent]() {
        Updater::UpdateInfo info = Updater::CheckForUpdates();
        if (info.available) {
            std::wstring msg = L"A new version of Checkmeg is available (" + Updater::Utf8ToWide(info.version) + L").\n\nDo you want to update now?";
            int result = MessageBoxW(NULL, msg.c_str(), L"Update Available", MB_YESNO | MB_ICONQUESTION | MB_SYSTEMMODAL);
            if (result == IDYES) {
                Updater::TriggerUpdate(info.downloadUrl);
            }
        } else if (!silent) {
            MessageBoxW(NULL, L"You are running the latest version.", L"Checkmeg", MB_OK | MB_ICONINFORMATION);
        }
    }).detach();
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    g_hInstance = hInstance;
    SetProcessDPIAware();
    g_hInst = hInstance;

    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    {
        Gdiplus::GdiplusStartupInput gdiplusStartupInput;
        Gdiplus::GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, nullptr);
    }
    
    LoadLogoPngIfPresent();

    TryAutoRestoreSupabaseSession();

    LoadHotkeySettings();
    
    RunAutoUpdateCheck(true);

    char path[MAX_PATH];
    GetModuleFileNameA(NULL, path, MAX_PATH);
    std::string exePath(path);
    std::string jsonPath = exePath.substr(0, exePath.find_last_of("\\/")) + "\\bookmarks.json";

    g_bookmarkManager = new BookmarkManager(jsonPath, false);
    ConfigureBookmarkSyncHooks();
    ReloadBookmarksFromActiveBackend(NULL, false);

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
            WNDCLASSEXW wc = {0};
            wc.cbSize = sizeof(WNDCLASSEXW);
            wc.lpfnWndProc = DefWindowProc;
            wc.hInstance = hInstance;
            wc.lpszClassName = L"CheckmegTemp";
            RegisterClassExW(&wc);
            g_hSearchWnd = CreateWindowExW(0, L"CheckmegTemp", L"", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

            if (binaryMode) {
                std::ifstream file(contentToBookmark, std::ios::binary);
                if (file) {
                    std::ostringstream ss;
                    ss << file.rdbuf();
                    std::string data = ss.str();
                    std::string base64 = base64_encode((unsigned char*)data.data(), data.size());
                    
                    std::string filename = contentToBookmark.substr(contentToBookmark.find_last_of("\\/") + 1);

                    Bookmark nb;
                    nb.type = BookmarkType::Binary;
                    nb.typeExplicit = true;
                    nb.content = filename;
                    nb.binaryData = base64;
                    nb.mimeType = "application/octet-stream";
                    nb.tags.clear();
                    nb.timestamp = std::time(nullptr);
                    nb.lastUsed = nb.timestamp;
                    nb.deviceId = GetCurrentDeviceId();
                    nb.validOnAnyDevice = true;
                    (void)EditBookmarkAtIndex((size_t)-1, &nb);
                } else {
                    MessageBoxW(NULL, L"Failed to open file for binary bookmarking.", L"Error", MB_OK);
                }
            } else {
                Bookmark nb;
                nb.type = BookmarkType::Text;
                nb.typeExplicit = false;
                nb.content = contentToBookmark;
                nb.tags.clear();
                nb.timestamp = std::time(nullptr);
                nb.lastUsed = nb.timestamp;
                nb.deviceId = GetCurrentDeviceId();
                nb.validOnAnyDevice = true;
                (void)EditBookmarkAtIndex((size_t)-1, &nb);
            }
        }

        LocalFree(argv);
        delete g_bookmarkManager;
        if (g_gdiplusToken) Gdiplus::GdiplusShutdown(g_gdiplusToken);
        return 0;
    }
    LocalFree(argv);

    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CheckmegMain";
    wc.hIcon = GetAppIcon();
    wc.hIconSm = GetAppIcon();
    RegisterClassExW(&wc);

    HWND hWnd = CreateWindowExW(0, L"CheckmegMain", L"Checkmeg", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);
    g_hMainWnd = hWnd;

    InitTrayIcon(hWnd);

    g_hKeyboardHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);
    if (!g_hKeyboardHook) {
        MessageBoxW(NULL, L"Failed to install global keyboard hook.", L"Checkmeg", MB_OK | MB_ICONERROR);
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    RemoveTrayIcon(hWnd);
    delete g_bookmarkManager;

    if (g_hKeyboardHook) {
        UnhookWindowsHookEx(g_hKeyboardHook);
        g_hKeyboardHook = NULL;
    }

    CleanupFaviconCache();

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

static void CleanupFaviconCache() {
    std::lock_guard<std::mutex> lock(g_faviconMutex);
    for (auto& kv : g_faviconByHost) {
        if (kv.second) DestroyIcon(kv.second);
    }
    g_faviconByHost.clear();
    g_faviconInFlight.clear();
    g_faviconFailed.clear();
}

static bool CrackUrlAnyScheme(const std::wstring& url, std::wstring* outHost, INTERNET_PORT* outPort, std::wstring* outPath, bool* outSecure) {
    if (!outHost || !outPort || !outPath || !outSecure) return false;

    URL_COMPONENTS parts{};
    parts.dwStructSize = sizeof(parts);
    wchar_t host[256] = {};
    wchar_t path[2048] = {};
    parts.lpszHostName = host;
    parts.dwHostNameLength = _countof(host);
    parts.lpszUrlPath = path;
    parts.dwUrlPathLength = _countof(path);

    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &parts)) return false;

    if (parts.nScheme != INTERNET_SCHEME_HTTP && parts.nScheme != INTERNET_SCHEME_HTTPS) return false;
    *outSecure = (parts.nScheme == INTERNET_SCHEME_HTTPS);
    *outHost = std::wstring(parts.lpszHostName, parts.dwHostNameLength);
    *outPort = parts.nPort;
    std::wstring p = std::wstring(parts.lpszUrlPath, parts.dwUrlPathLength);
    if (parts.dwExtraInfoLength && parts.lpszExtraInfo) {
        p += std::wstring(parts.lpszExtraInfo, parts.dwExtraInfoLength);
    }
    if (p.empty()) p = L"/";
    *outPath = p;
    return true;
}

static std::wstring ExtractUrlHostForFavicon(const std::string& urlUtf8) {
    std::wstring w = Utf8ToWide(urlUtf8);
    if (w.empty()) return L"";

    // BookmarkManager treats "www." as URL; normalize it to https for parsing.
    if (w.rfind(L"www.", 0) == 0) {
        w = L"https://" + w;
    }

    std::wstring host;
    INTERNET_PORT port = 0;
    std::wstring path;
    bool secure = false;
    if (!CrackUrlAnyScheme(w, &host, &port, &path, &secure)) return L"";
    if (host.empty()) return L"";
    return host;
}

static bool HttpGetBytes(const std::wstring& host, INTERNET_PORT port, bool secure, const std::wstring& path, std::vector<unsigned char>* outBytes) {
    if (!outBytes) return false;
    outBytes->clear();

    HINTERNET hSession = WinHttpOpen(L"Checkmeg/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return false;

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), port, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        return false;
    }

    DWORD flags = secure ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    WinHttpSetTimeouts(hRequest, 5000, 5000, 5000, 5000);
    DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
    WinHttpSetOption(hRequest, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

    BOOL ok = WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0);
    if (!ok) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    ok = WinHttpReceiveResponse(hRequest, nullptr);
    if (!ok) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    DWORD statusCode = 0;
    DWORD statusSize = sizeof(statusCode);
    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX,
        &statusCode, &statusSize, WINHTTP_NO_HEADER_INDEX);
    if (statusCode < 200 || statusCode >= 400) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    for (;;) {
        DWORD available = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &available)) break;
        if (available == 0) break;
        size_t offset = outBytes->size();
        outBytes->resize(offset + available);
        DWORD read = 0;
        if (!WinHttpReadData(hRequest, outBytes->data() + offset, available, &read)) break;
        if (read == 0) break;
        if (read < available) outBytes->resize(offset + read);
        // Avoid absurdly large downloads for a favicon.
        if (outBytes->size() > 256 * 1024) break;
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return !outBytes->empty();
}

static HICON CreateHiconFromImageBytes(const std::vector<unsigned char>& bytes, int iconSize) {
    if (!g_gdiplusToken) return NULL;
    if (bytes.empty()) return NULL;

    IStream* stream = NULL;
    if (CreateStreamOnHGlobal(NULL, TRUE, &stream) != S_OK) return NULL;
    ULONG written = 0;
    stream->Write(bytes.data(), (ULONG)bytes.size(), &written);
    LARGE_INTEGER zero{};
    stream->Seek(zero, STREAM_SEEK_SET, NULL);

    std::unique_ptr<Gdiplus::Image> img(Gdiplus::Image::FromStream(stream));
    stream->Release();

    if (!img || img->GetLastStatus() != Gdiplus::Ok) return NULL;

    Gdiplus::Bitmap scaled(iconSize, iconSize, PixelFormat32bppARGB);
    Gdiplus::Graphics g(&scaled);
    g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
    g.DrawImage(img.get(), 0, 0, iconSize, iconSize);

    HICON hIcon = NULL;
    scaled.GetHICON(&hIcon);
    return hIcon;
}

static void EnsureFaviconFetchForHostAsync(const std::wstring& host) {
    if (host.empty()) return;

    {
        std::lock_guard<std::mutex> lock(g_faviconMutex);
        if (g_faviconByHost.count(host)) return;
        if (g_faviconFailed.count(host)) return;
        if (g_faviconInFlight.count(host)) return;
        g_faviconInFlight.insert(host);
    }

    std::thread([host]() {
        std::vector<unsigned char> bytes;
        bool ok = HttpGetBytes(host, INTERNET_DEFAULT_HTTPS_PORT, true, L"/favicon.ico", &bytes);
        if (!ok) {
            bytes.clear();
            ok = HttpGetBytes(host, INTERNET_DEFAULT_HTTP_PORT, false, L"/favicon.ico", &bytes);
        }

        HICON icon = NULL;
        if (ok) {
            icon = CreateHiconFromImageBytes(bytes, 16);
        }

        {
            std::lock_guard<std::mutex> lock(g_faviconMutex);
            g_faviconInFlight.erase(host);
            if (icon) {
                g_faviconByHost[host] = icon;
            } else {
                g_faviconFailed.insert(host);
            }
        }

        if (g_hSearchWnd && IsWindow(g_hSearchWnd)) {
            PostMessageW(g_hSearchWnd, WM_APP_FAVICON_UPDATED, 0, 0);
        }
    }).detach();
}

static LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        const KBDLLHOOKSTRUCT* kb = (const KBDLLHOOKSTRUCT*)lParam;
        if (kb) {
            const DWORD vk = NormalizeVkFromHook(kb);
            const bool isDown = (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);
            const bool isUp = (wParam == WM_KEYUP || wParam == WM_SYSKEYUP);

            // Don't fire global hotkeys while the Options window is in front or while the user is rebinding.
            if (IsOptionsForeground() || g_hotkeyCaptureMode) {
                return CallNextHookEx(NULL, nCode, wParam, lParam);
            }

            auto handleSingle = [&](const HotkeySpec& spec, bool& downFlag, bool& usedFlag, UINT msgToPost) {
                if (!spec.singleKey) return false;
                if (spec.vk == 0) return false;
                if (vk == spec.vk) {
                    if (isDown) {
                        downFlag = true;
                        // Suppress AltGr (Ctrl+RightAlt) only for Right Alt single-key binding.
                        if (spec.vk == VK_RMENU) {
                            usedFlag = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                        } else {
                            usedFlag = false;
                        }
                    } else if (isUp) {
                        bool trigger = downFlag && !usedFlag;
                        downFlag = false;
                        usedFlag = false;
                        if (trigger && g_hMainWnd && IsWindow(g_hMainWnd)) {
                            PostMessageW(g_hMainWnd, msgToPost, 0, 0);
                        }
                    }
                    return true;
                }
                if (downFlag && isDown) {
                    usedFlag = true;
                }
                return false;
            };

            // First, handle possible single-key bindings.
            (void)handleSingle(g_hotkeySearch, g_searchSingleDown, g_searchSingleUsedWithOther, WM_APP_TOGGLE_SEARCH);
            (void)handleSingle(g_hotkeyCapture, g_captureSingleDown, g_captureSingleUsedWithOther, WM_APP_CAPTURE_BOOKMARK);

            // Then, handle chord bindings.
            auto handleChord = [&](const HotkeySpec& spec, UINT msgToPost) -> bool {
                if (spec.singleKey) return false;
                if (spec.vk == 0) return false;
                if (!isDown) return false;
                if (vk != spec.vk) return false;
                if (!HotkeyModifiersMatch(spec)) return false;
                if (g_hMainWnd && IsWindow(g_hMainWnd)) {
                    PostMessageW(g_hMainWnd, msgToPost, 0, 0);
                }
                return true;
            };

            if (handleChord(g_hotkeyCapture, WM_APP_CAPTURE_BOOKMARK)) return 1;
            if (handleChord(g_hotkeySearch, WM_APP_TOGGLE_SEARCH)) return 1;
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

static LRESULT CALLBACK OptionsHotkeyButtonSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) {
    if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN || msg == WM_KEYUP || msg == WM_SYSKEYUP) {
        HWND root = GetAncestor(hWnd, GA_ROOT);
        if (root && IsWindow(root)) {
            SendMessageW(root, msg, wParam, lParam);
            return 0;
        }
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

static LRESULT CALLBACK OptionsTabPageSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) {
    // Buttons inside tab pages send WM_COMMAND to the page, not the main Options window.
    // Forward those commands to the root so the existing handler works.
    if (msg == WM_COMMAND) {
        HWND root = GetAncestor(hWnd, GA_ROOT);
        if (root && IsWindow(root)) {
            SendMessageW(root, msg, wParam, lParam);
            return 0;
        }
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

static LRESULT CALLBACK SearchChildAltDSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) {
    if (msg == WM_SYSKEYDOWN) {
        if (wParam == 'D' || wParam == 'd') {
            if (g_hEdit) {
                SetFocus(g_hEdit);
                SendMessageW(g_hEdit, EM_SETSEL, 0, -1);
            }
            return 0;
        }
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

static void ConfigureBookmarkSyncHooks() {
    if (!g_bookmarkManager) return;

    g_bookmarkManager->onUpsert = [](const Bookmark& b) {
        if (!g_supabaseAuth.IsLoggedIn()) return;
        if (b.type == BookmarkType::Binary && !IsBinarySyncEnabled()) return;
        std::string err;
        (void)g_supabaseBookmarks.Upsert(b, &err);
    };

    g_bookmarkManager->onDelete = [](const Bookmark& b) {
        if (!g_supabaseAuth.IsLoggedIn()) return;
        if (b.type == BookmarkType::Binary && !IsBinarySyncEnabled()) return;
        std::string err;
        (void)g_supabaseBookmarks.DeleteById(b.id, &err);
    };
}

static void PumpUiMessages() {
    MSG msg;
    while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

class ModalProgress {
public:
    ModalProgress(HWND owner, const wchar_t* title, const wchar_t* message, int maxValue /* 0 => marquee */)
        : m_owner(owner), m_max(maxValue) {
        EnsureClassRegistered();

        if (m_owner && IsWindow(m_owner)) {
            EnableWindow(m_owner, FALSE);
        }

        const DWORD style = WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU;
        const DWORD exStyle = WS_EX_DLGMODALFRAME | WS_EX_TOPMOST;

        int w = 360;
        int h = 130;

        RECT r{ 0, 0, w, h };
        AdjustWindowRectEx(&r, style, FALSE, exStyle);
        w = r.right - r.left;
        h = r.bottom - r.top;

        int x = GetSystemMetrics(SM_CXSCREEN) / 2 - w / 2;
        int y = GetSystemMetrics(SM_CYSCREEN) / 2 - h / 2;

        m_hWnd = CreateWindowExW(
            exStyle,
            L"CheckmegProgress",
            title ? title : L"Working...",
            style,
            x, y, w, h,
            m_owner,
            NULL,
            g_hInst,
            this);

        if (!m_hWnd) {
            if (m_owner && IsWindow(m_owner)) EnableWindow(m_owner, TRUE);
            return;
        }

        ShowWindow(m_hWnd, SW_SHOW);
        UpdateWindow(m_hWnd);
        SetText(message);
        SetMax(maxValue);
        PumpUiMessages();
    }

    ~ModalProgress() {
        if (m_hWnd && IsWindow(m_hWnd)) {
            DestroyWindow(m_hWnd);
            m_hWnd = NULL;
        }
        if (m_owner && IsWindow(m_owner)) {
            EnableWindow(m_owner, TRUE);
            SetForegroundWindow(m_owner);
        }
        PumpUiMessages();
    }

    ModalProgress(const ModalProgress&) = delete;
    ModalProgress& operator=(const ModalProgress&) = delete;

    void SetText(const wchar_t* message) {
        if (!m_hText || !IsWindow(m_hText)) return;
        SetWindowTextW(m_hText, message ? message : L"");
        PumpUiMessages();
    }

    void SetMax(int maxValue /* 0 => marquee */) {
        m_max = maxValue;
        if (!m_hProg || !IsWindow(m_hProg)) return;

        if (m_max <= 0) {
            // Marquee / indeterminate
            LONG_PTR style = GetWindowLongPtrW(m_hProg, GWL_STYLE);
            style |= PBS_MARQUEE;
            SetWindowLongPtrW(m_hProg, GWL_STYLE, style);
            SendMessageW(m_hProg, PBM_SETMARQUEE, TRUE, 30);
        } else {
            // Determinate
            LONG_PTR style = GetWindowLongPtrW(m_hProg, GWL_STYLE);
            style &= ~PBS_MARQUEE;
            SetWindowLongPtrW(m_hProg, GWL_STYLE, style);
            SendMessageW(m_hProg, PBM_SETRANGE32, 0, (LPARAM)m_max);
            SendMessageW(m_hProg, PBM_SETPOS, 0, 0);
        }

        PumpUiMessages();
    }

    void SetPos(int pos) {
        if (!m_hProg || !IsWindow(m_hProg)) return;
        if (m_max <= 0) return;
        SendMessageW(m_hProg, PBM_SETPOS, (WPARAM)pos, 0);
        PumpUiMessages();
    }

private:
    static void EnsureClassRegistered() {
        static bool s_registered = false;
        if (s_registered) return;

        WNDCLASSEXW wc = { 0 };
        wc.cbSize = sizeof(wc);
        wc.hInstance = g_hInst;
        wc.lpszClassName = L"CheckmegProgress";
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.hIcon = GetAppIcon();
        wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT {
            ModalProgress* self = (ModalProgress*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);
            if (msg == WM_NCCREATE) {
                CREATESTRUCTW* cs = (CREATESTRUCTW*)lParam;
                SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
                return TRUE;
            }
            switch (msg) {
                case WM_CREATE: {
                    self = (ModalProgress*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);
                    if (!self) return 0;

                    HDC hdc = GetDC(hWnd);
                    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                    ReleaseDC(hWnd, hdc);
                    float scale = dpiY / 96.0f;

                    self->m_hText = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE,
                        (int)(16 * scale), (int)(14 * scale), (int)(320 * scale), (int)(20 * scale),
                        hWnd, NULL, g_hInst, NULL);
                    SendMessageW(self->m_hText, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                    self->m_hProg = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD | WS_VISIBLE,
                        (int)(16 * scale), (int)(44 * scale), (int)(320 * scale), (int)(18 * scale),
                        hWnd, NULL, g_hInst, NULL);
                    SendMessageW(self->m_hProg, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    return 0;
                }
                case WM_CLOSE:
                    // Prevent user from cancelling mid-sync.
                    return 0;
            }
            return DefWindowProcW(hWnd, msg, wParam, lParam);
        };

        RegisterClassExW(&wc);
        s_registered = true;
    }

    HWND m_owner = NULL;
    HWND m_hWnd = NULL;
    HWND m_hText = NULL;
    HWND m_hProg = NULL;
    int m_max = 0;
};

static void ReloadBookmarksFromActiveBackend(HWND hWndForUi, bool showErrors) {
    if (!g_bookmarkManager) return;

    if (g_supabaseAuth.IsLoggedIn()) {
        std::unique_ptr<ModalProgress> progress;
        if (hWndForUi && IsWindow(hWndForUi)) {
            progress = std::make_unique<ModalProgress>(hWndForUi, L"Supabase", L"Fetching bookmarks...", 0);
        }
        // Remote-only mode: ignore local file completely.
        g_bookmarkManager->SetUseLocalFile(false);

        std::vector<Bookmark> remote;
        std::string err;
        if (!g_supabaseBookmarks.FetchAll(&remote, &err, IsBinarySyncEnabled())) {
            g_bookmarkManager->ReplaceAll({});
            if (showErrors && hWndForUi) {
                progress.reset();
                MessageBoxW(hWndForUi, Utf8ToWide(err).c_str(), L"Supabase sync failed", MB_OK | MB_ICONERROR);
            }
        } else {
            g_bookmarkManager->ReplaceAll(remote);
        }
    } else {
        // Local mode.
        g_bookmarkManager->SetUseLocalFile(true);
        g_bookmarkManager->load(true);
    }
}

static void RefreshSearchResultsAfterBookmarkReload() {
    if (!g_hSearchWnd || !g_hEdit) return;
    if (!IsWindow(g_hSearchWnd)) return;

    wchar_t buffer[256] = {};
    GetWindowTextW(g_hEdit, buffer, _countof(buffer));
    UpdateSearchList(WideToUtf8(buffer));
}

static std::string GetBookmarksJsonPathA() {
    char path[MAX_PATH];
    GetModuleFileNameA(NULL, path, MAX_PATH);
    std::string exePath(path);
    return exePath.substr(0, exePath.find_last_of("\\/")) + "\\bookmarks.json";
}

static void SyncLocalBookmarksToSupabase(HWND hWndForUi) {
    if (!g_supabaseAuth.IsLoggedIn()) {
        MessageBoxW(hWndForUi, L"Please log in first.", L"Supabase", MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::unique_ptr<ModalProgress> progress;
    if (hWndForUi && IsWindow(hWndForUi)) {
        progress = std::make_unique<ModalProgress>(hWndForUi, L"Supabase", L"Preparing sync...", 0);
    }

    // Load local bookmarks from disk (even if the app is in remote-only mode).
    std::string jsonPath = GetBookmarksJsonPathA();
    BookmarkManager local(jsonPath, true);
    local.SetUseLocalFile(true);
    // Ensure IDs exist for older files.
    // load() already generates missing ids.

    const bool allowBinary = IsBinarySyncEnabled();

    // Fetch remote bookmarks.
    std::vector<Bookmark> remote;
    std::string err;
    if (!g_supabaseBookmarks.FetchAll(&remote, &err, allowBinary)) {
        progress.reset();
        MessageBoxW(hWndForUi, Utf8ToWide(err).c_str(), L"Supabase sync failed", MB_OK | MB_ICONERROR);
        return;
    }

    auto KeyFor = [&](const Bookmark& b) -> std::string {
        // Local-priority conflict resolution: treat same normalized content as duplicates.
        return NormalizeContentForDedup(b.content);
    };

    // De-dupe local by content key (keep first).
    std::unordered_set<std::string> localKeys;
    std::unordered_set<std::string> localIds;
    std::vector<Bookmark> uniqueLocal;
    uniqueLocal.reserve(local.bookmarks.size());
    for (const auto& b : local.bookmarks) {
        if (!allowBinary && b.type == BookmarkType::Binary) continue;
        if (b.id.empty()) continue;
        std::string k = KeyFor(b);
        if (k.empty()) continue;
        if (localKeys.insert(k).second) {
            uniqueLocal.push_back(b);
            localIds.insert(b.id);
        }
    }

    // 1) Upsert local rows (local wins for same id).
    int upserted = 0;
    int done = 0;
    const int totalOps = (int)uniqueLocal.size() + (int)remote.size();
    if (progress) {
        progress->SetText(L"Syncing to Supabase...");
        progress->SetMax(totalOps);
    }
    for (const auto& b : uniqueLocal) {
        std::string upErr;
        if (g_supabaseBookmarks.Upsert(b, &upErr)) {
            upserted++;
        }
        if (progress) progress->SetPos(++done);
    }

    // 2) Remove remote duplicates that conflict with local by content.
    int deleted = 0;
    for (const auto& rb : remote) {
        if (rb.id.empty()) continue;
        if (localIds.count(rb.id)) continue; // local already controls this row
        std::string k = KeyFor(rb);
        if (k.empty()) continue;
        if (localKeys.count(k)) {
            std::string delErr;
            if (g_supabaseBookmarks.DeleteById(rb.id, &delErr)) {
                deleted++;
            }
        }
        if (progress) progress->SetPos(++done);
    }

    // Reload remote view and refresh UI.
    ReloadBookmarksFromActiveBackend(hWndForUi, true);
    RefreshSearchResultsAfterBookmarkReload();

    progress.reset();
    std::wstring msg = L"Sync complete. Upserted: " + std::to_wstring(upserted) + L". Removed remote duplicates: " + std::to_wstring(deleted) + L".";
    MessageBoxW(hWndForUi, msg.c_str(), L"Supabase Sync", MB_OK | MB_ICONINFORMATION);
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

                // Login indicator (bottom-right): green if logged in, gray otherwise.
                {
                    int cw = rc.right - rc.left;
                    int ch = rc.bottom - rc.top;
                    int r = std::max(3, std::min(cw, ch) / 6);
                    int pad = std::max(2, r / 3);
                    int cx = cw - pad - r;
                    int cy = ch - pad - r;

                    bool loggedIn = g_supabaseAuth.IsLoggedIn();
                    Gdiplus::Color fill = loggedIn
                        ? Gdiplus::Color(255, 0, 200, 0)
                        : Gdiplus::Color(255, 160, 160, 160);

                    Gdiplus::SolidBrush brush(fill);
                    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
                    g.FillEllipse(&brush, cx - r, cy - r, r * 2, r * 2);
                }
            }

            EndPaint(hWnd, &ps);
            return 0;
        }
    }
    return DefWindowProcW(hWnd, message, wParam, lParam);
}

static void TryAutoRestoreSupabaseSession() {
    std::string err;
    (void)g_supabaseAuth.TryRestoreOrRefresh(&err);
}

static void RefreshOptionsLoginButtons(HWND hBtnLogin, HWND hBtnSignup, HWND hBtnLogout) {
    if (!hBtnLogin || !hBtnSignup || !hBtnLogout) return;
    if (g_supabaseAuth.IsLoggedIn()) {
        ShowWindow(hBtnLogin, SW_HIDE);
        ShowWindow(hBtnSignup, SW_HIDE);
        std::wstring emailW = Utf8ToWide(g_supabaseAuth.Session().email);
        if (emailW.empty()) emailW = L"(signed in)";
        std::wstring text = L"Log out (" + emailW + L")";
        SetWindowTextW(hBtnLogout, text.c_str());
        ShowWindow(hBtnLogout, SW_SHOW);
    } else {
        ShowWindow(hBtnLogout, SW_HIDE);
        ShowWindow(hBtnLogin, SW_SHOW);
        ShowWindow(hBtnSignup, SW_SHOW);
    }
}

static bool ShowEmailPasswordDialog(const wchar_t* title, std::string* outEmail, std::string* outPassword) {
    if (!outEmail || !outPassword) return false;
    *outEmail = "";
    *outPassword = "";

    static std::wstring s_email;
    static std::wstring s_password;
    static bool s_ok;
    s_email.clear();
    s_password.clear();
    s_ok = false;

    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT {
        static HWND hEmail = NULL;
        static HWND hPass = NULL;
        static HWND hOk = NULL;
        static HWND hCancel = NULL;
        static HWND hLblEmail = NULL;
        static HWND hLblPass = NULL;

        switch (msg) {
            case WM_CREATE: {
                hLblEmail = CreateWindowExW(0, L"STATIC", L"Email:", WS_CHILD | WS_VISIBLE,
                    12, 14, 360, 18, hWnd, NULL, g_hInst, NULL);
                hEmail = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
                    12, 34, 360, 24, hWnd, (HMENU)10, g_hInst, NULL);

                hLblPass = CreateWindowExW(0, L"STATIC", L"Password:", WS_CHILD | WS_VISIBLE,
                    12, 68, 360, 18, hWnd, NULL, g_hInst, NULL);
                hPass = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_PASSWORD | ES_AUTOHSCROLL,
                    12, 88, 360, 24, hWnd, (HMENU)11, g_hInst, NULL);

                hOk = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                    212, 128, 75, 28, hWnd, (HMENU)1, g_hInst, NULL);
                hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    297, 128, 75, 28, hWnd, (HMENU)2, g_hInst, NULL);

                if (g_hUiFont) {
                    SendMessageW(hLblEmail, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hEmail, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hLblPass, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hPass, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hOk, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hCancel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                }

                SetWindowSubclass(hEmail, EditDialogChildSubclassProc, 100, 0);
                SetWindowSubclass(hPass, EditDialogChildSubclassProc, 101, 0);
                SetFocus(hEmail);
                return 0;
            }
            case WM_COMMAND: {
                int id = LOWORD(wParam);
                if (id == 1) {
                    int elen = GetWindowTextLengthW(hEmail);
                    std::vector<wchar_t> ebuf(elen + 1);
                    GetWindowTextW(hEmail, ebuf.data(), elen + 1);
                    s_email = ebuf.data();

                    int plen = GetWindowTextLengthW(hPass);
                    std::vector<wchar_t> pbuf(plen + 1);
                    GetWindowTextW(hPass, pbuf.data(), plen + 1);
                    s_password = pbuf.data();

                    s_ok = true;
                    DestroyWindow(hWnd);
                    return 0;
                }
                if (id == 2) {
                    DestroyWindow(hWnd);
                    return 0;
                }
                return 0;
            }
            case WM_CLOSE:
                DestroyWindow(hWnd);
                return 0;
        }
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    };
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"CheckmegAuth";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = GetAppIcon();
    RegisterClassExW(&wc);

    int w = 400;
    int h = 200;
    int x = GetSystemMetrics(SM_CXSCREEN) / 2 - w / 2;
    int y = GetSystemMetrics(SM_CYSCREEN) / 2 - h / 2;

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST | WS_EX_CONTROLPARENT, L"CheckmegAuth", title,
        WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU,
        x, y, w, h,
        g_hSearchWnd, NULL, g_hInst, NULL);

    if (g_hSearchWnd && IsWindow(g_hSearchWnd)) EnableWindow(g_hSearchWnd, FALSE);

    MSG msg;
    BOOL bRet;
    while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0) {
        if (bRet == -1) break;
        if (!IsWindow(hDlg)) break;
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE && (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd))) {
            DestroyWindow(hDlg);
            continue;
        }
        if (IsDialogMessageW(hDlg, &msg)) continue;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (g_hSearchWnd && IsWindow(g_hSearchWnd)) EnableWindow(g_hSearchWnd, TRUE);

    UnregisterClassW(L"CheckmegAuth", g_hInst);

    if (!s_ok) return false;
    std::string email = WideToUtf8(s_email);
    std::string pass = WideToUtf8(s_password);
    if (email.empty() || pass.empty()) return false;
    *outEmail = email;
    *outPassword = pass;
    return true;
}

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_APP_TOGGLE_SEARCH:
            ToggleSearchWindow();
            break;
        case WM_APP_CAPTURE_BOOKMARK:
            CaptureAndBookmark();
            break;
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
                AppendMenuW(hMenu, MF_STRING, ID_TRAY_CHECK_UPDATE, L"Check for updates");
                AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                AppendMenuW(hMenu, MF_STRING, ID_TRAY_EXIT, L"Quit");
                int selection = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
                DestroyMenu(hMenu);
                if (selection == ID_TRAY_EXIT) {
                    PostQuitMessage(0);
                } else if (selection == ID_TRAY_OPTIONS) {
                    ShowOptionsDialog();
                } else if (selection == ID_TRAY_CHECK_UPDATE) {
                    RunAutoUpdateCheck(false);
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
            return j;
      }    
   }

   free(pImageCodecInfo);
    return -1;
}

void CaptureAndBookmark() {
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

    for (int i = 0; i < 5; ++i) {
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            CloseClipboard();
            break;
        }
        Sleep(20);
    }

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

    bool gotData = false;
    bool isText = false;
    bool isDib = false;
    bool isDrop = false;

    for (int i = 0; i < 20; ++i) {
        Sleep(50); 
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
        return;
    }

    bool clipboardOpened = false;
    for (int i = 0; i < 10; ++i) {
        if (OpenClipboard(NULL)) {
            clipboardOpened = true;
            break;
        }
        Sleep(50);
    }

    if (!clipboardOpened) {
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
                
                std::string text = WideToUtf8(wtext);
                if (!text.empty()) {
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

                    Bookmark nb;
                    nb.type = BookmarkType::Text;
                    nb.typeExplicit = false;
                    nb.content = text;
                    nb.tags.clear();
                    nb.timestamp = std::time(nullptr);
                    nb.lastUsed = nb.timestamp;
                    nb.deviceId = GetCurrentDeviceId();
                    nb.validOnAnyDevice = true;
                    (void)EditBookmarkAtIndex((size_t)-1, &nb);
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
                            Bookmark nb;
                            nb.type = BookmarkType::Binary;
                            nb.typeExplicit = true;
                            nb.content = name;
                            nb.binaryData = base64;
                            nb.mimeType = "image/png";
                            nb.tags.clear();
                            nb.timestamp = std::time(nullptr);
                            nb.lastUsed = nb.timestamp;
                            nb.deviceId = GetCurrentDeviceId();
                            nb.validOnAnyDevice = true;
                            (void)EditBookmarkAtIndex((size_t)-1, &nb);
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

                    Bookmark nb;
                    nb.type = BookmarkType::Binary;
                    nb.typeExplicit = true;
                    nb.content = filename;
                    nb.binaryData = base64;
                    nb.mimeType = "application/octet-stream";
                    nb.tags.clear();
                    nb.timestamp = std::time(nullptr);
                    nb.lastUsed = nb.timestamp;
                    nb.deviceId = GetCurrentDeviceId();
                    nb.validOnAnyDevice = true;
                    (void)EditBookmarkAtIndex((size_t)-1, &nb);
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
        
        // Improved focus stealing
        HWND hFore = GetForegroundWindow();
        DWORD dwForeID = GetWindowThreadProcessId(hFore, NULL);
        DWORD dwCurID = GetCurrentThreadId();
        
        if (dwForeID != dwCurID) {
            AttachThreadInput(dwForeID, dwCurID, TRUE);
            ShowWindow(g_hSearchWnd, SW_SHOW);
            SetForegroundWindow(g_hSearchWnd);
            SetFocus(g_hEdit);
            AttachThreadInput(dwForeID, dwCurID, FALSE);
        } else {
            ShowWindow(g_hSearchWnd, SW_SHOW);
            SetForegroundWindow(g_hSearchWnd);
            SetFocus(g_hEdit);
        }

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
        case WM_APP_FAVICON_UPDATED:
            if (g_hList && IsWindow(g_hList)) {
                InvalidateRect(g_hList, NULL, TRUE);
                UpdateWindow(g_hList);
            }
            return 0;
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
                    Bookmark nb;
                    nb.type = BookmarkType::Text;
                    nb.typeExplicit = false;
                    nb.content = text;
                    nb.tags.clear();
                    nb.timestamp = std::time(nullptr);
                    nb.lastUsed = nb.timestamp;
                    nb.deviceId = GetCurrentDeviceId();
                    nb.validOnAnyDevice = true;

                    std::string saved;
                    if (EditBookmarkAtIndex((size_t)-1, &nb, &saved)) {
                        SetWindowTextW(g_hEdit, Utf8ToWide(saved).c_str());
                        UpdateSearchList(saved);
                    } else {
                        // No-op on cancel
                    }
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
                        const std::string rawContentUtf8 = b.content;
                        text = b.sensitive ? L"***" : Utf8ToWide(b.content);
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

                    // Icon (favicon for URLs, emoji for everything else)
                    std::wstring icon;
                    if (type == BookmarkType::URL) icon = L"\U0001F310";       // fallback 🌐 while favicon loads
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

                    RECT rcIcon = rc;
                    rcIcon.right = rcIcon.left + 28;

                    bool drewFavicon = false;
                    if (type == BookmarkType::URL) {
                        std::string hostSource = (originalIdx >= 0 && originalIdx < (int)g_bookmarkManager->bookmarks.size())
                            ? g_bookmarkManager->bookmarks[originalIdx].content
                            : WideToUtf8(text);
                        std::wstring host = ExtractUrlHostForFavicon(hostSource);
                        if (!host.empty()) {
                            HICON hFav = NULL;
                            {
                                std::lock_guard<std::mutex> lock(g_faviconMutex);
                                auto it = g_faviconByHost.find(host);
                                if (it != g_faviconByHost.end()) hFav = it->second;
                            }

                            if (hFav) {
                                int iconSize = 16;
                                int x = rcIcon.left + 2;
                                int y = ((dis->rcItem.top + dis->rcItem.bottom) / 2) - (iconSize / 2);
                                DrawIconEx(dis->hDC, x, y, hFav, iconSize, iconSize, 0, NULL, DI_NORMAL);
                                drewFavicon = true;
                            } else {
                                EnsureFaviconFetchForHostAsync(host);
                            }
                        }
                    }

                    HFONT oldFont = (HFONT)SelectObject(dis->hDC, g_hEmojiFont ? g_hEmojiFont : g_hUiFont);
                    if (!drewFavicon) {
                        DrawTextW(dis->hDC, icon.c_str(), (int)icon.size(), &rcIcon, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
                    }

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
    SetWindowSubclass(g_hCreateButton, SearchChildAltDSubclassProc, 1, 0);
    
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
        case WM_SYSKEYDOWN:
            if (wParam == 'D' || wParam == 'd') {
                SetFocus(g_hEdit);
                SendMessageW(g_hEdit, EM_SETSEL, 0, -1);
                return 0;
            }
            break;
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
                            preview = g_bookmarkManager->bookmarks[originalIdx].sensitive ? L"***" : Utf8ToWide(g_bookmarkManager->bookmarks[originalIdx].content);
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

static void ClearClipboardBestEffort() {
    for (int i = 0; i < 5; ++i) {
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            CloseClipboard();
            return;
        }
        Sleep(20);
    }
}

void PasteText(const std::string& text, bool clearClipboardAfter) {
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

    if (clearClipboardAfter) {
        ClearClipboardBestEffort();
    }
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
                    PasteText(b.content, b.sensitive);
                }
            } else if (b.type == BookmarkType::Command) {
                if (!TryRunPowershell(wContent)) {
                    // If it fails, fall back to typing it as text.
                    WaitForTargetFocus();
                    PasteText(b.content, b.sensitive);
                }
            } else {
                WaitForTargetFocus();
                PasteText(b.content, b.sensitive);
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

static bool IsBinarySyncEnabled() {
    DWORD value = 1;
    HKEY hKey = NULL;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Checkmeg", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        DWORD type = REG_DWORD;
        DWORD size = sizeof(value);
        (void)RegQueryValueExW(hKey, L"SyncBinaryToSupabase", NULL, &type, (LPBYTE)&value, &size);
        RegCloseKey(hKey);
    }
    return value != 0;
}

static void SetBinarySyncEnabled(bool enable) {
    HKEY hKey = NULL;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Checkmeg", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        DWORD value = enable ? 1 : 0;
        RegSetValueExW(hKey, L"SyncBinaryToSupabase", 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
        RegCloseKey(hKey);
    }
}

void ShowOptionsDialog() {
    // Singleton: if already open, focus it and return.
    if (g_hOptionsWnd && IsWindow(g_hOptionsWnd)) {
        ShowWindow(g_hOptionsWnd, SW_SHOW);
        SetForegroundWindow(g_hOptionsWnd);
        BringWindowToTop(g_hOptionsWnd);
        return;
    }

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
        static HWND hTab = NULL;
        static HWND hPageLogin = NULL;
        static HWND hPageSystem = NULL;
        static HWND hPageHotkeys = NULL;
        static HWND hPageData = NULL;
        static HWND hBtnCtx = NULL;
        static HWND hBtnRun = NULL;
        static HWND hBtnOpenJson = NULL;
        static HWND hBtnSyncLocal = NULL;
        static HWND hBtnRefreshSupabase = NULL;
        static HWND hChkBinarySync = NULL;
        static HWND hBtnOk = NULL;
        static HWND hBtnLogin = NULL;
        static HWND hBtnSignup = NULL;
        static HWND hBtnLogout = NULL;
        static HWND hBtnBindSearch = NULL;
        static HWND hBtnBindCapture = NULL;
        static int activeBind = 0; // 0 none, 1 search, 2 capture
        static DWORD pendingModifierVk = 0;
        static bool sawNonModifier = false;
        static bool isRegistered = false;
        static bool isRunAtStartup = false;

        auto IsModifierVk = [](DWORD vk) -> bool {
            return vk == VK_SHIFT || vk == VK_LSHIFT || vk == VK_RSHIFT ||
                   vk == VK_CONTROL || vk == VK_LCONTROL || vk == VK_RCONTROL ||
                   vk == VK_MENU || vk == VK_LMENU || vk == VK_RMENU ||
                   vk == VK_LWIN || vk == VK_RWIN;
        };

        auto ShowOnlyPage = [&](HWND page) {
            if (hPageLogin) ShowWindow(hPageLogin, (page == hPageLogin) ? SW_SHOW : SW_HIDE);
            if (hPageSystem) ShowWindow(hPageSystem, (page == hPageSystem) ? SW_SHOW : SW_HIDE);
            if (hPageHotkeys) ShowWindow(hPageHotkeys, (page == hPageHotkeys) ? SW_SHOW : SW_HIDE);
            if (hPageData) ShowWindow(hPageData, (page == hPageData) ? SW_SHOW : SW_HIDE);
        };

        auto LayoutTabPages = [&](float scale) {
            if (!hTab) return;
            RECT rc;
            GetClientRect(hTab, &rc);
            TabCtrl_AdjustRect(hTab, FALSE, &rc);
            int x = rc.left;
            int y = rc.top;
            int w = rc.right - rc.left;
            int h = rc.bottom - rc.top;
            if (hPageLogin) SetWindowPos(hPageLogin, NULL, x, y, w, h, SWP_NOZORDER);
            if (hPageSystem) SetWindowPos(hPageSystem, NULL, x, y, w, h, SWP_NOZORDER);
            if (hPageHotkeys) SetWindowPos(hPageHotkeys, NULL, x, y, w, h, SWP_NOZORDER);
            if (hPageData) SetWindowPos(hPageData, NULL, x, y, w, h, SWP_NOZORDER);
        };

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

                // Ensure current hotkeys are loaded
                LoadHotkeySettings();

                // Tabs
                hTab = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                    (int)(20 * scale), (int)(95 * scale), (int)(360 * scale), (int)(440 * scale), hWnd, (HMENU)100, g_hInst, NULL);
                SendMessageW(hTab, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                TCITEMW item{};
                item.mask = TCIF_TEXT;
                item.pszText = (LPWSTR)L"Login";
                TabCtrl_InsertItem(hTab, 0, &item);
                item.pszText = (LPWSTR)L"System";
                TabCtrl_InsertItem(hTab, 1, &item);
                item.pszText = (LPWSTR)L"Hotkeys";
                TabCtrl_InsertItem(hTab, 2, &item);
                item.pszText = (LPWSTR)L"Data";
                TabCtrl_InsertItem(hTab, 3, &item);

                hPageLogin = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE,
                    0, 0, 0, 0, hTab, NULL, g_hInst, NULL);
                hPageSystem = CreateWindowExW(0, L"STATIC", L"", WS_CHILD,
                    0, 0, 0, 0, hTab, NULL, g_hInst, NULL);
                hPageHotkeys = CreateWindowExW(0, L"STATIC", L"", WS_CHILD,
                    0, 0, 0, 0, hTab, NULL, g_hInst, NULL);
                hPageData = CreateWindowExW(0, L"STATIC", L"", WS_CHILD,
                    0, 0, 0, 0, hTab, NULL, g_hInst, NULL);

                SetWindowSubclass(hPageLogin, OptionsTabPageSubclassProc, 1, 0);
                SetWindowSubclass(hPageSystem, OptionsTabPageSubclassProc, 1, 0);
                SetWindowSubclass(hPageHotkeys, OptionsTabPageSubclassProc, 1, 0);
                SetWindowSubclass(hPageData, OptionsTabPageSubclassProc, 1, 0);

                LayoutTabPages(scale);
                ShowOnlyPage(hPageLogin);

                // Login tab
                hBtnLogin = CreateWindowExW(0, L"BUTTON", L"Log in", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(25 * scale), (int)(300 * scale), (int)(30 * scale), hPageLogin, (HMENU)10, g_hInst, NULL);
                hBtnSignup = CreateWindowExW(0, L"BUTTON", L"Sign up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(60 * scale), (int)(300 * scale), (int)(30 * scale), hPageLogin, (HMENU)11, g_hInst, NULL);
                hBtnLogout = CreateWindowExW(0, L"BUTTON", L"Log out", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(42 * scale), (int)(300 * scale), (int)(30 * scale), hPageLogin, (HMENU)12, g_hInst, NULL);

                SendMessageW(hBtnLogin, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SendMessageW(hBtnSignup, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SendMessageW(hBtnLogout, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                RefreshOptionsLoginButtons(hBtnLogin, hBtnSignup, hBtnLogout);

                // System tab
                hBtnRun = CreateWindowExW(0, L"BUTTON", L"Run at Startup", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    (int)(20 * scale), (int)(25 * scale), (int)(300 * scale), (int)(25 * scale), hPageSystem, (HMENU)3, g_hInst, NULL);
                SendMessageW(hBtnRun, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SendMessageW(hBtnRun, BM_SETCHECK, isRunAtStartup ? BST_CHECKED : BST_UNCHECKED, 0);

                // Context Menu Button
                const wchar_t* btnText = isRegistered ? L"Remove from File Context Menu" : L"Add to File Context Menu";
                hBtnCtx = CreateWindowExW(0, L"BUTTON", btnText, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(60 * scale), (int)(300 * scale), (int)(30 * scale), hPageSystem, (HMENU)2, g_hInst, NULL);
                SendMessageW(hBtnCtx, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                // Hotkeys tab
                CreateWindowExW(0, L"STATIC", L"Search", WS_CHILD | WS_VISIBLE,
                    (int)(20 * scale), (int)(20 * scale), (int)(80 * scale), (int)(22 * scale), hPageHotkeys, NULL, g_hInst, NULL);

                hBtnBindSearch = CreateWindowExW(0, L"BUTTON", HotkeyToDisplayString(g_hotkeySearch).c_str(),
                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(70 * scale), (int)(300 * scale), (int)(30 * scale), hPageHotkeys, (HMENU)20, g_hInst, NULL);
                SendMessageW(hBtnBindSearch, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SetWindowSubclass(hBtnBindSearch, OptionsHotkeyButtonSubclassProc, 1, 0);

                CreateWindowExW(0, L"STATIC", L"Bookmark", WS_CHILD | WS_VISIBLE,
                    (int)(20 * scale), (int)(120 * scale), (int)(120 * scale), (int)(22 * scale), hPageHotkeys, NULL, g_hInst, NULL);

                hBtnBindCapture = CreateWindowExW(0, L"BUTTON", HotkeyToDisplayString(g_hotkeyCapture).c_str(),
                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(170 * scale), (int)(300 * scale), (int)(30 * scale), hPageHotkeys, (HMENU)21, g_hInst, NULL);
                SendMessageW(hBtnBindCapture, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SetWindowSubclass(hBtnBindCapture, OptionsHotkeyButtonSubclassProc, 1, 0);

                // Data tab
                hBtnRefreshSupabase = CreateWindowExW(0, L"BUTTON", L"Refresh from Supabase", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(25 * scale), (int)(300 * scale), (int)(30 * scale), hPageData, (HMENU)7, g_hInst, NULL);
                SendMessageW(hBtnRefreshSupabase, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                hBtnOpenJson = CreateWindowExW(0, L"BUTTON", L"Export Bookmarks File", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(60 * scale), (int)(300 * scale), (int)(30 * scale), hPageData, (HMENU)4, g_hInst, NULL);
                SendMessageW(hBtnOpenJson, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                // Sync local -> Supabase (local priority)
                hBtnSyncLocal = CreateWindowExW(0, L"BUTTON", L"Sync local data to Supabase", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                    (int)(20 * scale), (int)(95 * scale), (int)(300 * scale), (int)(30 * scale), hPageData, (HMENU)5, g_hInst, NULL);
                SendMessageW(hBtnSyncLocal, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);

                // Toggle Binary syncing
                hChkBinarySync = CreateWindowExW(0, L"BUTTON", L"Sync binary bookmarks to Supabase",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    (int)(20 * scale), (int)(135 * scale), (int)(320 * scale), (int)(25 * scale), hPageData, (HMENU)6, g_hInst, NULL);
                SendMessageW(hChkBinarySync, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SendMessageW(hChkBinarySync, BM_SETCHECK, IsBinarySyncEnabled() ? BST_CHECKED : BST_UNCHECKED, 0);

                hBtnOk = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                    (int)(300 * scale), (int)(545 * scale), (int)(80 * scale), (int)(30 * scale), hWnd, (HMENU)1, g_hInst, NULL);
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

                // Separator under header
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(220, 220, 220));
                HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, (int)(20 * scale), (int)(85 * scale), NULL);
                LineTo(hdc, (int)(380 * scale), (int)(85 * scale));
                SelectObject(hdc, oldPen);
                DeleteObject(hPen);

                SelectObject(hdc, oldFont);
                EndPaint(hWnd, &ps);
                return 0;
            }
            case WM_NOTIFY: {
                const NMHDR* nm = (const NMHDR*)lParam;
                if (nm && nm->hwndFrom == hTab && nm->code == TCN_SELCHANGE) {
                    int sel = TabCtrl_GetCurSel(hTab);
                    if (sel == 0) ShowOnlyPage(hPageLogin);
                    else if (sel == 1) ShowOnlyPage(hPageSystem);
                    else if (sel == 2) ShowOnlyPage(hPageHotkeys);
                    else if (sel == 3) ShowOnlyPage(hPageData);
                    return 0;
                }
                break;
            }
            case WM_KEYDOWN:
            case WM_SYSKEYDOWN: {
                if (activeBind == 0) break;
                DWORD vk = NormalizeModifierVkFromLparam((DWORD)wParam, lParam);
                if (vk == VK_ESCAPE) {
                    activeBind = 0;
                    g_hotkeyCaptureMode = false;
                    pendingModifierVk = 0;
                    sawNonModifier = false;
                    if (hBtnBindSearch) SetWindowTextW(hBtnBindSearch, HotkeyToDisplayString(g_hotkeySearch).c_str());
                    if (hBtnBindCapture) SetWindowTextW(hBtnBindCapture, HotkeyToDisplayString(g_hotkeyCapture).c_str());
                    return 0;
                }

                if (IsModifierVk(vk)) {
                    pendingModifierVk = vk;
                    return 0;
                }
                sawNonModifier = true;

                HotkeySpec hk;
                hk.vk = vk;
                bool winL = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0;
                bool winR = (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
                bool ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                bool shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                bool altL = (GetAsyncKeyState(VK_LMENU) & 0x8000) != 0;
                bool altR = (GetAsyncKeyState(VK_RMENU) & 0x8000) != 0;

                if (winL || winR) {
                    hk.mods |= HKMOD_WIN;
                    if (winL && !winR) hk.winSide = 1;
                    else if (winR && !winL) hk.winSide = 2;
                }
                if (ctrl) hk.mods |= HKMOD_CTRL;
                if (shift) hk.mods |= HKMOD_SHIFT;
                if (altL || altR) {
                    hk.mods |= HKMOD_ALT;
                    if (altL && !altR) hk.altSide = 1;
                    else if (altR && !altL) hk.altSide = 2;
                }

                // Auto-detect: non-modifier key binds as a combo (mods may be 0)
                hk.singleKey = false;

                // Prevent the same hotkey for both actions.
                if (activeBind == 1) {
                    if (HotkeyEquals(hk, g_hotkeyCapture)) {
                        MessageBoxW(hWnd, L"That hotkey is already used for Bookmark.\n\nChoose a different one.", L"Hotkeys", MB_OK | MB_ICONWARNING);
                        return 0;
                    }
                    g_hotkeySearch = hk;
                    SaveHotkeyToRegistry(L"HotkeySearch", g_hotkeySearch);
                    if (hBtnBindSearch) SetWindowTextW(hBtnBindSearch, HotkeyToDisplayString(g_hotkeySearch).c_str());
                } else if (activeBind == 2) {
                    if (HotkeyEquals(hk, g_hotkeySearch)) {
                        MessageBoxW(hWnd, L"That hotkey is already used for Search.\n\nChoose a different one.", L"Hotkeys", MB_OK | MB_ICONWARNING);
                        return 0;
                    }
                    g_hotkeyCapture = hk;
                    SaveHotkeyToRegistry(L"HotkeyCapture", g_hotkeyCapture);
                    if (hBtnBindCapture) SetWindowTextW(hBtnBindCapture, HotkeyToDisplayString(g_hotkeyCapture).c_str());
                }

                activeBind = 0;
                g_hotkeyCaptureMode = false;
                pendingModifierVk = 0;
                sawNonModifier = false;
                return 0;
            }
            case WM_KEYUP:
            case WM_SYSKEYUP: {
                if (activeBind == 0) break;
                DWORD vk = NormalizeModifierVkFromLparam((DWORD)wParam, lParam);
                if (!sawNonModifier && pendingModifierVk != 0 && vk == pendingModifierVk) {
                    // Modifier-only binding (press+release with no non-modifier key).
                    HotkeySpec hk;
                    hk.singleKey = true;
                    hk.vk = vk;
                    hk.mods = 0;
                    hk.altSide = 0;
                    hk.winSide = 0;

                    if (activeBind == 1) {
                        if (HotkeyEquals(hk, g_hotkeyCapture)) {
                            MessageBoxW(hWnd, L"That hotkey is already used for Bookmark.\n\nChoose a different one.", L"Hotkeys", MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                        g_hotkeySearch = hk;
                        SaveHotkeyToRegistry(L"HotkeySearch", g_hotkeySearch);
                        if (hBtnBindSearch) SetWindowTextW(hBtnBindSearch, HotkeyToDisplayString(g_hotkeySearch).c_str());
                    } else if (activeBind == 2) {
                        if (HotkeyEquals(hk, g_hotkeySearch)) {
                            MessageBoxW(hWnd, L"That hotkey is already used for Search.\n\nChoose a different one.", L"Hotkeys", MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                        g_hotkeyCapture = hk;
                        SaveHotkeyToRegistry(L"HotkeyCapture", g_hotkeyCapture);
                        if (hBtnBindCapture) SetWindowTextW(hBtnBindCapture, HotkeyToDisplayString(g_hotkeyCapture).c_str());
                    }

                    activeBind = 0;
                    g_hotkeyCaptureMode = false;
                    pendingModifierVk = 0;
                    sawNonModifier = false;
                    return 0;
                }
                break;
            }
            case WM_COMMAND: {
                int id = LOWORD(wParam);
                if (id == 1) { // OK
                    DestroyWindow(hWnd);
                } else if (id == 20 || id == 21) {
                    // Start capturing a new hotkey
                    activeBind = (id == 20) ? 1 : 2;
                    g_hotkeyCaptureMode = true;
                    pendingModifierVk = 0;
                    sawNonModifier = false;
                    if (id == 20 && hBtnBindSearch) SetWindowTextW(hBtnBindSearch, L"Press keys… (Esc to cancel)");
                    if (id == 21 && hBtnBindCapture) SetWindowTextW(hBtnBindCapture, L"Press keys… (Esc to cancel)");
                    // Keep focus on the button so keystrokes are forwarded via subclass.
                    if (id == 20 && hBtnBindSearch) SetFocus(hBtnBindSearch);
                    if (id == 21 && hBtnBindCapture) SetFocus(hBtnBindCapture);
                } else if (id == 10) { // Log in
                    std::string email;
                    std::string password;
                    if (!ShowEmailPasswordDialog(L"Log in", &email, &password)) return 0;

                    std::string err;
                    bool ok = g_supabaseAuth.SignInWithPassword(email, password, &err);
                    if (ok) {
                        MessageBoxW(hWnd, L"Logged in.", L"Success", MB_OK | MB_ICONINFORMATION);
                    } else {
                        MessageBoxW(hWnd, Utf8ToWide(err).c_str(), L"Login failed", MB_OK | MB_ICONERROR);
                    }

                    if (ok && g_supabaseAuth.IsLoggedIn()) {
                        ReloadBookmarksFromActiveBackend(hWnd, true);
                        RefreshSearchResultsAfterBookmarkReload();
                    }
                    RefreshOptionsLoginButtons(hBtnLogin, hBtnSignup, hBtnLogout);
                    if (g_hLogo) InvalidateRect(g_hLogo, NULL, TRUE);
                } else if (id == 11) { // Sign up
                    std::string email;
                    std::string password;
                    if (!ShowEmailPasswordDialog(L"Sign up", &email, &password)) return 0;

                    std::string err;
                    bool ok = g_supabaseAuth.SignUpWithPassword(email, password, &err);
                    if (ok && g_supabaseAuth.IsLoggedIn()) {
                        MessageBoxW(hWnd, L"Signed up and logged in.", L"Success", MB_OK | MB_ICONINFORMATION);
                    } else if (ok) {
                        // Signup succeeded but may require email confirmation.
                        if (err.empty()) err = "Sign up succeeded. Please log in.";
                        MessageBoxW(hWnd, Utf8ToWide(err).c_str(), L"Sign up", MB_OK | MB_ICONINFORMATION);
                    } else {
                        MessageBoxW(hWnd, Utf8ToWide(err).c_str(), L"Sign up failed", MB_OK | MB_ICONERROR);
                    }

                    if (ok && g_supabaseAuth.IsLoggedIn()) {
                        ReloadBookmarksFromActiveBackend(hWnd, true);
                        RefreshSearchResultsAfterBookmarkReload();
                    }
                    RefreshOptionsLoginButtons(hBtnLogin, hBtnSignup, hBtnLogout);
                    if (g_hLogo) InvalidateRect(g_hLogo, NULL, TRUE);
                } else if (id == 12) { // Log out
                    g_supabaseAuth.Logout();
                    ReloadBookmarksFromActiveBackend(hWnd, false);
                    RefreshSearchResultsAfterBookmarkReload();
                    RefreshOptionsLoginButtons(hBtnLogin, hBtnSignup, hBtnLogout);
                    if (g_hLogo) InvalidateRect(g_hLogo, NULL, TRUE);
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
                } else if (id == 5) { // Sync local -> Supabase
                    SyncLocalBookmarksToSupabase(hWnd);
                } else if (id == 7) { // Refresh from Supabase
                    if (!g_supabaseAuth.IsLoggedIn()) {
                        MessageBoxW(hWnd, L"Please log in first.", L"Supabase", MB_OK | MB_ICONINFORMATION);
                        return 0;
                    }
                    ReloadBookmarksFromActiveBackend(hWnd, true);
                    RefreshSearchResultsAfterBookmarkReload();
                } else if (id == 6) { // Toggle binary syncing
                    if (HIWORD(wParam) == BN_CLICKED && hChkBinarySync) {
                        bool check = (SendMessageW(hChkBinarySync, BM_GETCHECK, 0, 0) == BST_CHECKED);
                        SetBinarySyncEnabled(check);
                        if (g_supabaseAuth.IsLoggedIn()) {
                            ReloadBookmarksFromActiveBackend(hWnd, false);
                            RefreshSearchResultsAfterBookmarkReload();
                        }
                    }
                }
                return 0;
            }
            case WM_DESTROY:
                if (hTitleFont) DeleteObject(hTitleFont);
                if (hVersionFont) DeleteObject(hVersionFont);
                if (hSectionFont) DeleteObject(hSectionFont);
                g_hotkeyCaptureMode = false;
                g_hOptionsWnd = NULL;
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

    // Size window by desired CLIENT size to avoid clipped bottom buttons.
    int clientW = (int)(400 * scale);
    int clientH = (int)(590 * scale);
    RECT r{ 0, 0, clientW, clientH };
    DWORD style = WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU;
    DWORD exStyle = WS_EX_DLGMODALFRAME | WS_EX_TOPMOST;
    AdjustWindowRectEx(&r, style, FALSE, exStyle);
    int w = r.right - r.left;
    int h = r.bottom - r.top;
    int x = GetSystemMetrics(SM_CXSCREEN) / 2 - w / 2;
    int y = GetSystemMetrics(SM_CYSCREEN) / 2 - h / 2;

    HWND hDlg = CreateWindowExW(exStyle, L"CheckmegOptions", L"Options",
        style,
        x, y, w, h,
        g_hSearchWnd, NULL, g_hInst, NULL);

    g_hOptionsWnd = hDlg;

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

bool EditBookmarkAtIndex(size_t originalIdx, const Bookmark* pDuplicateSource, std::string* outSavedContentUtf8) {
    Bookmark b;
    bool isNew = false;
    if (pDuplicateSource) {
        b = *pDuplicateSource;
        isNew = true;
        b.deviceId = GetCurrentDeviceId();
        b.timestamp = std::time(nullptr);
    } else {
        if (originalIdx >= g_bookmarkManager->bookmarks.size()) return false;
        b = g_bookmarkManager->bookmarks[originalIdx];
    }

    // Never show sensitive content in plaintext.
    std::string originalSensitiveContentUtf8 = b.sensitive ? b.content : std::string();
    std::wstring currentContent = b.sensitive ? L"" : Utf8ToWide(b.content);
    std::string tagsText = FormatTags(b.tags);
    std::wstring timestampText = FormatLocalTime(b.timestamp);
    std::wstring deviceIdText = Utf8ToWide(b.deviceId);
    if (deviceIdText.empty()) deviceIdText = L"Unknown";


    static std::wstring s_editResult;
    static int s_typeSelection;
    static std::wstring s_tagsResult;
    static std::wstring s_timeResult;
    static std::wstring s_deviceIdResult;
    static bool s_validOnAnyDevice;
    static bool s_sensitive;
    static bool s_saved;
    static size_t s_originalIdx;
    static std::string s_originalSensitiveContentUtf8;
    static bool s_sensitivePlaceholder;
    s_saved = false;
    s_editResult = currentContent;
    s_tagsResult = Utf8ToWide(tagsText);
    s_timeResult = timestampText;
    s_deviceIdResult = deviceIdText;
    s_validOnAnyDevice = b.validOnAnyDevice;
    s_sensitive = b.sensitive;
    s_originalIdx = isNew ? (size_t)-1 : originalIdx;
    s_originalSensitiveContentUtf8 = originalSensitiveContentUtf8;
    s_sensitivePlaceholder = false;

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
        static HWND hSensitiveCheck;
        static HWND hSave;
        static HWND hCancel;
        auto SwapEditToSensitiveMode = [&](bool enableSensitive) {
            if (!hEdit || !IsWindow(hEdit)) return;

            // Capture current visible text if transitioning from non-sensitive to sensitive.
            if (enableSensitive) {
                if (!s_sensitive) {
                    int len = GetWindowTextLengthW(hEdit);
                    if (len > 0) {
                        std::vector<wchar_t> buf(len + 1);
                        GetWindowTextW(hEdit, buf.data(), len + 1);
                        s_originalSensitiveContentUtf8 = WideToUtf8(std::wstring(buf.data()));
                    }
                }
            }

            // Replace the edit control:
            // - Sensitive: single-line password edit (never shows plaintext)
            // - Non-sensitive: multiline edit
            RECT rc;
            GetWindowRect(hEdit, &rc);
            MapWindowPoints(NULL, hWnd, (LPPOINT)&rc, 2);

            DestroyWindow(hEdit);
            hEdit = NULL;

            if (enableSensitive) {
                hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"***",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
                    rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
                    hWnd, (HMENU)10, g_hInst, NULL);
                if (g_hUiFont) SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SetWindowSubclass(hEdit, EditDialogChildSubclassProc, 1, 0);
                SendMessageW(hEdit, EM_SETSEL, 0, -1);
                s_sensitivePlaceholder = true;
            } else {
                std::wstring restore = Utf8ToWide(s_originalSensitiveContentUtf8);
                hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", restore.c_str(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL,
                    rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
                    hWnd, (HMENU)10, g_hInst, NULL);
                if (g_hUiFont) SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                SetWindowSubclass(hEdit, EditDialogChildSubclassProc, 1, 0);
                s_sensitivePlaceholder = false;
            }

            SetFocus(hEdit);
        };
        switch(msg) {
            case WM_CREATE:
            {
                hCombo = CreateWindowExW(0, L"COMBOBOX", L"",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                    10, 10, 360, 250, hWnd, (HMENU)12, g_hInst, NULL);

                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F9EA  Auto");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4C4  Text");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F310  URL");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4C1  File");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F5A5  Command");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"\U0001F4BE  Binary");
                SendMessageW(hCombo, CB_SETCURSEL, (WPARAM)s_typeSelection, 0);

                hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", ((LPCWSTR)((LPCREATESTRUCT)lParam)->lpCreateParams), 
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL, 10, 40, 360, 70, hWnd, (HMENU)10, g_hInst, NULL);

                hTagsLabel = CreateWindowExW(0, L"STATIC", L"Tags (comma-separated):", WS_CHILD | WS_VISIBLE,
                    10, 115, 360, 18, hWnd, NULL, g_hInst, NULL);
                hTags = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", s_tagsResult.c_str(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 10, 135, 360, 22, hWnd, (HMENU)11, g_hInst, NULL);

                std::wstring deviceLine = L"Device ID: " + s_deviceIdResult;
                hDeviceLabel = CreateWindowExW(0, L"STATIC", deviceLine.c_str(), WS_CHILD | WS_VISIBLE,
                    10, 165, 360, 18, hWnd, NULL, g_hInst, NULL);

                std::wstring timeLine = L"Saved: " + s_timeResult;
                hTimeLabel = CreateWindowExW(0, L"STATIC", timeLine.c_str(), WS_CHILD | WS_VISIBLE,
                    10, 185, 360, 18, hWnd, NULL, g_hInst, NULL);

                hValidCheck = CreateWindowExW(0, L"BUTTON", L"Valid on any device", 
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                    10, 205, 360, 20, hWnd, (HMENU)13, g_hInst, NULL);
                if (s_validOnAnyDevice) {
                    SendMessageW(hValidCheck, BM_SETCHECK, BST_CHECKED, 0);
                }

                hSensitiveCheck = CreateWindowExW(0, L"BUTTON", L"Sensitive", 
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                    10, 228, 360, 20, hWnd, (HMENU)14, g_hInst, NULL);
                if (s_sensitive) {
                    SendMessageW(hSensitiveCheck, BM_SETCHECK, BST_CHECKED, 0);
                }

                hSave = CreateWindowExW(0, L"BUTTON", L"Save", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, 200, 240, 80, 30, hWnd, (HMENU)1, g_hInst, NULL);
                hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 290, 240, 80, 30, hWnd, (HMENU)2, g_hInst, NULL);

                if (g_hUiFont) {
                    SendMessageW(hCombo, WM_SETFONT, (WPARAM)(g_hEmojiFont ? g_hEmojiFont : g_hUiFont), TRUE);
                    SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTagsLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTags, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hDeviceLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hTimeLabel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hValidCheck, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hSensitiveCheck, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hSave, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                    SendMessageW(hCancel, WM_SETFONT, (WPARAM)g_hUiFont, TRUE);
                }

                SetWindowSubclass(hEdit, EditDialogChildSubclassProc, 1, 0);
                SetWindowSubclass(hTags, EditDialogChildSubclassProc, 2, 0);
                SetWindowSubclass(hCombo, EditDialogChildSubclassProc, 3, 0);
                SetWindowSubclass(hValidCheck, EditDialogChildSubclassProc, 4, 0);
                SetWindowSubclass(hSensitiveCheck, EditDialogChildSubclassProc, 5, 0);

                RECT rc;
                GetClientRect(hWnd, &rc);
                SendMessageW(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));

                if (s_sensitive) {
                    // Start in sensitive mode: never show plaintext.
                    SwapEditToSensitiveMode(true);
                } else {
                    SetFocus(hEdit);
                }
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

                SetWindowPos(hSensitiveCheck, NULL, margin, y + 4, w, 20, SWP_NOZORDER);
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
                if (LOWORD(wParam) == 10 && HIWORD(wParam) == EN_CHANGE) {
                    // If the sensitive field was showing the placeholder and the user started typing,
                    // treat it as a real replacement value.
                    if (s_sensitive && s_sensitivePlaceholder) {
                        int len = GetWindowTextLengthW(hEdit);
                        if (len != 3) {
                            s_sensitivePlaceholder = false;
                        } else {
                            wchar_t buf3[8] = {};
                            GetWindowTextW(hEdit, buf3, _countof(buf3));
                            if (wcscmp(buf3, L"***") != 0) {
                                s_sensitivePlaceholder = false;
                            }
                        }
                    }
                    break;
                }

                if (LOWORD(wParam) == 14 && HIWORD(wParam) == BN_CLICKED) {
                    bool nowSensitive = (SendMessageW(hSensitiveCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    if (nowSensitive) {
                        // Disallow creating an empty sensitive bookmark (must have a secret).
                        if (s_originalIdx == (size_t)-1) {
                            int len = GetWindowTextLengthW(hEdit);
                            if (len <= 0) {
                                MessageBoxW(hWnd,
                                    L"Sensitive bookmark content cannot be empty.",
                                    L"Sensitive",
                                    MB_OK | MB_TOPMOST | MB_SETFOREGROUND);
                                SendMessageW(hSensitiveCheck, BM_SETCHECK, BST_UNCHECKED, 0);
                                return 0;
                            }
                        }
                        SwapEditToSensitiveMode(true);
                        s_sensitive = true;
                    } else {
                        // Leaving sensitive mode: reveal the current value (since it's no longer sensitive).
                        // If the password field is still the placeholder, restore the original secret.
                        if (s_sensitive) {
                            int len = GetWindowTextLengthW(hEdit);
                            std::vector<wchar_t> buf(len + 1);
                            GetWindowTextW(hEdit, buf.data(), len + 1);
                            std::wstring curW = buf.data();
                            if (!(s_sensitivePlaceholder && curW == L"***")) {
                                s_originalSensitiveContentUtf8 = WideToUtf8(curW);
                            }
                        }
                        SwapEditToSensitiveMode(false);
                        s_sensitive = false;
                    }
                    break;
                }

                if (LOWORD(wParam) == 1) {
                    int len = GetWindowTextLengthW(hEdit);
                    std::vector<wchar_t> buf(len + 1);
                    GetWindowTextW(hEdit, &buf[0], len + 1);
                    std::wstring candidateW = &buf[0];
                    int candidateTypeSel = (int)SendMessageW(hCombo, CB_GETCURSEL, 0, 0);

                    bool isSensitiveNow = (SendMessageW(hSensitiveCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    if (isSensitiveNow) {
                        if (s_sensitivePlaceholder && candidateW == L"***" && !s_originalSensitiveContentUtf8.empty()) {
                            candidateW = Utf8ToWide(s_originalSensitiveContentUtf8);
                        }
                    }

                    int tlen = GetWindowTextLengthW(hTags);
                    std::vector<wchar_t> tbuf(tlen + 1);
                    GetWindowTextW(hTags, &tbuf[0], tlen + 1);
                    std::wstring tagsW = &tbuf[0];

                    if (g_bookmarkManager) {
                        g_bookmarkManager->loadIfChanged(true);
                        if (candidateTypeSel != 5) {
                            int exclude = (s_originalIdx == (size_t)-1) ? -1 : (int)s_originalIdx;
                            int dup = FindDuplicateIndex(WideToUtf8(candidateW), exclude);
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
                    s_sensitive = (SendMessageW(hSensitiveCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);

                    if (s_sensitive) {
                        // DPAPI encrypted payload is device/user-scoped.
                        s_validOnAnyDevice = false;

                        // Don't allow creating an empty sensitive bookmark.
                        if (s_originalIdx == (size_t)-1 && s_editResult.empty()) {
                            MessageBoxW(hWnd,
                                L"Sensitive bookmark content cannot be empty.",
                                L"Sensitive",
                                MB_OK | MB_TOPMOST | MB_SETFOREGROUND);
                            SetFocus(hEdit);
                            return 0;
                        }
                    }

                    s_saved = true;
                    DestroyWindow(hWnd);
                } else if (LOWORD(wParam) == 2) {
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

    HWND hEditWnd = NULL;
    {
        // Size window by desired CLIENT size to avoid clipped bottom buttons.
        int clientW = 440;
        int clientH = 350;
        RECT r{ 0, 0, clientW, clientH };
        DWORD style = WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU;
        DWORD exStyle = WS_EX_DLGMODALFRAME | WS_EX_TOPMOST | WS_EX_CONTROLPARENT;
        AdjustWindowRectEx(&r, style, FALSE, exStyle);
        int w = r.right - r.left;
        int h = r.bottom - r.top;
        int x = GetSystemMetrics(SM_CXSCREEN) / 2 - w / 2;
        int y = GetSystemMetrics(SM_CYSCREEN) / 2 - h / 2;

        hEditWnd = CreateWindowExW(exStyle, L"CheckmegEdit", L"Edit Bookmark",
            style,
            x, y, w, h,
            g_hSearchWnd, NULL, g_hInst, (LPVOID)currentContent.c_str());
    }

    // True modal behavior: prevent the search window from handling Esc, etc.
    if (g_hSearchWnd && IsWindow(g_hSearchWnd)) EnableWindow(g_hSearchWnd, FALSE);

    // Message loop for the modal window
    MSG msg;
    BOOL bRet;
    while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0) {
        if (bRet == -1) break;
        if (!IsWindow(hEditWnd)) break;
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE && (msg.hwnd == hEditWnd || IsChild(hEditWnd, msg.hwnd))) {
            DestroyWindow(hEditWnd);
            continue;
        }
        if (IsDialogMessageW(hEditWnd, &msg)) continue;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (g_hSearchWnd && IsWindow(g_hSearchWnd)) EnableWindow(g_hSearchWnd, TRUE);
    
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
        std::string newContent = WideToUtf8(s_editResult);
        std::string newDeviceId = WideToUtf8(s_deviceIdResult);
        bool newSensitive = s_sensitive;

        if (newSensitive) {
            // Keep existing sensitive content if user didn't enter a replacement.
            if (!isNew && newContent.empty() && !s_originalSensitiveContentUtf8.empty()) {
                newContent = s_originalSensitiveContentUtf8;
            }
        }

        // Preserve original device ID when editing
        if (isNew) {
            if (hasExplicitType && explicitType == BookmarkType::Binary) {
                std::string mime = b.mimeType.empty() ? "application/octet-stream" : b.mimeType;

                bool oldSuppress = g_bookmarkManager->suppressSyncCallbacks;
                g_bookmarkManager->suppressSyncCallbacks = true;
                g_bookmarkManager->addBinary(newContent, b.binaryData, mime, newDeviceId, s_validOnAnyDevice);
                size_t newIdx = g_bookmarkManager->bookmarks.size() - 1;
                g_bookmarkManager->bookmarks[newIdx].tags = tags;
                g_bookmarkManager->bookmarks[newIdx].validOnAnyDevice = s_validOnAnyDevice;
                g_bookmarkManager->bookmarks[newIdx].deviceId = newDeviceId;
                g_bookmarkManager->bookmarks[newIdx].sensitive = newSensitive;
                g_bookmarkManager->save();
                g_bookmarkManager->suppressSyncCallbacks = oldSuppress;
                if (!g_bookmarkManager->suppressSyncCallbacks && g_bookmarkManager->onUpsert) {
                    g_bookmarkManager->onUpsert(g_bookmarkManager->bookmarks[newIdx]);
                }
            } else {
                g_bookmarkManager->add(newContent, newDeviceId, s_validOnAnyDevice);
                size_t newIdx = g_bookmarkManager->bookmarks.size() - 1;
                g_bookmarkManager->update(newIdx, newContent, hasExplicitType, explicitType, tags, newDeviceId, s_validOnAnyDevice, newSensitive);
            }
        } else {
            g_bookmarkManager->update(originalIdx, newContent, hasExplicitType, explicitType, tags, newDeviceId, s_validOnAnyDevice, newSensitive);
        }

        if (outSavedContentUtf8) *outSavedContentUtf8 = newContent;
        // Refresh list
        wchar_t buffer[256];
        GetWindowTextW(g_hEdit, buffer, 256);
        UpdateSearchList(WideToUtf8(buffer));
    }

    return s_saved;
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
        std::string lowerContent;
        if (!b.sensitive) {
            lowerContent = b.content;
            std::transform(lowerContent.begin(), lowerContent.end(), lowerContent.begin(), ::tolower);
        }

        std::string lowerTags;
        for (const auto& t : b.tags) {
            if (!lowerTags.empty()) lowerTags += " ";
            std::string lt = t;
            std::transform(lt.begin(), lt.end(), lt.begin(), ::tolower);
            lowerTags += lt;
        }

        if (query.empty() || (!lowerContent.empty() && lowerContent.find(lowerQuery) != std::string::npos) || lowerTags.find(lowerQuery) != std::string::npos) {
            // Owner-drawn listbox: store content as string, render emoji icon in WM_DRAWITEM.
            std::wstring wDisplay = b.sensitive ? L"***" : Utf8ToWide(b.content);
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
