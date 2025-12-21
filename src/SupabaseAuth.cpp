#define _WIN32_WINNT 0x0600
#include "SupabaseAuth.h"
#include "SupabaseConfig.h"

#include <windows.h>
#include <winhttp.h>
#include <shlobj.h>

#include <vector>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <ctime>

#pragma comment(lib, "winhttp.lib")

namespace {

struct HttpResponse {
    DWORD status = 0;
    std::string body;
};

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return std::wstring();
    int sizeNeeded = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    std::wstring out(sizeNeeded, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), out.data(), sizeNeeded);
    return out;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return std::string();
    int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string out(sizeNeeded, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), out.data(), sizeNeeded, nullptr, nullptr);
    return out;
}

static bool IsSupabaseConfigured() {
    return !SupabaseConfig::SUPABASE_URL.empty() && !SupabaseConfig::SUPABASE_ANON_KEY.empty();
}

static std::string JsonEscape(const std::string& s) {
    std::string out;
    out.reserve(s.size() + 16);
    for (char c : s) {
        switch (c) {
            case '\\': out += "\\\\"; break;
            case '"': out += "\\\""; break;
            case '\n': out += "\\n"; break;
            case '\r': out += "\\r"; break;
            case '\t': out += "\\t"; break;
            default: out += c; break;
        }
    }
    return out;
}

static std::string JsonUnescape(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        char c = s[i];
        if (c != '\\' || i + 1 >= s.size()) {
            out.push_back(c);
            continue;
        }
        char n = s[++i];
        switch (n) {
            case '\\': out.push_back('\\'); break;
            case '"': out.push_back('"'); break;
            case 'n': out.push_back('\n'); break;
            case 'r': out.push_back('\r'); break;
            case 't': out.push_back('\t'); break;
            default: out.push_back(n); break;
        }
    }
    return out;
}

static bool JsonFindString(const std::string& json, const std::string& key, std::string* out) {
    if (!out) return false;
    std::string needle = "\"" + key + "\"";
    size_t pos = json.find(needle);
    if (pos == std::string::npos) return false;
    pos = json.find(':', pos + needle.size());
    if (pos == std::string::npos) return false;

    // Skip whitespace
    pos++;
    while (pos < json.size() && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n')) pos++;
    if (pos >= json.size() || json[pos] != '"') return false;
    pos++; // after opening quote

    std::string raw;
    for (; pos < json.size(); ++pos) {
        char c = json[pos];
        if (c == '"') {
            *out = JsonUnescape(raw);
            return true;
        }
        if (c == '\\' && pos + 1 < json.size()) {
            raw.push_back('\\');
            raw.push_back(json[pos + 1]);
            pos++;
            continue;
        }
        raw.push_back(c);
    }
    return false;
}

static bool JsonFindInt64(const std::string& json, const std::string& key, std::int64_t* out) {
    if (!out) return false;
    std::string needle = "\"" + key + "\"";
    size_t pos = json.find(needle);
    if (pos == std::string::npos) return false;
    pos = json.find(':', pos + needle.size());
    if (pos == std::string::npos) return false;
    pos++;
    while (pos < json.size() && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n')) pos++;

    bool neg = false;
    if (pos < json.size() && json[pos] == '-') { neg = true; pos++; }

    std::int64_t value = 0;
    bool any = false;
    while (pos < json.size() && json[pos] >= '0' && json[pos] <= '9') {
        any = true;
        value = value * 10 + (json[pos] - '0');
        pos++;
    }
    if (!any) return false;
    *out = neg ? -value : value;
    return true;
}

static bool CrackUrlHttps(const std::string& urlUtf8, std::wstring* outHost, INTERNET_PORT* outPort, std::wstring* outPath) {
    if (!outHost || !outPort || !outPath) return false;

    std::wstring url = Utf8ToWide(urlUtf8);
    URL_COMPONENTS parts{};
    parts.dwStructSize = sizeof(parts);

    wchar_t host[256] = {};
    wchar_t path[2048] = {};
    parts.lpszHostName = host;
    parts.dwHostNameLength = _countof(host);
    parts.lpszUrlPath = path;
    parts.dwUrlPathLength = _countof(path);

    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &parts)) return false;

    if (parts.nScheme != INTERNET_SCHEME_HTTPS) return false;

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

static HttpResponse HttpPostJson(const std::string& url, const std::string& jsonBody, const std::vector<std::pair<std::string, std::string>>& headers) {
    HttpResponse resp;

    std::wstring host;
    INTERNET_PORT port = INTERNET_DEFAULT_HTTPS_PORT;
    std::wstring path;
    if (!CrackUrlHttps(url, &host, &port, &path)) {
        resp.status = 0;
        resp.body = "Invalid URL (expected https://...)";
        return resp;
    }

    HINTERNET hSession = WinHttpOpen(L"Checkmeg/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) {
        resp.status = 0;
        resp.body = "WinHttpOpen failed";
        return resp;
    }

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), port, 0);
    if (!hConnect) {
        resp.status = 0;
        resp.body = "WinHttpConnect failed";
        WinHttpCloseHandle(hSession);
        return resp;
    }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
    if (!hRequest) {
        resp.status = 0;
        resp.body = "WinHttpOpenRequest failed";
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return resp;
    }

    // Timeouts (ms)
    WinHttpSetTimeouts(hRequest, 8000, 8000, 8000, 8000);

    // Build headers
    // NOTE: Supabase expects UTF-8 JSON. We send UTF-8 bytes in the request body.
    std::wstring allHeaders = L"Content-Type: application/json; charset=utf-8\r\n";
    for (const auto& kv : headers) {
        allHeaders += Utf8ToWide(kv.first + ": " + kv.second + "\r\n");
    }

    BOOL ok = WinHttpSendRequest(
        hRequest,
        allHeaders.c_str(),
        (DWORD)-1L,
        (LPVOID)(jsonBody.empty() ? nullptr : jsonBody.data()),
        (DWORD)jsonBody.size(),
        (DWORD)jsonBody.size(),
        0);

    if (!ok) {
        resp.status = 0;
        resp.body = "WinHttpSendRequest failed";
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return resp;
    }

    ok = WinHttpReceiveResponse(hRequest, nullptr);
    if (!ok) {
        resp.status = 0;
        resp.body = "WinHttpReceiveResponse failed";
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return resp;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX,
        &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
    resp.status = status;

    std::string body;
    for (;;) {
        DWORD available = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &available) || available == 0) break;
        std::vector<char> buf(available);
        DWORD read = 0;
        if (!WinHttpReadData(hRequest, buf.data(), available, &read) || read == 0) break;
        body.append(buf.data(), buf.data() + read);
    }
    resp.body = body;

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    return resp;
}

static std::string BestEffortExtractError(const std::string& json) {
    std::string err;
    if (JsonFindString(json, "error_description", &err) && !err.empty()) return err;
    if (JsonFindString(json, "message", &err) && !err.empty()) return err;
    if (JsonFindString(json, "msg", &err) && !err.empty()) return err;
    if (JsonFindString(json, "error", &err) && !err.empty()) return err;
    return "Request failed";
}

static std::int64_t NowUnix() {
    return (std::int64_t)std::time(nullptr);
}

} // namespace

SupabaseAuth::SupabaseAuth() = default;

std::string SupabaseAuth::BuildJsonEmailPassword(const std::string& email, const std::string& password) {
    return std::string("{") +
        "\"email\":\"" + JsonEscape(email) + "\"," +
        "\"password\":\"" + JsonEscape(password) + "\"" +
        "}";
}

std::string SupabaseAuth::BuildJsonRefreshToken(const std::string& refreshToken) {
    return std::string("{") +
        "\"refresh_token\":\"" + JsonEscape(refreshToken) + "\"" +
        "}";
}

std::wstring SupabaseAuth::GetSessionFilePathW() const {
    PWSTR appData = nullptr;
    std::wstring dir;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, nullptr, &appData)) && appData) {
        dir = appData;
        CoTaskMemFree(appData);
    } else {
        dir = L".";
    }
    if (!dir.empty() && dir.back() != L'\\') dir += L"\\";
    dir += L"Checkmeg";
    if (!dir.empty() && dir.back() != L'\\') dir += L"\\";
    return dir + L"session.json";
}

bool SupabaseAuth::EnsureSessionDirExists() const {
    std::wstring filePath = GetSessionFilePathW();
    size_t pos = filePath.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return false;
    std::wstring dir = filePath.substr(0, pos);

    // CreateDirectory succeeds if exists.
    if (CreateDirectoryW(dir.c_str(), nullptr)) return true;
    DWORD err = GetLastError();
    return err == ERROR_ALREADY_EXISTS;
}

bool SupabaseAuth::SaveSessionToDisk(std::string* outError) const {
    if (!EnsureSessionDirExists()) {
        if (outError) *outError = "Failed to create session directory";
        return false;
    }

    std::wstring pathW = GetSessionFilePathW();
    std::ofstream f(pathW, std::ios::binary | std::ios::trunc);
    if (!f) {
        if (outError) *outError = "Failed to open session file for writing";
        return false;
    }

    std::ostringstream ss;
    ss << "{";
    ss << "\"email\":\"" << JsonEscape(session_.email) << "\",";
    ss << "\"refresh_token\":\"" << JsonEscape(session_.refreshToken) << "\"";
    ss << "}";

    const std::string payload = ss.str();
    f.write(payload.data(), payload.size());
    if (!f) {
        if (outError) *outError = "Failed to write session file";
        return false;
    }
    return true;
}

void SupabaseAuth::ClearSessionOnDisk() const {
    std::wstring pathW = GetSessionFilePathW();
    DeleteFileW(pathW.c_str());
}

bool SupabaseAuth::LoadSessionFromDisk() {
    session_ = SupabaseSession{};

    std::wstring pathW = GetSessionFilePathW();
    std::ifstream f(pathW, std::ios::binary);
    if (!f) return false;

    std::ostringstream ss;
    ss << f.rdbuf();
    std::string json = ss.str();

    std::string email;
    std::string refresh;
    if (!JsonFindString(json, "email", &email)) return false;
    if (!JsonFindString(json, "refresh_token", &refresh)) return false;

    session_.email = email;
    session_.refreshToken = refresh;
    session_.loggedIn = !session_.refreshToken.empty();
    return session_.loggedIn;
}

bool SupabaseAuth::RefreshWithToken(const std::string& refreshToken, std::string* outError) {
    if (!IsSupabaseConfigured()) {
        if (outError) *outError = "Supabase not configured (set SUPABASE_URL and SUPABASE_ANON_KEY in SupabaseConfig.h)";
        return false;
    }

    std::vector<std::pair<std::string, std::string>> headers;
    headers.push_back({"apikey", SupabaseConfig::SUPABASE_ANON_KEY});

    HttpResponse resp = HttpPostJson(SupabaseConfig::EndpointTokenRefresh(), BuildJsonRefreshToken(refreshToken), headers);
    if (resp.status < 200 || resp.status >= 300) {
        if (outError) *outError = BestEffortExtractError(resp.body);
        return false;
    }

    std::string access;
    std::string refreshNew;
    std::int64_t expiresIn = 0;

    if (!JsonFindString(resp.body, "access_token", &access) || access.empty()) {
        if (outError) *outError = "Refresh response missing access_token";
        return false;
    }

    // Supabase usually rotates refresh_token; handle both cases.
    if (!JsonFindString(resp.body, "refresh_token", &refreshNew) || refreshNew.empty()) {
        refreshNew = refreshToken;
    }

    (void)JsonFindInt64(resp.body, "expires_in", &expiresIn);

    session_.accessToken = access;
    session_.refreshToken = refreshNew;
    session_.expiresAtUnix = (expiresIn > 0) ? (NowUnix() + expiresIn) : 0;
    session_.loggedIn = true;

    std::string saveErr;
    SaveSessionToDisk(&saveErr);
    return true;
}

bool SupabaseAuth::TryRestoreOrRefresh(std::string* outError) {
    if (!LoadSessionFromDisk()) {
        if (outError) *outError = "No saved session";
        return false;
    }

    std::string err;
    if (RefreshWithToken(session_.refreshToken, &err)) {
        return true;
    }

    // If refresh fails, clear local session.
    session_ = SupabaseSession{};
    ClearSessionOnDisk();
    if (outError) *outError = err;
    return false;
}

bool SupabaseAuth::SignInWithPassword(const std::string& email, const std::string& password, std::string* outError) {
    if (!IsSupabaseConfigured()) {
        if (outError) *outError = "Supabase not configured (set SUPABASE_URL and SUPABASE_ANON_KEY in SupabaseConfig.h)";
        return false;
    }

    std::vector<std::pair<std::string, std::string>> headers;
    headers.push_back({"apikey", SupabaseConfig::SUPABASE_ANON_KEY});

    HttpResponse resp = HttpPostJson(SupabaseConfig::EndpointTokenPassword(), BuildJsonEmailPassword(email, password), headers);
    if (resp.status < 200 || resp.status >= 300) {
        if (outError) *outError = BestEffortExtractError(resp.body);
        return false;
    }

    std::string access;
    std::string refresh;
    std::int64_t expiresIn = 0;

    if (!JsonFindString(resp.body, "access_token", &access) || access.empty()) {
        if (outError) *outError = "Login response missing access_token";
        return false;
    }
    if (!JsonFindString(resp.body, "refresh_token", &refresh) || refresh.empty()) {
        if (outError) *outError = "Login response missing refresh_token";
        return false;
    }
    (void)JsonFindInt64(resp.body, "expires_in", &expiresIn);

    session_.loggedIn = true;
    session_.email = email;
    session_.accessToken = access;
    session_.refreshToken = refresh;
    session_.expiresAtUnix = (expiresIn > 0) ? (NowUnix() + expiresIn) : 0;

    std::string saveErr;
    if (!SaveSessionToDisk(&saveErr)) {
        // Still consider login successful, but surface warning.
        if (outError) *outError = "Logged in, but failed to save session: " + saveErr;
    }
    return true;
}

bool SupabaseAuth::SignUpWithPassword(const std::string& email, const std::string& password, std::string* outError) {
    if (!IsSupabaseConfigured()) {
        if (outError) *outError = "Supabase not configured (set SUPABASE_URL and SUPABASE_ANON_KEY in SupabaseConfig.h)";
        return false;
    }

    std::vector<std::pair<std::string, std::string>> headers;
    headers.push_back({"apikey", SupabaseConfig::SUPABASE_ANON_KEY});

    HttpResponse resp = HttpPostJson(SupabaseConfig::EndpointSignUp(), BuildJsonEmailPassword(email, password), headers);
    if (resp.status < 200 || resp.status >= 300) {
        if (outError) *outError = BestEffortExtractError(resp.body);
        return false;
    }

    // Some Supabase setups return session tokens immediately; others require email confirmation.
    std::string access;
    std::string refresh;
    std::int64_t expiresIn = 0;

    if (JsonFindString(resp.body, "access_token", &access) && JsonFindString(resp.body, "refresh_token", &refresh) && !refresh.empty()) {
        (void)JsonFindInt64(resp.body, "expires_in", &expiresIn);
        session_.loggedIn = true;
        session_.email = email;
        session_.accessToken = access;
        session_.refreshToken = refresh;
        session_.expiresAtUnix = (expiresIn > 0) ? (NowUnix() + expiresIn) : 0;

        std::string saveErr;
        SaveSessionToDisk(&saveErr);
        return true;
    }

    // Signup succeeded but no session.
    session_ = SupabaseSession{};
    if (outError) *outError = "Sign up succeeded. If email confirmation is enabled, confirm your email then log in.";
    return true;
}

void SupabaseAuth::Logout() {
    if (IsSupabaseConfigured() && session_.loggedIn && !session_.accessToken.empty()) {
        std::vector<std::pair<std::string, std::string>> headers;
        headers.push_back({"apikey", SupabaseConfig::SUPABASE_ANON_KEY});
        headers.push_back({"Authorization", std::string("Bearer ") + session_.accessToken});
        (void)HttpPostJson(SupabaseConfig::EndpointLogout(), "{}", headers);
    }

    session_ = SupabaseSession{};
    ClearSessionOnDisk();
}
