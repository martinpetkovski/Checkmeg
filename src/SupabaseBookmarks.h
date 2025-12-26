#pragma once

#include "Bookmark.h"
#include "SupabaseAuth.h"
#include "SupabaseConfig.h"

#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <winhttp.h>

#include <string>
#include <vector>
#include <utility>
#include <sstream>
#include <algorithm>

#include "SensitiveCrypto.h"
#include "SupabaseConfig.h"

#pragma comment(lib, "winhttp.lib")
class SupabaseBookmarks {
public:
    explicit SupabaseBookmarks(SupabaseAuth* auth = nullptr) : auth_(auth) {}

    void SetAuth(SupabaseAuth* auth) { auth_ = auth; }

    bool FetchAll(std::vector<Bookmark>* outBookmarks, std::string* outError, bool includeBinary = true) {
        if (!outBookmarks) return false;
        outBookmarks->clear();

        if (!auth_ || !auth_->IsLoggedIn()) {
            if (outError) *outError = "Not logged in";
            return false;
        }

        // Prefer explicit sensitive + contentEnc columns when present.
        std::string url = SupabaseConfig::EndpointBookmarksTable() +
            "?select=id,type,typeExplicit,content,contentEnc,sensitive,binaryData,mimeType,tags,timestamp,lastUsed,deviceId,validOnAnyDevice";
        if (!includeBinary) {
            // Schema uses lowercase: type in ('text','url','file','command','binary')
            url += "&type=neq.binary";
        }

        HttpResponse resp = HttpRequest("GET", url, "", BuildAuthHeaders());
        if (resp.status == 401 || resp.status == 403) {
            std::string ignored;
            auth_->TryRestoreOrRefresh(&ignored);
            resp = HttpRequest("GET", url, "", BuildAuthHeaders());
        }

        // Back-compat: if the project hasn't migrated the table yet (no contentEnc/sensitive columns), retry
        // with the older select list.
        if (resp.status >= 400 && resp.status < 500) {
            if (resp.body.find("contentEnc") != std::string::npos || resp.body.find("sensitive") != std::string::npos) {
                supportsSensitiveColumns_ = false;
                std::string urlOld = SupabaseConfig::EndpointBookmarksTable() +
                    "?select=id,type,typeExplicit,content,binaryData,mimeType,tags,timestamp,lastUsed,deviceId,validOnAnyDevice";
                if (!includeBinary) {
                    urlOld += "&type=neq.binary";
                }
                resp = HttpRequest("GET", urlOld, "", BuildAuthHeaders());
                if (resp.status == 401 || resp.status == 403) {
                    std::string ignored;
                    auth_->TryRestoreOrRefresh(&ignored);
                    resp = HttpRequest("GET", urlOld, "", BuildAuthHeaders());
                }
            }
        }

        if (resp.status < 200 || resp.status >= 300) {
            if (outError) *outError = BestEffortExtractError(resp.body);
            return false;
        }

        std::vector<std::string> objects = SplitTopLevelJsonObjects(resp.body);
        std::vector<Bookmark> result;
        result.reserve(objects.size());

        for (const auto& obj : objects) {
            Bookmark b;
            b.id = JsonGetString(obj, "id");
            if (b.id.empty()) continue;

            std::string type = JsonGetString(obj, "type");
            b.type = ParseType(type);
            b.typeExplicit = JsonGetBool(obj, "typeExplicit", false);

            const std::string rawContent = JsonGetString(obj, "content");
            const std::string rawContentEnc = supportsSensitiveColumns_ ? JsonGetString(obj, "contentEnc") : std::string();
            const bool rawSensitive = supportsSensitiveColumns_ ? JsonGetBool(obj, "sensitive", false) : false;

            b.deviceId = JsonGetString(obj, "deviceId");
            b.validOnAnyDevice = JsonGetBool(obj, "validOnAnyDevice", true);

            b.sensitive = rawSensitive;

            // Legacy inline marker support even without separate columns.
            if (!b.sensitive && SensitiveCrypto::HasLegacyInlineMarker(rawContent)) {
                b.sensitive = true;
            }

            if (b.sensitive) {
                std::string cipherB64;
                if (!rawContentEnc.empty()) {
                    cipherB64 = rawContentEnc;
                } else {
                    (void)SensitiveCrypto::TryParseLegacyInlineMarker(rawContent, &cipherB64);
                }

                std::string plain;
                if (!cipherB64.empty() && SensitiveCrypto::DecryptUtf8FromBase64Dpapi(cipherB64, SupabaseConfig::SENSITIVE_CRYPTO_KEY, &plain)) {
                    b.content = plain;
                    // If AES, it's portable. If DPAPI, it's device scoped.
                    if (SensitiveCrypto::HasAesMarker(cipherB64)) {
                        b.validOnAnyDevice = true;
                    } else {
                        b.validOnAnyDevice = false;
                    }
                } else {
                    b.content.clear();
                    b.validOnAnyDevice = false;
                }
            } else {
                b.content = rawContent;
            }
            b.binaryData = JsonGetString(obj, "binaryData");
            b.mimeType = JsonGetString(obj, "mimeType");
            b.tags = JsonGetStringArray(obj, "tags");
            b.timestamp = (std::time_t)JsonGetInt64(obj, "timestamp", 0);
            b.lastUsed = (std::time_t)JsonGetInt64(obj, "lastUsed", b.timestamp);

            b.binaryDataLoaded = true;
            b.originalFileIndex = (size_t)-1;

            result.push_back(std::move(b));
        }

        std::sort(result.begin(), result.end(), [](const Bookmark& a, const Bookmark& b) {
            if (a.lastUsed == b.lastUsed) return a.timestamp > b.timestamp;
            return a.lastUsed > b.lastUsed;
        });

        *outBookmarks = std::move(result);
        return true;
    }

    bool Upsert(const Bookmark& b, std::string* outError) {
        if (!auth_ || !auth_->IsLoggedIn()) {
            if (outError) *outError = "Not logged in";
            return false;
        }
        if (b.id.empty()) {
            if (outError) *outError = "Bookmark missing id";
            return false;
        }

        std::string url = SupabaseConfig::EndpointBookmarksTable() + "?on_conflict=id";
        std::string body = "[" + BuildBookmarkJsonObject(b, supportsSensitiveColumns_) + "]";

        auto headers = BuildAuthHeaders();
        headers.push_back({"Prefer", "resolution=merge-duplicates,return=minimal"});

        HttpResponse resp = HttpRequest("POST", url, body, headers);
        if (resp.status == 401 || resp.status == 403) {
            std::string ignored;
            auth_->TryRestoreOrRefresh(&ignored);
            resp = HttpRequest("POST", url, body, headers);
        }

        // Back-compat: if sensitive/contentEnc columns don't exist, retry without them.
        if (resp.status >= 400 && resp.status < 500) {
            if (resp.body.find("contentEnc") != std::string::npos || resp.body.find("sensitive") != std::string::npos) {
                supportsSensitiveColumns_ = false;
                body = "[" + BuildBookmarkJsonObject(b, false) + "]";
                resp = HttpRequest("POST", url, body, headers);
                if (resp.status == 401 || resp.status == 403) {
                    std::string ignored;
                    auth_->TryRestoreOrRefresh(&ignored);
                    resp = HttpRequest("POST", url, body, headers);
                }
            }
        }

        if (resp.status < 200 || resp.status >= 300) {
            if (outError) *outError = BestEffortExtractError(resp.body);
            return false;
        }
        return true;
    }

    bool DeleteById(const std::string& id, std::string* outError) {
        if (!auth_ || !auth_->IsLoggedIn()) {
            if (outError) *outError = "Not logged in";
            return false;
        }
        if (id.empty()) {
            if (outError) *outError = "Missing id";
            return false;
        }

        std::string url = SupabaseConfig::EndpointBookmarksTable() + "?id=eq." + UrlEscapeComponent(id);

        HttpResponse resp = HttpRequest("DELETE", url, "", BuildAuthHeaders());
        if (resp.status == 401 || resp.status == 403) {
            std::string ignored;
            auth_->TryRestoreOrRefresh(&ignored);
            resp = HttpRequest("DELETE", url, "", BuildAuthHeaders());
        }

        if (resp.status < 200 || resp.status >= 300) {
            if (outError) *outError = BestEffortExtractError(resp.body);
            return false;
        }
        return true;
    }

private:
    struct HttpResponse {
        DWORD status = 0;
        std::string body;
    };

    SupabaseAuth* auth_ = nullptr;
    bool supportsSensitiveColumns_ = true;

    static std::wstring Utf8ToWide(const std::string& s) {
        if (s.empty()) return std::wstring();
        int sizeNeeded = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
        std::wstring out(sizeNeeded, 0);
        MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), out.data(), sizeNeeded);
        return out;
    }

    static bool CrackUrlHttps(const std::string& urlUtf8, std::wstring* outHost, INTERNET_PORT* outPort, std::wstring* outPath) {
        if (!outHost || !outPort || !outPath) return false;

        std::wstring url = Utf8ToWide(urlUtf8);
        URL_COMPONENTS parts{};
        parts.dwStructSize = sizeof(parts);

        wchar_t host[256] = {};
        wchar_t path[4096] = {};
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

    static HttpResponse HttpRequest(const std::string& method, const std::string& url, const std::string& bodyUtf8,
        const std::vector<std::pair<std::string, std::string>>& headers) {
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

        std::wstring methodW = Utf8ToWide(method);
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, methodW.c_str(), path.c_str(), nullptr, WINHTTP_NO_REFERER,
            WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
        if (!hRequest) {
            resp.status = 0;
            resp.body = "WinHttpOpenRequest failed";
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return resp;
        }

        WinHttpSetTimeouts(hRequest, 8000, 8000, 8000, 8000);

        std::wstring allHeaders;
        allHeaders += L"Accept: application/json\r\n";
        if (!bodyUtf8.empty()) {
            allHeaders += L"Content-Type: application/json; charset=utf-8\r\n";
        }
        for (const auto& kv : headers) {
            allHeaders += Utf8ToWide(kv.first + ": " + kv.second + "\r\n");
        }

        BOOL ok = WinHttpSendRequest(
            hRequest,
            allHeaders.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : allHeaders.c_str(),
            (DWORD)-1L,
            (LPVOID)(bodyUtf8.empty() ? nullptr : bodyUtf8.data()),
            (DWORD)bodyUtf8.size(),
            (DWORD)bodyUtf8.size(),
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

    std::vector<std::pair<std::string, std::string>> BuildAuthHeaders() const {
        std::vector<std::pair<std::string, std::string>> headers;
        headers.push_back({"apikey", SupabaseConfig::SUPABASE_ANON_KEY});
        headers.push_back({"Authorization", std::string("Bearer ") + auth_->Session().accessToken});
        return headers;
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

    static std::string BestEffortExtractError(const std::string& json) {
        std::string m;
        if (JsonFindString(json, "message", &m) && !m.empty()) return m;
        if (JsonFindString(json, "msg", &m) && !m.empty()) return m;
        if (JsonFindString(json, "error_description", &m) && !m.empty()) return m;
        if (JsonFindString(json, "hint", &m) && !m.empty()) return m;
        if (JsonFindString(json, "details", &m) && !m.empty()) return m;
        return "Request failed";
    }

    static bool JsonFindString(const std::string& json, const std::string& key, std::string* out) {
        if (!out) return false;
        std::string needle = "\"" + key + "\"";
        size_t pos = json.find(needle);
        if (pos == std::string::npos) return false;
        pos = json.find(':', pos + needle.size());
        if (pos == std::string::npos) return false;

        pos++;
        while (pos < json.size() && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n')) pos++;
        if (pos >= json.size() || json[pos] != '"') return false;
        pos++;

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

    static bool JsonFindBool(const std::string& json, const std::string& key, bool* out) {
        if (!out) return false;
        std::string needle = "\"" + key + "\"";
        size_t pos = json.find(needle);
        if (pos == std::string::npos) return false;
        pos = json.find(':', pos + needle.size());
        if (pos == std::string::npos) return false;
        pos++;
        while (pos < json.size() && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n')) pos++;
        if (json.compare(pos, 4, "true") == 0) { *out = true; return true; }
        if (json.compare(pos, 5, "false") == 0) { *out = false; return true; }
        return false;
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

    static std::string JsonGetString(const std::string& jsonObj, const std::string& key) {
        std::string v;
        (void)JsonFindString(jsonObj, key, &v);
        return v;
    }

    static std::int64_t JsonGetInt64(const std::string& jsonObj, const std::string& key, std::int64_t def) {
        std::int64_t v = 0;
        if (!JsonFindInt64(jsonObj, key, &v)) return def;
        return v;
    }

    static bool JsonGetBool(const std::string& jsonObj, const std::string& key, bool def) {
        bool v = false;
        if (!JsonFindBool(jsonObj, key, &v)) return def;
        return v;
    }

    static std::vector<std::string> JsonGetStringArray(const std::string& jsonObj, const std::string& key) {
        std::vector<std::string> out;
        std::string needle = "\"" + key + "\"";
        size_t pos = jsonObj.find(needle);
        if (pos == std::string::npos) return out;
        pos = jsonObj.find(':', pos + needle.size());
        if (pos == std::string::npos) return out;
        pos++;
        while (pos < jsonObj.size() && (jsonObj[pos] == ' ' || jsonObj[pos] == '\t' || jsonObj[pos] == '\r' || jsonObj[pos] == '\n')) pos++;
        if (pos >= jsonObj.size() || jsonObj[pos] != '[') return out;
        bool inStr = false;
        bool esc = false;
        std::string cur;
        for (; pos < jsonObj.size(); ++pos) {
            char c = jsonObj[pos];
            if (!inStr) {
                if (c == ']') break;
                if (c == '"') {
                    inStr = true;
                    cur.clear();
                }
                continue;
            }

            if (esc) {
                cur.push_back('\\');
                cur.push_back(c);
                esc = false;
                continue;
            }
            if (c == '\\') {
                esc = true;
                continue;
            }
            if (c == '"') {
                out.push_back(JsonUnescape(cur));
                inStr = false;
                continue;
            }
            cur.push_back(c);
        }
        return out;
    }

    static std::vector<std::string> SplitTopLevelJsonObjects(const std::string& jsonArray) {
        std::vector<std::string> out;
        int depth = 0;
        bool inStr = false;
        bool esc = false;
        size_t start = std::string::npos;

        for (size_t i = 0; i < jsonArray.size(); ++i) {
            char c = jsonArray[i];
            if (inStr) {
                if (esc) {
                    esc = false;
                } else if (c == '\\') {
                    esc = true;
                } else if (c == '"') {
                    inStr = false;
                }
                continue;
            }

            if (c == '"') { inStr = true; continue; }
            if (c == '{') {
                if (depth == 0) start = i;
                depth++;
                continue;
            }
            if (c == '}') {
                depth--;
                if (depth == 0 && start != std::string::npos) {
                    out.push_back(jsonArray.substr(start, i - start + 1));
                    start = std::string::npos;
                }
                continue;
            }
        }
        return out;
    }

    static BookmarkType ParseType(const std::string& t) {
        if (t == "url") return BookmarkType::URL;
        if (t == "file") return BookmarkType::File;
        if (t == "command") return BookmarkType::Command;
        if (t == "binary") return BookmarkType::Binary;
        return BookmarkType::Text;
    }

    static std::string BuildBookmarkJsonObject(const Bookmark& b, bool includeSensitiveColumns) {
        std::ostringstream ss;
        ss << "{";
        ss << "\"id\":\"" << JsonEscape(b.id) << "\",";

        ss << "\"type\":\"" << JsonEscape(TypeToString(b.type)) << "\",";
        ss << "\"typeExplicit\":" << (b.typeExplicit ? "true" : "false") << ",";

        if (includeSensitiveColumns) {
            ss << "\"sensitive\":" << (b.sensitive ? "true" : "false") << ",";
        }

        if (b.sensitive) {
            std::string cipherB64;
            (void)SensitiveCrypto::EncryptUtf8ToBase64Dpapi(b.content, SupabaseConfig::SENSITIVE_CRYPTO_KEY, &cipherB64);

            if (includeSensitiveColumns) {
                ss << "\"contentEnc\":\"" << JsonEscape(cipherB64) << "\",";
                ss << "\"content\":\"\",";
            } else {
                // Schema didn't migrate: store ciphertext inline in content.
                ss << "\"content\":\"" << JsonEscape(SensitiveCrypto::MakeLegacyInlineMarker(cipherB64)) << "\",";
            }
        } else {
            ss << "\"content\":\"" << JsonEscape(b.content) << "\",";
        }

        ss << "\"tags\":[";
        for (size_t i = 0; i < b.tags.size(); ++i) {
            if (i) ss << ",";
            ss << "\"" << JsonEscape(b.tags[i]) << "\"";
        }
        ss << "],";

        ss << "\"timestamp\":" << (long long)b.timestamp << ",";
        ss << "\"lastUsed\":" << (long long)b.lastUsed << ",";

        ss << "\"deviceId\":\"" << JsonEscape(b.deviceId) << "\",";
        ss << "\"validOnAnyDevice\":" << (b.validOnAnyDevice ? "true" : "false") << ",";

        if (b.type == BookmarkType::Binary) {
            ss << "\"mimeType\":\"" << JsonEscape(b.mimeType) << "\",";
            ss << "\"binaryData\":\"" << JsonEscape(b.binaryData) << "\"";
        } else {
            ss << "\"mimeType\":\"\",";
            ss << "\"binaryData\":\"\"";
        }

        ss << "}";
        return ss.str();
    }

    static std::string TypeToString(BookmarkType t) {
        switch (t) {
            case BookmarkType::URL: return "url";
            case BookmarkType::File: return "file";
            case BookmarkType::Command: return "command";
            case BookmarkType::Binary: return "binary";
            default: return "text";
        }
    }

    static std::string UrlEscapeComponent(const std::string& s) {
        static const char* hex = "0123456789ABCDEF";
        std::string out;
        out.reserve(s.size() + 8);
        for (unsigned char c : s) {
            if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == '~') {
                out.push_back((char)c);
            } else {
                out.push_back('%');
                out.push_back(hex[(c >> 4) & 0xF]);
                out.push_back(hex[c & 0xF]);
            }
        }
        return out;
    }
};
