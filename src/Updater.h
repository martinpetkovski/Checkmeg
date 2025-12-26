#pragma once

#include <windows.h>
#include <shellapi.h>
#include <winhttp.h>
#include <string>
#include <vector>
#include <sstream>
#include <iostream>
#include "Version.h"

namespace Updater {

    struct UpdateInfo {
        bool available;
        std::string version;
        std::string downloadUrl;
    };

    // Minimal helpers copied/adapted to avoid dependencies
    static std::wstring Utf8ToWide(const std::string& str) {
        if (str.empty()) return L"";
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
        std::wstring wstrTo(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
        return wstrTo;
    }

    static bool CrackUrlHttps(const std::string& url, std::wstring* outHost, INTERNET_PORT* outPort, std::wstring* outPath) {
        URL_COMPONENTS parts{};
        parts.dwStructSize = sizeof(parts);
        parts.dwHostNameLength = 1;
        parts.dwUrlPathLength = 1;
        parts.dwExtraInfoLength = 1;

        std::wstring wurl = Utf8ToWide(url);
        if (!WinHttpCrackUrl(wurl.c_str(), (DWORD)wurl.size(), 0, &parts)) return false;

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

    static std::string SimpleGet(const std::string& url) {
        std::wstring host, path;
        INTERNET_PORT port;
        if (!CrackUrlHttps(url, &host, &port, &path)) return "";

        HINTERNET hSession = WinHttpOpen(L"Checkmeg/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!hSession) return "";

        HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), port, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return ""; }

        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return ""; }

        // GitHub API requires User-Agent
        std::wstring headers = L"User-Agent: Checkmeg\r\nAccept: application/vnd.github.v3+json\r\n";
        
        if (!WinHttpSendRequest(hRequest, headers.c_str(), (DWORD)-1L, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return "";
        }

        if (!WinHttpReceiveResponse(hRequest, nullptr)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return "";
        }

        std::string body;
        for (;;) {
            DWORD available = 0;
            if (!WinHttpQueryDataAvailable(hRequest, &available) || available == 0) break;
            std::vector<char> buf(available);
            DWORD read = 0;
            if (!WinHttpReadData(hRequest, buf.data(), available, &read) || read == 0) break;
            body.append(buf.data(), buf.data() + read);
        }

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return body;
    }

    static std::string JsonFindString(const std::string& json, const std::string& key) {
        std::string needle = "\"" + key + "\"";
        size_t pos = json.find(needle);
        if (pos == std::string::npos) return "";
        pos = json.find(':', pos + needle.size());
        if (pos == std::string::npos) return "";
        pos++;
        while (pos < json.size() && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\r' || json[pos] == '\n')) pos++;
        
        if (pos >= json.size() || json[pos] != '"') return "";
        pos++;
        
        std::string val;
        for (; pos < json.size(); ++pos) {
            if (json[pos] == '"') break;
            // Simple unescape for common chars
            if (json[pos] == '\\' && pos + 1 < json.size()) {
                pos++;
            }
            val += json[pos];
        }
        return val;
    }

    static UpdateInfo CheckForUpdates() {
        UpdateInfo info = {false, "", ""};
        std::string json = SimpleGet("https://api.github.com/repos/martinpetkovski/Checkmeg/releases/latest");
        if (json.empty()) return info;

        info.version = JsonFindString(json, "tag_name");
        
        // Find first asset download url
        // This is a very rough parse, assuming the first asset is the one we want or the structure is simple
        // We look for "browser_download_url" inside "assets" array.
        // Since we don't have a full parser, we'll just search for the key.
        // Ideally we should check if it ends in .zip
        
        size_t assetsPos = json.find("\"assets\"");
        if (assetsPos != std::string::npos) {
            size_t urlPos = json.find("\"browser_download_url\"", assetsPos);
            if (urlPos != std::string::npos) {
                // Extract value manually to reuse logic
                size_t colon = json.find(':', urlPos);
                if (colon != std::string::npos) {
                    size_t startQuote = json.find('"', colon);
                    if (startQuote != std::string::npos) {
                        size_t endQuote = json.find('"', startQuote + 1);
                        if (endQuote != std::string::npos) {
                            info.downloadUrl = json.substr(startQuote + 1, endQuote - startQuote - 1);
                        }
                    }
                }
            }
        }

        if (info.version.empty() || info.downloadUrl.empty()) return info;

        // Compare versions. Simple string compare for now, assuming format "1.YYMMDDHH"
        // If tag_name has 'v' prefix, strip it.
        std::string remote = info.version;
        if (!remote.empty() && remote[0] == 'v') remote = remote.substr(1);

        std::string local = CHECKMEG_VERSION;
        
        // If remote > local
        if (remote > local) {
            info.available = true;
        }

        return info;
    }

    static void TriggerUpdate(const std::string& downloadUrl) {
        char exePath[MAX_PATH];
        GetModuleFileNameA(NULL, exePath, MAX_PATH);
        std::string currentExe = exePath;

        // PowerShell script to download, unzip, replace, and restart
        // We use a temporary directory for extraction
        std::stringstream ps;
        ps << "$u='" << downloadUrl << "'; ";
        ps << "$d='$env:TEMP\\CheckmegUpdate.zip'; ";
        ps << "$x='$env:TEMP\\CheckmegUpdate'; ";
        ps << "$exe='" << currentExe << "'; ";
        ps << "Write-Host 'Downloading update...'; ";
        ps << "Invoke-RestMethod $u -OutFile $d; ";
        ps << "Write-Host 'Extracting...'; ";
        ps << "Expand-Archive -Path $d -DestinationPath $x -Force; ";
        ps << "$new = Get-ChildItem $x -Filter Checkmeg.exe -Recurse | Select-Object -First 1; ";
        ps << "if ($new) { ";
        ps << "  Write-Host 'Replacing executable...'; ";
        ps << "  Stop-Process -Name Checkmeg -Force -ErrorAction SilentlyContinue; ";
        ps << "  Start-Sleep -Seconds 1; ";
        ps << "  Move-Item -Path $new.FullName -Destination $exe -Force; ";
        ps << "  Start-Process $exe; ";
        ps << "}";

        std::string cmd = ps.str();
        
        // Convert to wide char for ShellExecute
        std::wstring wCmd = Utf8ToWide("-WindowStyle Hidden -Command \"" + cmd + "\"");
        
        ShellExecuteW(NULL, L"open", L"powershell.exe", wCmd.c_str(), NULL, SW_HIDE);
        exit(0);
    }
}
