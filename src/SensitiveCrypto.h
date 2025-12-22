#pragma once

// Minimal DPAPI (Crypt32.dll) encrypt/decrypt helpers, loaded dynamically
// so we don't need to link against crypt32.lib.

#define _WIN32_WINNT 0x0600
#include <windows.h>

#include <string>
#include <vector>

namespace SensitiveCrypto {

struct DataBlob {
    DWORD cbData;
    BYTE* pbData;
};

using CryptProtectDataFn = BOOL(WINAPI*)(
    const DataBlob* pDataIn,
    LPCWSTR szDataDescr,
    const DataBlob* pOptionalEntropy,
    PVOID pvReserved,
    PVOID pPromptStruct,
    DWORD dwFlags,
    DataBlob* pDataOut);

using CryptUnprotectDataFn = BOOL(WINAPI*)(
    const DataBlob* pDataIn,
    LPWSTR* ppszDataDescr,
    const DataBlob* pOptionalEntropy,
    PVOID pvReserved,
    PVOID pPromptStruct,
    DWORD dwFlags,
    DataBlob* pDataOut);

static inline std::string Base64Encode(const std::vector<unsigned char>& bytes) {
    static const char* k = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    out.reserve(((bytes.size() + 2) / 3) * 4);

    size_t i = 0;
    while (i + 2 < bytes.size()) {
        unsigned int n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        out.push_back(k[(n >> 18) & 63]);
        out.push_back(k[(n >> 12) & 63]);
        out.push_back(k[(n >> 6) & 63]);
        out.push_back(k[n & 63]);
        i += 3;
    }

    if (i < bytes.size()) {
        unsigned int n = (bytes[i] << 16);
        out.push_back(k[(n >> 18) & 63]);
        if (i + 1 < bytes.size()) {
            n |= (bytes[i + 1] << 8);
            out.push_back(k[(n >> 12) & 63]);
            out.push_back(k[(n >> 6) & 63]);
            out.push_back('=');
        } else {
            out.push_back(k[(n >> 12) & 63]);
            out.push_back('=');
            out.push_back('=');
        }
    }

    return out;
}

static inline bool Base64Decode(const std::string& s, std::vector<unsigned char>* outBytes) {
    if (!outBytes) return false;

    auto decodeChar = [](char c) -> int {
        if (c >= 'A' && c <= 'Z') return c - 'A';
        if (c >= 'a' && c <= 'z') return c - 'a' + 26;
        if (c >= '0' && c <= '9') return c - '0' + 52;
        if (c == '+') return 62;
        if (c == '/') return 63;
        return -1;
    };

    std::vector<unsigned char> out;
    out.reserve((s.size() / 4) * 3);

    int val[4];
    int valCount = 0;

    for (char c : s) {
        if (c == '\r' || c == '\n' || c == ' ' || c == '\t') continue;
        if (c == '=') {
            val[valCount++] = -2;
        } else {
            int v = decodeChar(c);
            if (v < 0) return false;
            val[valCount++] = v;
        }

        if (valCount == 4) {
            if (val[0] < 0 || val[1] < 0) return false;
            unsigned int n = ((unsigned int)val[0] << 18) | ((unsigned int)val[1] << 12);

            out.push_back((unsigned char)((n >> 16) & 0xFF));

            if (val[2] == -2) {
                // == padding => done
                valCount = 0;
                break;
            }
            if (val[2] < 0) return false;
            n |= ((unsigned int)val[2] << 6);
            out.push_back((unsigned char)((n >> 8) & 0xFF));

            if (val[3] == -2) {
                valCount = 0;
                break;
            }
            if (val[3] < 0) return false;
            n |= (unsigned int)val[3];
            out.push_back((unsigned char)(n & 0xFF));

            valCount = 0;
        }
    }

    if (valCount != 0) {
        // Non-multiple of 4 base64 is invalid.
        return false;
    }

    *outBytes = std::move(out);
    return true;
}

static inline std::vector<unsigned char> BuildEntropyBytes(const std::string& userKeyUtf8) {
    // Domain separation + optional user-provided secret.
    // NOTE: userKeyUtf8 should be consistent across runs to decrypt old data.
    std::string ent = "CheckmegSensitiveV1";
    if (!userKeyUtf8.empty()) {
        ent += ":";
        ent += userKeyUtf8;
    }
    return std::vector<unsigned char>(ent.begin(), ent.end());
}

static inline bool EncryptUtf8ToBase64Dpapi(const std::string& plaintextUtf8, const std::string& entropyKeyUtf8, std::string* outCipherB64) {
    if (!outCipherB64) return false;
    outCipherB64->clear();

    HMODULE hCrypt32 = LoadLibraryW(L"Crypt32.dll");
    if (!hCrypt32) return false;

    auto pCryptProtectData = (CryptProtectDataFn)GetProcAddress(hCrypt32, "CryptProtectData");
    if (!pCryptProtectData) {
        FreeLibrary(hCrypt32);
        return false;
    }

    DataBlob in{};
    in.cbData = (DWORD)plaintextUtf8.size();
    in.pbData = (BYTE*)(plaintextUtf8.empty() ? nullptr : (BYTE*)plaintextUtf8.data());

    // Optional entropy (we use domain separation + caller-provided secret)
    std::vector<unsigned char> entropyBytes = BuildEntropyBytes(entropyKeyUtf8);
    DataBlob entropy{};
    entropy.cbData = (DWORD)entropyBytes.size();
    entropy.pbData = entropyBytes.empty() ? nullptr : (BYTE*)entropyBytes.data();

    DataBlob out{};
    const DWORD CRYPTPROTECT_UI_FORBIDDEN = 0x1;

    BOOL ok = pCryptProtectData(&in, L"Checkmeg Sensitive", &entropy, nullptr, nullptr, CRYPTPROTECT_UI_FORBIDDEN, &out);
    if (!ok || !out.pbData || out.cbData == 0) {
        FreeLibrary(hCrypt32);
        return false;
    }

    std::vector<unsigned char> bytes(out.pbData, out.pbData + out.cbData);
    LocalFree(out.pbData);

    *outCipherB64 = Base64Encode(bytes);

    FreeLibrary(hCrypt32);
    return !outCipherB64->empty();
}

static inline bool DecryptUtf8FromBase64Dpapi(const std::string& cipherB64, const std::string& entropyKeyUtf8, std::string* outPlaintextUtf8) {
    if (!outPlaintextUtf8) return false;
    outPlaintextUtf8->clear();

    std::vector<unsigned char> cipherBytes;
    if (!Base64Decode(cipherB64, &cipherBytes)) return false;

    HMODULE hCrypt32 = LoadLibraryW(L"Crypt32.dll");
    if (!hCrypt32) return false;

    auto pCryptUnprotectData = (CryptUnprotectDataFn)GetProcAddress(hCrypt32, "CryptUnprotectData");
    if (!pCryptUnprotectData) {
        FreeLibrary(hCrypt32);
        return false;
    }

    DataBlob in{};
    in.cbData = (DWORD)cipherBytes.size();
    in.pbData = cipherBytes.empty() ? nullptr : (BYTE*)cipherBytes.data();

    auto tryDecryptWithEntropy = [&](const std::vector<unsigned char>& entropyBytes) -> bool {
        DataBlob entropy{};
        entropy.cbData = (DWORD)entropyBytes.size();
        entropy.pbData = entropyBytes.empty() ? nullptr : (BYTE*)entropyBytes.data();

        DataBlob out{};
        const DWORD CRYPTPROTECT_UI_FORBIDDEN = 0x1;

        BOOL ok = pCryptUnprotectData(&in, nullptr, &entropy, nullptr, nullptr, CRYPTPROTECT_UI_FORBIDDEN, &out);
        if (!ok || !out.pbData) {
            return false;
        }
        outPlaintextUtf8->assign((const char*)out.pbData, (size_t)out.cbData);
        LocalFree(out.pbData);
        return true;
    };

    if (!tryDecryptWithEntropy(BuildEntropyBytes(entropyKeyUtf8))) {
        FreeLibrary(hCrypt32);
        return false;
    }

    FreeLibrary(hCrypt32);
    return true;
}

static inline bool HasLegacyInlineMarker(const std::string& content) {
    return content.rfind("enc:v1:", 0) == 0;
}

static inline std::string MakeLegacyInlineMarker(const std::string& cipherB64) {
    return std::string("enc:v1:") + cipherB64;
}

static inline bool TryParseLegacyInlineMarker(const std::string& content, std::string* outCipherB64) {
    if (!outCipherB64) return false;
    outCipherB64->clear();
    if (!HasLegacyInlineMarker(content)) return false;
    *outCipherB64 = content.substr(std::string("enc:v1:").size());
    return !outCipherB64->empty();
}

} // namespace SensitiveCrypto
