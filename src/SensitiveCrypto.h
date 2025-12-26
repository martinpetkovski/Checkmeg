#pragma once

// Minimal DPAPI (Crypt32.dll) and CNG (Bcrypt.dll) encrypt/decrypt helpers.
// Loaded dynamically so we don't need to link against crypt32.lib or bcrypt.lib.

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif

#include <windows.h>
#include <wincrypt.h> // For DPAPI definitions if available

#include <string>
#include <vector>
#include <algorithm>
#include <random>

namespace SensitiveCrypto {

// DPAPI definitions
#ifndef CRYPTPROTECT_UI_FORBIDDEN
#define CRYPTPROTECT_UI_FORBIDDEN 0x1
#endif

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

// BCrypt definitions
#ifndef NTSTATUS
typedef LONG NTSTATUS;
#endif
#ifndef STATUS_SUCCESS
#define STATUS_SUCCESS ((NTSTATUS)0x00000000L)
#endif

#define BCRYPT_AES_ALGORITHM    L"AES"
#define BCRYPT_SHA256_ALGORITHM L"SHA256"
#define BCRYPT_CHAINING_MODE    L"ChainingMode"
#define BCRYPT_CHAIN_MODE_CBC   L"ChainingModeCBC"
#define BCRYPT_BLOCK_PADDING    0x00000001

using BCryptOpenAlgorithmProviderFn = NTSTATUS(WINAPI*)(PVOID*, LPCWSTR, LPCWSTR, ULONG);
using BCryptCloseAlgorithmProviderFn = NTSTATUS(WINAPI*)(PVOID, ULONG);
using BCryptSetPropertyFn = NTSTATUS(WINAPI*)(PVOID, LPCWSTR, PUCHAR, ULONG, ULONG);
using BCryptGenerateSymmetricKeyFn = NTSTATUS(WINAPI*)(PVOID, PVOID*, PUCHAR, ULONG, PUCHAR, ULONG, ULONG);
using BCryptDestroyKeyFn = NTSTATUS(WINAPI*)(PVOID);
using BCryptEncryptFn = NTSTATUS(WINAPI*)(PVOID, PUCHAR, ULONG, PVOID, PUCHAR, ULONG, PUCHAR, ULONG, ULONG*, ULONG);
using BCryptDecryptFn = NTSTATUS(WINAPI*)(PVOID, PUCHAR, ULONG, PVOID, PUCHAR, ULONG, PUCHAR, ULONG, ULONG*, ULONG);
using BCryptCreateHashFn = NTSTATUS(WINAPI*)(PVOID, PVOID*, PUCHAR, ULONG, PUCHAR, ULONG, ULONG);
using BCryptHashDataFn = NTSTATUS(WINAPI*)(PVOID, PUCHAR, ULONG, ULONG);
using BCryptFinishHashFn = NTSTATUS(WINAPI*)(PVOID, PUCHAR, ULONG, ULONG);
using BCryptDestroyHashFn = NTSTATUS(WINAPI*)(PVOID);

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

    if (valCount != 0) return false;
    *outBytes = std::move(out);
    return true;
}

static inline std::vector<unsigned char> BuildEntropyBytes(const std::string& userKeyUtf8) {
    std::string ent = "CheckmegSensitiveV1";
    if (!userKeyUtf8.empty()) {
        ent += ":";
        ent += userKeyUtf8;
    }
    return std::vector<unsigned char>(ent.begin(), ent.end());
}

// --- DPAPI Implementation ---

static inline bool EncryptDpapi(const std::string& plaintextUtf8, const std::string& entropyKeyUtf8, std::string* outCipherB64) {
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

    std::vector<unsigned char> entropyBytes = BuildEntropyBytes(entropyKeyUtf8);
    DataBlob entropy{};
    entropy.cbData = (DWORD)entropyBytes.size();
    entropy.pbData = entropyBytes.empty() ? nullptr : (BYTE*)entropyBytes.data();

    DataBlob out{};

    BOOL ok = pCryptProtectData(&in, L"Checkmeg Sensitive", &entropy, nullptr, nullptr, CRYPTPROTECT_UI_FORBIDDEN, &out);
    if (!ok || !out.pbData || out.cbData == 0) {
        FreeLibrary(hCrypt32);
        return false;
    }

    std::vector<unsigned char> bytes(out.pbData, out.pbData + out.cbData);
    LocalFree(out.pbData);
    FreeLibrary(hCrypt32);

    *outCipherB64 = Base64Encode(bytes);
    return !outCipherB64->empty();
}

static inline bool DecryptDpapi(const std::string& cipherB64, const std::string& entropyKeyUtf8, std::string* outPlaintextUtf8) {
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

// --- AES Implementation (CNG) ---

static inline bool AesEncrypt(const std::string& plaintext, const std::string& keyStr, std::string* outCipherB64) {
    HMODULE hBcrypt = LoadLibraryW(L"Bcrypt.dll");
    if (!hBcrypt) return false;

    auto pBCryptOpenAlgorithmProvider = (BCryptOpenAlgorithmProviderFn)GetProcAddress(hBcrypt, "BCryptOpenAlgorithmProvider");
    auto pBCryptCloseAlgorithmProvider = (BCryptCloseAlgorithmProviderFn)GetProcAddress(hBcrypt, "BCryptCloseAlgorithmProvider");
    auto pBCryptSetProperty = (BCryptSetPropertyFn)GetProcAddress(hBcrypt, "BCryptSetProperty");
    auto pBCryptGenerateSymmetricKey = (BCryptGenerateSymmetricKeyFn)GetProcAddress(hBcrypt, "BCryptGenerateSymmetricKey");
    auto pBCryptDestroyKey = (BCryptDestroyKeyFn)GetProcAddress(hBcrypt, "BCryptDestroyKey");
    auto pBCryptEncrypt = (BCryptEncryptFn)GetProcAddress(hBcrypt, "BCryptEncrypt");
    auto pBCryptCreateHash = (BCryptCreateHashFn)GetProcAddress(hBcrypt, "BCryptCreateHash");
    auto pBCryptHashData = (BCryptHashDataFn)GetProcAddress(hBcrypt, "BCryptHashData");
    auto pBCryptFinishHash = (BCryptFinishHashFn)GetProcAddress(hBcrypt, "BCryptFinishHash");
    auto pBCryptDestroyHash = (BCryptDestroyHashFn)GetProcAddress(hBcrypt, "BCryptDestroyHash");

    if (!pBCryptOpenAlgorithmProvider || !pBCryptEncrypt) {
        FreeLibrary(hBcrypt);
        return false;
    }

    // 1. Hash the keyStr to get 32 bytes (256 bits) key
    PVOID hHashAlg = nullptr;
    if (pBCryptOpenAlgorithmProvider(&hHashAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0) != STATUS_SUCCESS) { FreeLibrary(hBcrypt); return false; }
    
    PVOID hHash = nullptr;
    if (pBCryptCreateHash(hHashAlg, &hHash, nullptr, 0, nullptr, 0, 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    if (pBCryptHashData(hHash, (PUCHAR)keyStr.data(), (ULONG)keyStr.size(), 0) != STATUS_SUCCESS) { pBCryptDestroyHash(hHash); pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    std::vector<unsigned char> keyBytes(32);
    if (pBCryptFinishHash(hHash, keyBytes.data(), 32, 0) != STATUS_SUCCESS) { pBCryptDestroyHash(hHash); pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    pBCryptDestroyHash(hHash);
    pBCryptCloseAlgorithmProvider(hHashAlg, 0);

    // 2. Prepare AES
    PVOID hAesAlg = nullptr;
    if (pBCryptOpenAlgorithmProvider(&hAesAlg, BCRYPT_AES_ALGORITHM, nullptr, 0) != STATUS_SUCCESS) { FreeLibrary(hBcrypt); return false; }
    
    if (pBCryptSetProperty(hAesAlg, BCRYPT_CHAINING_MODE, (PUCHAR)BCRYPT_CHAIN_MODE_CBC, sizeof(BCRYPT_CHAIN_MODE_CBC), 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false; }

    PVOID hKey = nullptr;
    // Key object buffer
    std::vector<unsigned char> keyObj(1024); // Should be enough
    if (pBCryptGenerateSymmetricKey(hAesAlg, &hKey, keyObj.data(), (ULONG)keyObj.size(), keyBytes.data(), (ULONG)keyBytes.size(), 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false; }

    // 3. Generate IV (16 bytes)
    std::vector<unsigned char> iv(16);
    std::random_device rd;
    std::generate(iv.begin(), iv.end(), [&](){ return (unsigned char)rd(); });
    
    std::vector<unsigned char> ivCopy = iv; // Encrypt modifies IV

    // 4. Encrypt
    ULONG cbResult = 0;
    std::vector<unsigned char> plainBytes(plaintext.begin(), plaintext.end());
    
    // Get size
    if (pBCryptEncrypt(hKey, plainBytes.data(), (ULONG)plainBytes.size(), nullptr, ivCopy.data(), (ULONG)ivCopy.size(), nullptr, 0, &cbResult, BCRYPT_BLOCK_PADDING) != STATUS_SUCCESS) {
        pBCryptDestroyKey(hKey); pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false;
    }

    std::vector<unsigned char> cipherBytes(cbResult);
    ivCopy = iv; // Reset IV
    if (pBCryptEncrypt(hKey, plainBytes.data(), (ULONG)plainBytes.size(), nullptr, ivCopy.data(), (ULONG)ivCopy.size(), cipherBytes.data(), (ULONG)cipherBytes.size(), &cbResult, BCRYPT_BLOCK_PADDING) != STATUS_SUCCESS) {
        pBCryptDestroyKey(hKey); pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false;
    }

    pBCryptDestroyKey(hKey);
    pBCryptCloseAlgorithmProvider(hAesAlg, 0);
    FreeLibrary(hBcrypt);

    // 5. Format: IV + Ciphertext
    std::vector<unsigned char> finalBytes = iv;
    finalBytes.insert(finalBytes.end(), cipherBytes.begin(), cipherBytes.end());

    *outCipherB64 = Base64Encode(finalBytes);
    return true;
}

static inline bool AesDecrypt(const std::string& cipherB64, const std::string& keyStr, std::string* outPlaintext) {
    std::vector<unsigned char> data;
    if (!Base64Decode(cipherB64, &data)) return false;
    if (data.size() < 16) return false; // IV is 16 bytes

    std::vector<unsigned char> iv(data.begin(), data.begin() + 16);
    std::vector<unsigned char> cipherBytes(data.begin() + 16, data.end());

    HMODULE hBcrypt = LoadLibraryW(L"Bcrypt.dll");
    if (!hBcrypt) return false;

    auto pBCryptOpenAlgorithmProvider = (BCryptOpenAlgorithmProviderFn)GetProcAddress(hBcrypt, "BCryptOpenAlgorithmProvider");
    auto pBCryptCloseAlgorithmProvider = (BCryptCloseAlgorithmProviderFn)GetProcAddress(hBcrypt, "BCryptCloseAlgorithmProvider");
    auto pBCryptSetProperty = (BCryptSetPropertyFn)GetProcAddress(hBcrypt, "BCryptSetProperty");
    auto pBCryptGenerateSymmetricKey = (BCryptGenerateSymmetricKeyFn)GetProcAddress(hBcrypt, "BCryptGenerateSymmetricKey");
    auto pBCryptDestroyKey = (BCryptDestroyKeyFn)GetProcAddress(hBcrypt, "BCryptDestroyKey");
    auto pBCryptDecrypt = (BCryptDecryptFn)GetProcAddress(hBcrypt, "BCryptDecrypt");
    auto pBCryptCreateHash = (BCryptCreateHashFn)GetProcAddress(hBcrypt, "BCryptCreateHash");
    auto pBCryptHashData = (BCryptHashDataFn)GetProcAddress(hBcrypt, "BCryptHashData");
    auto pBCryptFinishHash = (BCryptFinishHashFn)GetProcAddress(hBcrypt, "BCryptFinishHash");
    auto pBCryptDestroyHash = (BCryptDestroyHashFn)GetProcAddress(hBcrypt, "BCryptDestroyHash");

    if (!pBCryptOpenAlgorithmProvider || !pBCryptDecrypt) {
        FreeLibrary(hBcrypt);
        return false;
    }

    // 1. Hash key
    PVOID hHashAlg = nullptr;
    if (pBCryptOpenAlgorithmProvider(&hHashAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0) != STATUS_SUCCESS) { FreeLibrary(hBcrypt); return false; }
    
    PVOID hHash = nullptr;
    if (pBCryptCreateHash(hHashAlg, &hHash, nullptr, 0, nullptr, 0, 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    if (pBCryptHashData(hHash, (PUCHAR)keyStr.data(), (ULONG)keyStr.size(), 0) != STATUS_SUCCESS) { pBCryptDestroyHash(hHash); pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    std::vector<unsigned char> keyBytes(32);
    if (pBCryptFinishHash(hHash, keyBytes.data(), 32, 0) != STATUS_SUCCESS) { pBCryptDestroyHash(hHash); pBCryptCloseAlgorithmProvider(hHashAlg, 0); FreeLibrary(hBcrypt); return false; }
    
    pBCryptDestroyHash(hHash);
    pBCryptCloseAlgorithmProvider(hHashAlg, 0);

    // 2. AES
    PVOID hAesAlg = nullptr;
    if (pBCryptOpenAlgorithmProvider(&hAesAlg, BCRYPT_AES_ALGORITHM, nullptr, 0) != STATUS_SUCCESS) { FreeLibrary(hBcrypt); return false; }
    
    if (pBCryptSetProperty(hAesAlg, BCRYPT_CHAINING_MODE, (PUCHAR)BCRYPT_CHAIN_MODE_CBC, sizeof(BCRYPT_CHAIN_MODE_CBC), 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false; }

    PVOID hKey = nullptr;
    std::vector<unsigned char> keyObj(1024);
    if (pBCryptGenerateSymmetricKey(hAesAlg, &hKey, keyObj.data(), (ULONG)keyObj.size(), keyBytes.data(), (ULONG)keyBytes.size(), 0) != STATUS_SUCCESS) { pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false; }

    // 3. Decrypt
    ULONG cbResult = 0;
    // Get size
    if (pBCryptDecrypt(hKey, cipherBytes.data(), (ULONG)cipherBytes.size(), nullptr, iv.data(), (ULONG)iv.size(), nullptr, 0, &cbResult, BCRYPT_BLOCK_PADDING) != STATUS_SUCCESS) {
        pBCryptDestroyKey(hKey); pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false;
    }

    std::vector<unsigned char> plainBytes(cbResult);
    // Reset IV
    std::vector<unsigned char> ivCopy = std::vector<unsigned char>(data.begin(), data.begin() + 16);
    
    if (pBCryptDecrypt(hKey, cipherBytes.data(), (ULONG)cipherBytes.size(), nullptr, ivCopy.data(), (ULONG)ivCopy.size(), plainBytes.data(), (ULONG)plainBytes.size(), &cbResult, BCRYPT_BLOCK_PADDING) != STATUS_SUCCESS) {
        pBCryptDestroyKey(hKey); pBCryptCloseAlgorithmProvider(hAesAlg, 0); FreeLibrary(hBcrypt); return false;
    }

    plainBytes.resize(cbResult);
    outPlaintext->assign(plainBytes.begin(), plainBytes.end());

    pBCryptDestroyKey(hKey);
    pBCryptCloseAlgorithmProvider(hAesAlg, 0);
    FreeLibrary(hBcrypt);
    return true;
}

// --- Unified Interface ---

static inline bool HasLegacyInlineMarker(const std::string& content) {
    return content.rfind("enc:v1:", 0) == 0;
}

static inline bool HasAesMarker(const std::string& content) {
    return content.rfind("enc:aes:", 0) == 0;
}

static inline bool EncryptSensitive(const std::string& plaintext, const std::string& key, std::string* outCipher) {
    std::string b64;
    if (!AesEncrypt(plaintext, key, &b64)) return false;
    *outCipher = "enc:aes:" + b64;
    return true;
}

static inline bool DecryptSensitive(const std::string& cipher, const std::string& key, std::string* outPlaintext) {
    if (HasAesMarker(cipher)) {
        std::string b64 = cipher.substr(8); // "enc:aes:" length
        return AesDecrypt(b64, key, outPlaintext);
    } else if (HasLegacyInlineMarker(cipher)) {
        std::string b64 = cipher.substr(7); // "enc:v1:" length
        return DecryptDpapi(b64, key, outPlaintext);
    } else {
        // Try raw DPAPI (legacy)
        return DecryptDpapi(cipher, key, outPlaintext);
    }
}

// Deprecated aliases for compatibility if needed, but we should update callers.
// We redirect them to the new unified functions to ensure new data is AES encrypted.
static inline bool EncryptUtf8ToBase64Dpapi(const std::string& plaintext, const std::string& key, std::string* outCipher) {
    return EncryptSensitive(plaintext, key, outCipher);
}

static inline bool DecryptUtf8FromBase64Dpapi(const std::string& cipher, const std::string& key, std::string* outPlaintext) {
    return DecryptSensitive(cipher, key, outPlaintext);
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
