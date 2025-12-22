#pragma once
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <ctime>
#include <algorithm>
#include <map>
#include <limits>
#include <filesystem>
#include <random>
#include <functional>

#include "SensitiveCrypto.h"
#include "SupabaseConfig.h"

enum class BookmarkType {
    Text,
    URL,
    File,
    Command,
    Binary
};

struct Bookmark {
    std::string id;
    BookmarkType type;
    bool typeExplicit = false;
    std::string content;
    bool sensitive = false;
    std::string binaryData;
    std::string mimeType;
    std::vector<std::string> tags;
    std::time_t timestamp;
    std::time_t lastUsed;
    std::string deviceId;
    bool validOnAnyDevice = true;
    bool binaryDataLoaded = true;
    size_t originalFileIndex = -1;
};

class BookmarkManager {
public:
    std::vector<Bookmark> bookmarks;
    std::string filePath;

    bool useLocalFile = true;

    bool suppressSyncCallbacks = false;
    std::function<void(const Bookmark&)> onUpsert;
    std::function<void(const Bookmark&)> onDelete;

    bool hasLastWriteTime = false;
    std::filesystem::file_time_type lastWriteTime;

    BookmarkManager(const std::string& path, bool autoLoad = true) : filePath(path) {
        if (autoLoad && useLocalFile) load(true);
    }

    void SetUseLocalFile(bool enable) {
        useLocalFile = enable;
    }

    void ReplaceAll(const std::vector<Bookmark>& newBookmarks) {
        bool old = suppressSyncCallbacks;
        suppressSyncCallbacks = true;
        bookmarks = newBookmarks;
        suppressSyncCallbacks = old;
    }

    void loadIfChanged(bool skipBinary = false) {
        if (!useLocalFile) return;
        std::error_code ec;
        auto wt = std::filesystem::last_write_time(filePath, ec);
        if (ec) {
            load(skipBinary);
            return;
        }

        if (!hasLastWriteTime || wt != lastWriteTime) {
            load(skipBinary);
        }
    }

    void add(const std::string& content, const std::string& deviceId = "", bool validOnAnyDevice = true) {
        Bookmark b;
        b.id = newUuid();
        b.content = content;
        b.timestamp = std::time(nullptr);
        b.lastUsed = b.timestamp;
        b.typeExplicit = false;
        b.type = detectType(content);
        b.tags.clear();
        b.deviceId = deviceId;
        b.validOnAnyDevice = validOnAnyDevice;
        bookmarks.push_back(b);
        save();

        if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks.back());
    }

    void addBinary(const std::string& content, const std::string& data, const std::string& mime, const std::string& deviceId = "", bool validOnAnyDevice = true) {
        Bookmark b;
        b.id = newUuid();
        b.content = content;
        b.binaryData = data;
        b.mimeType = mime;
        b.timestamp = std::time(nullptr);
        b.lastUsed = b.timestamp;
        b.typeExplicit = true;
        b.type = BookmarkType::Binary;
        b.tags.clear();
        b.deviceId = deviceId;
        b.validOnAnyDevice = validOnAnyDevice;
        bookmarks.push_back(b);
        save();

        if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks.back());
    }

    void remove(size_t index) {
        if (index < bookmarks.size()) {
            Bookmark removed = bookmarks[index];
            bookmarks.erase(bookmarks.begin() + index);
            save();

            if (!suppressSyncCallbacks && onDelete && !removed.id.empty()) onDelete(removed);
        }
    }

    void update(size_t index, const std::string& content) {
        if (index < bookmarks.size()) {
            bookmarks[index].content = content;
            if (!bookmarks[index].typeExplicit) {
                bookmarks[index].type = detectType(content);
            }
            bookmarks[index].timestamp = std::time(nullptr);
            save();

            if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks[index]);
        }
    }

    void update(size_t index, const std::string& content, bool hasExplicitType, BookmarkType explicitType) {
        if (index < bookmarks.size()) {
            bookmarks[index].content = content;
            if (hasExplicitType) {
                bookmarks[index].typeExplicit = true;
                bookmarks[index].type = explicitType;
            } else {
                bookmarks[index].typeExplicit = false;
                bookmarks[index].type = detectType(content);
            }
            bookmarks[index].timestamp = std::time(nullptr);
            save();

            if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks[index]);
        }
    }

    void update(size_t index, const std::string& content, bool hasExplicitType, BookmarkType explicitType, const std::vector<std::string>& tags, const std::string& deviceId, bool validOnAnyDevice) {
        if (index < bookmarks.size()) {
            bookmarks[index].content = content;
            bookmarks[index].tags = tags;
            if (hasExplicitType) {
                bookmarks[index].typeExplicit = true;
                bookmarks[index].type = explicitType;
            } else {
                bookmarks[index].typeExplicit = false;
                bookmarks[index].type = detectType(content);
            }
            bookmarks[index].timestamp = std::time(nullptr);
            bookmarks[index].deviceId = deviceId;
            bookmarks[index].validOnAnyDevice = validOnAnyDevice;
            save();

            if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks[index]);
        }
    }

    void update(size_t index, const std::string& content, bool hasExplicitType, BookmarkType explicitType, const std::vector<std::string>& tags, const std::string& deviceId, bool validOnAnyDevice, bool sensitive) {
        if (index < bookmarks.size()) {
            bookmarks[index].content = content;
            bookmarks[index].tags = tags;
            if (hasExplicitType) {
                bookmarks[index].typeExplicit = true;
                bookmarks[index].type = explicitType;
            } else {
                bookmarks[index].typeExplicit = false;
                bookmarks[index].type = detectType(content);
            }
            bookmarks[index].timestamp = std::time(nullptr);
            bookmarks[index].deviceId = deviceId;
            bookmarks[index].sensitive = sensitive;
            // Sensitive is stored via DPAPI encryption; treat it as device/user scoped.
            bookmarks[index].validOnAnyDevice = sensitive ? false : validOnAnyDevice;
            save();

            if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks[index]);
        }
    }

    void updateBinaryData(size_t index, const std::string& data) {
        if (index < bookmarks.size()) {
            bookmarks[index].binaryData = data;
            bookmarks[index].binaryDataLoaded = true;
            save();

            if (!suppressSyncCallbacks && onUpsert) onUpsert(bookmarks[index]);
        }
    }
void updateLastUsed(size_t index) {
        if (index < bookmarks.size()) {
            bookmarks[index].lastUsed = std::time(nullptr);
            save();
        }
    }
    void ensureBinaryDataLoaded(size_t index) {
        if (!useLocalFile) return;
        if (index < bookmarks.size() && bookmarks[index].type == BookmarkType::Binary && !bookmarks[index].binaryDataLoaded) {
            std::string data = loadBinaryDataForIndex(bookmarks[index].originalFileIndex);
            if (!data.empty()) {
                bookmarks[index].binaryData = data;
                bookmarks[index].binaryDataLoaded = true;
            }
        }
    }    
    BookmarkType detectType(const std::string& content) {
        if (content.find("http://") == 0 || content.find("https://") == 0 || content.find("www.") == 0) {
            return BookmarkType::URL;
        } else if (content.length() > 3 && content[1] == ':' && (content[2] == '\\' || content[2] == '/')) {
            return BookmarkType::File;
        } else {
            return BookmarkType::Text;
        }
    }

    void save() {
        if (!useLocalFile) return;
        bool needToLoad = false;
        for (const auto& b : bookmarks) {
            if (b.type == BookmarkType::Binary && !b.binaryDataLoaded) {
                needToLoad = true;
                break;
            }
        }

        std::map<size_t, std::string> binaryMap;
        if (needToLoad) {
            binaryMap = loadBinaryDataMap();
        }

        std::ofstream out(filePath);
        out << "[\n";
        for (size_t i = 0; i < bookmarks.size(); ++i) {
            const auto& b = bookmarks[i];
            out << "  {\n";
            out << "    \"id\": \"" << escape(b.id) << "\",\n";
            out << "    \"type\": \"" << typeToString(b.type) << "\",\n";
            out << "    \"typeExplicit\": " << (b.typeExplicit ? "true" : "false") << ",\n";
            out << "    \"sensitive\": " << (b.sensitive ? "true" : "false") << ",\n";
            if (b.type == BookmarkType::Binary) {
                out << "    \"mimeType\": \"" << escape(b.mimeType) << "\",\n";
                std::string dataToWrite = b.binaryData;
                if (!b.binaryDataLoaded) {
                    if (binaryMap.count(b.originalFileIndex)) {
                        dataToWrite = binaryMap[b.originalFileIndex];
                    }
                }
                out << "    \"binaryData\": \"" << escape(dataToWrite) << "\",\n";
            }
            out << "    \"lastUsed\": " << b.lastUsed << ",\n";
            out << "    \"tags\": [\n";
            for (size_t t = 0; t < b.tags.size(); ++t) {
                out << "      \"" << escape(b.tags[t]) << "\"" << (t < b.tags.size() - 1 ? "," : "") << "\n";
            }
            out << "    ],\n";
            if (b.sensitive) {
                std::string cipherB64;
                if (SensitiveCrypto::EncryptUtf8ToBase64Dpapi(b.content, SupabaseConfig::SENSITIVE_CRYPTO_KEY, &cipherB64)) {
                    out << "    \"contentEnc\": \"" << escape(cipherB64) << "\",\n";
                } else {
                    // Worst case: avoid writing plaintext if encryption fails.
                    out << "    \"contentEnc\": \"\",\n";
                }
                out << "    \"content\": \"\",\n";
            } else {
                out << "    \"content\": \"" << escape(b.content) << "\",\n";
            }
            out << "    \"timestamp\": " << b.timestamp << ",\n";
            out << "    \"deviceId\": \"" << escape(b.deviceId) << "\",\n";
            out << "    \"validOnAnyDevice\": " << (b.validOnAnyDevice ? "true" : "false") << "\n";
            out << "  }" << (i < bookmarks.size() - 1 ? "," : "") << "\n";
        }
        out << "]\n";

        std::error_code ec;
        lastWriteTime = std::filesystem::last_write_time(filePath, ec);
        hasLastWriteTime = !ec;
    }

    void load(bool skipBinary = false) {
        if (!useLocalFile) return;
        bookmarks.clear();
        std::ifstream in(filePath);
        std::string line;
        Bookmark current;
        bool inside = false;
        bool insideTags = false;
        size_t currentIndex = 0;

        std::string pendingContentEnc;

        while (smartGetLine(in, line, skipBinary)) {
            line.erase(0, line.find_first_not_of(" \t"));
            current.lastUsed = 0;
                
            if (line.find("{") == 0) {
                inside = true;
                current = Bookmark();
                current.id.clear();
                current.tags.clear();
                current.validOnAnyDevice = true; 
                current.deviceId = "";
                current.originalFileIndex = currentIndex;
                current.binaryDataLoaded = true;
                current.sensitive = false;
                pendingContentEnc.clear();
                insideTags = false;
            } else if (line.find("}") == 0) {
                if (inside) {
                    if (current.id.empty()) current.id = newUuid();

                    if (current.sensitive) {
                        // If we have an encrypted payload, decrypt it into current.content.
                        std::string cipherB64;
                        if (!pendingContentEnc.empty()) {
                            cipherB64 = pendingContentEnc;
                        } else {
                            // Legacy inline marker support in case content was stored as enc:v1:<b64>
                            (void)SensitiveCrypto::TryParseLegacyInlineMarker(current.content, &cipherB64);
                        }

                        if (!cipherB64.empty()) {
                            std::string plain;
                            if (SensitiveCrypto::DecryptUtf8FromBase64Dpapi(cipherB64, SupabaseConfig::SENSITIVE_CRYPTO_KEY, &plain)) {
                                current.content = plain;
                            } else {
                                // Can't decrypt on this machine/user.
                                current.content.clear();
                                current.validOnAnyDevice = false;
                            }
                        } else {
                            // Marked sensitive but no ciphertext; avoid retaining any accidental plaintext.
                            current.content.clear();
                        }

                        // DPAPI is device/user scoped; don't treat as global.
                        current.validOnAnyDevice = false;
                    }

                    bookmarks.push_back(current);
                    inside = false;
                    insideTags = false;
                    currentIndex++;
                }
            } else if (inside) {
                if (insideTags) {
                    if (line.find("]") == 0) {
                        insideTags = false;
                        continue;
                    }
                    std::string val = extractQuotedString(line);
                    if (!val.empty()) current.tags.push_back(unescape(val));
                    continue;
                }

                if (line.find("\"id\":") == 0) {
                    current.id = unescape(extractValue(line));
                } else if (line.find("\"type\":") == 0) {
                    std::string val = extractValue(line);
                    if (val == "url") current.type = BookmarkType::URL;
                    else if (val == "file") current.type = BookmarkType::File;
                    else if (val == "command") current.type = BookmarkType::Command;
                    else if (val == "binary") current.type = BookmarkType::Binary;
                    else current.type = BookmarkType::Text;
                } else if (line.find("\"typeExplicit\":") == 0) {
                    if (line.find("true") != std::string::npos) current.typeExplicit = true;
                    else current.typeExplicit = false;
                } else if (line.find("\"mimeType\":") == 0) {
                    current.mimeType = unescape(extractValue(line));
                } else if (line.find("\"binaryData\":") == 0) {
                    if (!skipBinary) {
                        current.binaryData = unescape(extractValue(line));
                        current.binaryDataLoaded = true;
                    } else {
                        current.binaryDataLoaded = false;
                    }
                } else if (line.find("\"sensitive\":") == 0) {
                    current.sensitive = (line.find("true") != std::string::npos);
                } else if (line.find("\"tags\":") == 0) {
                    insideTags = true;
                } else if (line.find("\"contentEnc\":") == 0) {
                    pendingContentEnc = unescape(extractValue(line));
                } else if (line.find("\"content\":") == 0) {
                    current.content = unescape(extractValue(line));
                } else if (line.find("\"timestamp\":") == 0) {
                    std::string val = line.substr(line.find(":") + 1);
                    if (val.find(",") != std::string::npos) val = val.substr(0, val.find(","));
                    val.erase(0, val.find_first_not_of(" \t\r\n"));
                    val.erase(val.find_last_not_of(" \t\r\n") + 1);
                    current.timestamp = std::stoll(val);
                } else if (line.find("\"lastUsed\":") == 0) {
                    std::string val = line.substr(line.find(":") + 1);
                    if (val.find(",") != std::string::npos) val = val.substr(0, val.find(","));
                    val.erase(0, val.find_first_not_of(" \t\r\n"));
                    val.erase(val.find_last_not_of(" \t\r\n") + 1);
                    current.lastUsed = std::stoll(val);
                } else if (line.find("\"deviceId\":") == 0) {
                    current.deviceId = unescape(extractValue(line));
                } else if (line.find("\"validOnAnyDevice\":") == 0) {
                    if (line.find("true") != std::string::npos) current.validOnAnyDevice = true;
                    else current.validOnAnyDevice = false;
                }
            }
        }

        std::error_code ec;
        lastWriteTime = std::filesystem::last_write_time(filePath, ec);
        hasLastWriteTime = !ec;
    }

private:
    std::string newUuid() {
        static thread_local std::mt19937_64 rng([] {
            std::random_device rd;
            std::seed_seq seq{ rd(), rd(), rd(), rd(), rd(), rd(), rd(), rd() };
            return std::mt19937_64(seq);
        }());

        std::uniform_int_distribution<uint64_t> dist;
        uint64_t a = dist(rng);
        uint64_t b = dist(rng);

        b = (b & 0x3FFFFFFFFFFFFFFFULL) | 0x8000000000000000ULL;
        a = (a & 0xFFFFFFFFFFFF0FFFULL) | 0x0000000000004000ULL;

        auto hex = [](uint64_t v, int n) {
            static const char* k = "0123456789abcdef";
            std::string s;
            s.reserve(n);
            for (int i = (n - 1) * 4; i >= 0; i -= 4) s.push_back(k[(v >> i) & 0xF]);
            return s;
        };

        std::string out;
        out.reserve(36);
        out += hex((a >> 32) & 0xFFFFFFFFULL, 8);
        out += "-";
        out += hex((a >> 16) & 0xFFFFULL, 4);
        out += "-";
        out += hex(a & 0xFFFFULL, 4);
        out += "-";
        out += hex((b >> 48) & 0xFFFFULL, 4);
        out += "-";
        out += hex(b & 0xFFFFFFFFFFFFULL, 12);
        return out;
    }

    std::string typeToString(BookmarkType t) {
        switch(t) {
            case BookmarkType::URL: return "url";
            case BookmarkType::File: return "file";
            case BookmarkType::Command: return "command";
            case BookmarkType::Binary: return "binary";
            default: return "text";
        }
    }

    std::string escape(const std::string& s) {
        std::string out;
        for (char c : s) {
            if (c == '"') out += "\\\"";
            else if (c == '\\') out += "\\\\";
            else if (c == '\n') out += "\\n";
            else out += c;
        }
        return out;
    }

    std::string unescape(const std::string& s) {
        std::string out;
        for (size_t i = 0; i < s.length(); ++i) {
            if (s[i] == '\\' && i + 1 < s.length()) {
                char next = s[i+1];
                if (next == '"') out += '"';
                else if (next == '\\') out += '\\';
                else if (next == 'n') out += '\n';
                i++;
            } else {
                out += s[i];
            }
        }
        return out;
    }

    std::string extractValue(const std::string& line) {
        size_t start = line.find(":");
        if (start == std::string::npos) return "";
        start = line.find("\"", start);
        if (start == std::string::npos) return "";
        start++;
        size_t end = line.find("\"", start);
        while (end != std::string::npos && line[end-1] == '\\') {
             end = line.find("\"", end + 1);
        }
        if (end == std::string::npos) return "";
        return line.substr(start, end - start);
    }

    std::string extractQuotedString(const std::string& line) {
        size_t start = line.find('"');
        if (start == std::string::npos) return "";
        start++;
        size_t end = line.find('"', start);
        while (end != std::string::npos && line[end - 1] == '\\') {
            end = line.find('"', end + 1);
        }
        if (end == std::string::npos) return "";
        return line.substr(start, end - start);
    }

    std::string loadBinaryDataForIndex(size_t targetIndex) {
        std::ifstream in(filePath);
        std::string line;
        size_t currentIndex = 0;
        bool inside = false;
        
        while (smartGetLine(in, line, currentIndex != targetIndex)) {
            line.erase(0, line.find_first_not_of(" \t"));
            
            if (line.find("{") == 0) {
                inside = true;
            } else if (line.find("}") == 0) {
                if (inside) {
                    inside = false;
                    currentIndex++;
                    if (currentIndex > targetIndex) return ""; 
                }
            } else if (inside && currentIndex == targetIndex) {
                if (line.find("\"binaryData\":") == 0) {
                    return unescape(extractValue(line));
                }
            }
        }
        return "";
    }

    std::map<size_t, std::string> loadBinaryDataMap() {
        std::map<size_t, std::string> dataMap;
        std::ifstream in(filePath);
        std::string line;
        size_t currentIndex = 0;
        bool inside = false;
        std::string currentBinaryData;

        while (std::getline(in, line)) {
            line.erase(0, line.find_first_not_of(" \t"));
            if (line.find("{") == 0) {
                inside = true;
                currentBinaryData = "";
            } else if (line.find("}") == 0) {
                if (inside) {
                    if (!currentBinaryData.empty()) {
                        dataMap[currentIndex] = currentBinaryData;
                    }
                    inside = false;
                    currentIndex++;
                }
            } else if (inside) {
                if (line.find("\"binaryData\":") == 0) {
                    currentBinaryData = unescape(extractValue(line));
                }
            }
        }
        return dataMap;
    }

    bool smartGetLine(std::ifstream& in, std::string& line, bool skipBinary) {
        if (!skipBinary) {
            return (bool)std::getline(in, line);
        }

        line.clear();
        if (line.capacity() < 512) line.reserve(512);

        char c;
        bool checkKey = true;

        while (in.get(c)) {
            if (c == '\n') return true;
            line.push_back(c);

            if (checkKey && line.length() <= 64) {
                if (line.find("\"binaryData\":") != std::string::npos) {
                    in.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
                    return true;
                }
            } else if (line.length() > 64) {
                checkKey = false;
            }
        }
        return !line.empty();
    }
};
