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

enum class BookmarkType {
    Text,
    URL,
    File,
    Command,
    Binary
};

struct Bookmark {
    BookmarkType type;
    bool typeExplicit = false;
    std::string content;
    std::string binaryData; // For Binary type
    std::string mimeType; // For Binary type
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

    bool hasLastWriteTime = false;
    std::filesystem::file_time_type lastWriteTime;

    BookmarkManager(const std::string& path) : filePath(path) {
        load(true);
    }

    void loadIfChanged(bool skipBinary = false) {
        std::error_code ec;
        auto wt = std::filesystem::last_write_time(filePath, ec);
        if (ec) {
            // If we can't stat the file (missing, permissions), fall back to load.
            load(skipBinary);
            return;
        }

        if (!hasLastWriteTime || wt != lastWriteTime) {
            load(skipBinary);
        }
    }

    void add(const std::string& content, const std::string& deviceId = "", bool validOnAnyDevice = true) {
        Bookmark b;
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
    }

    void addBinary(const std::string& content, const std::string& data, const std::string& mime, const std::string& deviceId = "", bool validOnAnyDevice = true) {
        Bookmark b;
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
    }

    void remove(size_t index) {
        if (index < bookmarks.size()) {
            bookmarks.erase(bookmarks.begin() + index);
            save();
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
        }
    }

    void updateBinaryData(size_t index, const std::string& data) {
        if (index < bookmarks.size()) {
            bookmarks[index].binaryData = data;
            bookmarks[index].binaryDataLoaded = true;
            save();
        }
    }
void updateLastUsed(size_t index) {
        if (index < bookmarks.size()) {
            bookmarks[index].lastUsed = std::time(nullptr);
            save();
        }
    }
    void ensureBinaryDataLoaded(size_t index) {
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
        // Check if we need to load missing binary data
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
            out << "    \"type\": \"" << typeToString(b.type) << "\",\n";
            out << "    \"typeExplicit\": " << (b.typeExplicit ? "true" : "false") << ",\n";
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
            out << "    \"content\": \"" << escape(b.content) << "\",\n";
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
        // For the "start simple" phase, we will implement a very basic parser 
        // that assumes the exact format we write.
        // In a real app, use nlohmann/json.
        bookmarks.clear();
        std::ifstream in(filePath);
        std::string line;
        Bookmark current;
        bool inside = false;
        bool insideTags = false;
        size_t currentIndex = 0;

        while (smartGetLine(in, line, skipBinary)) {
            // Trim whitespace
            line.erase(0, line.find_first_not_of(" \t"));
            current.lastUsed = 0;
                
            if (line.find("{") == 0) {
                inside = true;
                current = Bookmark();
                current.tags.clear();
                // Default values for new fields if missing in JSON
                current.validOnAnyDevice = true; 
                current.deviceId = "";
                current.originalFileIndex = currentIndex;
                current.binaryDataLoaded = true;
                insideTags = false;
            } else if (line.find("}") == 0) {
                if (inside) {
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
                    // Tag line like: "value",
                    std::string val = extractQuotedString(line);
                    if (!val.empty()) current.tags.push_back(unescape(val));
                    continue;
                }

                if (line.find("\"type\":") == 0) {
                    std::string val = extractValue(line);
                    if (val == "url") current.type = BookmarkType::URL;
                    else if (val == "file") current.type = BookmarkType::File;
                    else if (val == "command") current.type = BookmarkType::Command;
                    else if (val == "binary") current.type = BookmarkType::Binary;
                    else current.type = BookmarkType::Text;
                } else if (line.find("\"typeExplicit\":") == 0) {
                    // very simple bool parse
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
                } else if (line.find("\"tags\":") == 0) {
                    insideTags = true;
                } else if (line.find("\"content\":") == 0) {
                    current.content = unescape(extractValue(line));
                } else if (line.find("\"timestamp\":") == 0) {
                    std::string val = line.substr(line.find(":") + 1);
                    // remove comma if present
                    if (val.find(",") != std::string::npos) val = val.substr(0, val.find(","));
                    // trim
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

        // Backwards compatibility: if old files didn't have typeExplicit,
        // default is false (auto type) which is already the struct default.

        std::error_code ec;
        lastWriteTime = std::filesystem::last_write_time(filePath, ec);
        hasLastWriteTime = !ec;
    }

private:
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
        // Handle escaped quotes in value? Our simple parser might fail here if not careful.
        // For now, assume simple structure.
        while (end != std::string::npos && line[end-1] == '\\') {
             end = line.find("\"", end + 1);
        }
        if (end == std::string::npos) return "";
        return line.substr(start, end - start);
    }

    // Extracts the first JSON string literal found in the line, without requiring a key/value colon.
    // Intended for parsing array entries like: "tag", or "tag"
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
                // Detect the key early and skip the (potentially huge) base64 payload.
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
