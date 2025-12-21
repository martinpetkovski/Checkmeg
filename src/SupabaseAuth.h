#pragma once

#include <string>
#include <cstdint>

struct SupabaseSession {
    bool loggedIn = false;
    std::string email;
    std::string accessToken;
    std::string refreshToken;
    std::int64_t expiresAtUnix = 0;
};

class SupabaseAuth {
public:
    SupabaseAuth();

    const SupabaseSession& Session() const { return session_; }
    bool IsLoggedIn() const { return session_.loggedIn && !session_.refreshToken.empty(); }
    bool LoadSessionFromDisk();
    bool TryRestoreOrRefresh(std::string* outError);

    bool SignInWithPassword(const std::string& email, const std::string& password, std::string* outError);
    bool SignUpWithPassword(const std::string& email, const std::string& password, std::string* outError);
    void Logout();

private:
    SupabaseSession session_;

    std::wstring GetSessionFilePathW() const;
    bool EnsureSessionDirExists() const;

    bool SaveSessionToDisk(std::string* outError) const;
    void ClearSessionOnDisk() const;

    bool RefreshWithToken(const std::string& refreshToken, std::string* outError);

    static std::string BuildJsonEmailPassword(const std::string& email, const std::string& password);
    static std::string BuildJsonRefreshToken(const std::string& refreshToken);
};
