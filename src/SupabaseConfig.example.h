#pragma once
#include <string>

namespace SupabaseConfig {
inline const std::string SUPABASE_URL = "https://YOUR_PROJECT.supabase.co";
inline const std::string SUPABASE_ANON_KEY = "YOUR_SUPABASE_ANON_KEY";

inline std::string AuthBaseUrl() {
    return SUPABASE_URL + "/auth/v1";
}

inline std::string RestBaseUrl() {
    return SUPABASE_URL + "/rest/v1";
}
inline const std::string BOOKMARKS_TABLE = "bookmarks";

inline std::string EndpointBookmarksTable() {
    return RestBaseUrl() + "/" + BOOKMARKS_TABLE;
}

inline std::string EndpointSignUp() {
    return AuthBaseUrl() + "/signup";
}

inline std::string EndpointTokenPassword() {
    return AuthBaseUrl() + "/token?grant_type=password";
}

inline std::string EndpointTokenRefresh() {
    return AuthBaseUrl() + "/token?grant_type=refresh_token";
}

inline std::string EndpointLogout() {
    return AuthBaseUrl() + "/logout";
}

}
