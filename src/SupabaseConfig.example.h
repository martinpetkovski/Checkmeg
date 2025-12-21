#pragma once

// Supabase configuration (EXAMPLE)
//
// Copy this file to `src/SupabaseConfig.h` and fill in your project values.
// IMPORTANT: `src/SupabaseConfig.h` is git-ignored so you don't accidentally
// commit keys/URLs.
//
// - SUPABASE_URL:       e.g. "https://YOUR_PROJECT.supabase.co"
// - SUPABASE_ANON_KEY:  your project's anon public API key (sb_publishable_...)
//
// NOTE: The anon key is intended for client apps, but you still generally
// shouldn't publish your specific project details. RLS must be enabled.

#include <string>

namespace SupabaseConfig {

// NOTE: Keep the trailing slash OFF.
inline const std::string SUPABASE_URL = "https://YOUR_PROJECT.supabase.co";
inline const std::string SUPABASE_ANON_KEY = "YOUR_SUPABASE_ANON_KEY";

inline std::string AuthBaseUrl() {
    return SUPABASE_URL + "/auth/v1";
}

inline std::string RestBaseUrl() {
    return SUPABASE_URL + "/rest/v1";
}

// Change this if your table name differs.
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
