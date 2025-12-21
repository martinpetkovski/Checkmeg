cmake_minimum_required(VERSION 3.10)
project(Checkmeg)

set(CMAKE_CXX_STANDARD 20)

# Add source files
file(GLOB SOURCES "src/*.cpp" "src/*.h")

# Create executable
add_executable(Checkmeg ${SOURCES})

# Link libraries
target_link_libraries(Checkmeg user32 kernel32 gdi32 ole32 shell32 shlwapi winhttp)
