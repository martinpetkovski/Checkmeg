# Checkmeg

Checkmeg is a simple C++ Win32 application that allows you to bookmark links and text from any window and search for them later.

## Features
- **Capture Selection**: Press `WIN + ALT + C` to copy the currently selected text in any active window and save it as a bookmark.
- **Search Bookmarks**: Press `WIN + ALT + SPACE` to open the search bar.
- **Launch**: Double-click a bookmark in the search list to open it (URLs open in browser, files open in default app).
- **Storage**: Bookmarks are saved in `bookmarks.json` in the same directory as the executable.

## How to Build

### Prerequisites
- A C++ compiler (Clang, MSVC, or MinGW).
- CMake (optional).

### Building with Clang (Recommended)
Run the following command in the terminal:
```powershell
clang++ src/main.cpp -o Checkmeg.exe -luser32 -lkernel32 -lgdi32 -lole32 -lshell32 -lshlwapi
```

### Building with CMake
```powershell
mkdir build
cd build
cmake ..
cmake --build . --config Release
```

## Usage
1. Run `Checkmeg.exe`. It will run in the background (no window will appear initially).
2. Select some text in a web browser or text editor.
3. Press `WIN + ALT + C`. You should hear a beep indicating success.
4. Press `WIN + ALT + SPACE` to open the search UI.
5. Type to filter bookmarks.
6. Double-click an item to open it.
7. Press `ESC` to close the search UI.
