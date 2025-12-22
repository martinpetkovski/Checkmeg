# Checkmeg

Checkmeg is a tiny Windows app that lives in your system tray and helps you save little bits of information you want to reuse later — like links, file paths, short notes, or commands.

Think of it as a “personal clipboard with search”.

> **Download:** [Latest release](https://github.com/martinpetkovski/Checkmeg/releases/latest)
> 
> **Hotkeys (defaults):** Capture `Win + Left Alt + X` • Search `Right Alt` (configurable in Options)

**Quick links:**
- [How to use it](#the-simple-version-how-youll-use-it)
- [What gets saved](#what-gets-saved)
- [Where your stuff is stored](#where-your-stuff-is-stored)
- [Supabase sync (optional)](#advanced-optional-supabase-sync)
- [Build from source](#advanced-developers-build)

## The simple version (how you’ll use it)

1) Run `Checkmeg.exe`
	- It won’t open a big window.
	- It sits near the clock in your system tray.

2) Save something you might want again
	- Select something (a URL, text, a file path — anything you can copy).
	- Press `Win + Left Alt + X` to save it into Checkmeg.

3) Find it later
	- Press `Right Alt` to open the search window.
	- Start typing to filter your saved items.

4) Use it
	- Pick an item and run/open/insert it (what happens depends on what you saved).

## What gets saved

Checkmeg tries to figure out what you copied and treats it appropriately:

- **A website link** → opens in your browser
- **A file path** → opens the file (or the folder) normally
- **Plain text** → gets typed into whatever app you’re currently using
- **A command** → can be run via PowerShell
- **A file’s raw contents** (advanced) → can be saved back to disk when you use it

## Where your stuff is stored

- By default, everything is saved locally in `bookmarks.json` next to `Checkmeg.exe`.
- Nothing is uploaded anywhere unless you enable the optional Supabase sync (advanced) by logging in.

## Safety note (commands)

If you save something as a **Command**, running it will execute it via PowerShell.
Only run commands you trust.

## Helpful extras

- You can edit, duplicate, or delete saved items from the search list.
- You can add tags (comma-separated) to keep things organized.
- There’s an Options window (from the tray menu) with things like export and “Run at Startup”.

---

## Advanced (optional): Supabase sync

If you want to sync bookmarks across devices, Checkmeg can use Supabase.

- Sign up / sign in with email + password
- Session is stored at `%APPDATA%\Checkmeg\session.json`
- Sync features include refresh and a one-click “sync local data to Supabase”

**Setup (developers / self-hosters):**

1) Create `src/SupabaseConfig.h`
- Copy [src/SupabaseConfig.example.h](src/SupabaseConfig.example.h) → `src/SupabaseConfig.h`
- Fill in `SUPABASE_URL` and `SUPABASE_ANON_KEY`

2) Create the Supabase table + policies
- Run the SQL in [supabase/bookmarks_schema.sql](supabase/bookmarks_schema.sql) from the Supabase Dashboard SQL editor.

## Advanced (developers): Build

### Prerequisites
- Windows
- LLVM/Clang toolchain (including `clang++` and `llvm-rc`)

### Build
- `./build.ps1`

### Package / release
- `./package.ps1` creates a versioned zip under `package/`
- `./release.ps1` runs packaging then publishes a GitHub release using the `gh` CLI

## Docs site

A small static docs page lives in [docs/](docs/).

- Open `docs/index.html` in a browser, or publish the `docs/` folder to GitHub Pages.
