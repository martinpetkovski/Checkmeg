Stop-Process -Name Checkmeg -Force -ErrorAction SilentlyContinue
clang++ -std=c++20 -O2 -municode -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -DNOMINMAX -I. .\src\main.cpp -o .\Checkmeg.exe -luser32 -lgdi32 -lshell32 -lole32 -lcomdlg32 -lcomctl32 -ladvapi32 -lgdiplus
if ($LASTEXITCODE -eq 0) {
    Start-Process .\Checkmeg.exe
}
