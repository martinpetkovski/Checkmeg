Stop-Process -Name Checkmeg -Force -ErrorAction SilentlyContinue

# Compile resources
llvm-rc src/resource.rc /FO resource.res
if ($LASTEXITCODE -ne 0) {
    Write-Error "Resource compilation failed."
    exit 1
}

clang++ -std=c++20 -O2 -municode -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -DNOMINMAX -I. .\src\main.cpp .\src\SupabaseAuth.cpp resource.res -o .\Checkmeg.exe -luser32 -lgdi32 -lshell32 -lole32 -lcomdlg32 -lcomctl32 -ladvapi32 -lgdiplus -lwinhttp
if ($LASTEXITCODE -eq 0) {
    Start-Process .\Checkmeg.exe
}
