// AudiodgProbe.cpp — check whether audiodg.exe has loaded MyCompanyEfxApo.dll
// Build: cl AudiodgProbe.cpp /nologo /W3 /EHsc /utf-8

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <tlhelp32.h>
#include <psapi.h>
#include <cstdio>

static DWORD FindProcessId(const wchar_t* exeName)
{
    DWORD pid = 0;
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE) return 0;
    PROCESSENTRY32W pe = { sizeof(pe) };
    if (Process32FirstW(snap, &pe)) {
        do {
            if (!_wcsicmp(pe.szExeFile, exeName)) { pid = pe.th32ProcessID; break; }
        } while (Process32NextW(snap, &pe));
    }
    CloseHandle(snap);
    return pid;
}

int wmain()
{
    const wchar_t* targetExe = L"audiodg.exe";
    const wchar_t* dllName   = L"MyCompanyEfxApo.dll";

    DWORD pid = FindProcessId(targetExe);
    if (!pid) {
        wprintf(L"[!] %s not running.\n", targetExe);
        return 1;
    }
    HANDLE h = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, pid);
    if (!h) {
        wprintf(L"[!] OpenProcess failed (%lu)\n", GetLastError());
        return 1;
    }

    HMODULE hMods[1024];
    DWORD cbNeeded = 0;
    if (!EnumProcessModules(h, hMods, sizeof(hMods), &cbNeeded)) {
        wprintf(L"[!] EnumProcessModules failed (%lu)\n", GetLastError());
        CloseHandle(h);
        return 1;
    }
    unsigned count = cbNeeded / sizeof(HMODULE);
    bool found = false;
    for (unsigned i = 0; i < count; ++i) {
        wchar_t path[MAX_PATH] = {0};
        if (GetModuleFileNameExW(h, hMods[i], path, MAX_PATH)) {
            const wchar_t* base = wcsrchr(path, L'\\');
            base = base ? (base + 1) : path;
            if (!_wcsicmp(base, dllName)) {
                wprintf(L"[OK] Found in audiodg: %s\n", path);
                found = true;
                break;
            }
        }
    }
    if (!found) {
        wprintf(L"[X] NOT found in audiodg: %s\n", dllName);
    }
    CloseHandle(h);
    return found ? 0 : 2;
}
