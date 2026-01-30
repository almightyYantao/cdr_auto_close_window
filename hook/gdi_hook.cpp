// gdi_hook.cpp
// 编译: cl /LD /EHsc gdi_hook.cpp user32.lib gdi32.lib /Fe:gdi_hook.dll
// 或 MinGW: g++ -shared -o gdi_hook.dll gdi_hook.cpp -lgdi32 -luser32

#include <windows.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <mutex>

// 共享内存结构
#define SHARED_MEM_NAME L"CDRPopupHandlerSharedMem"
#define MAX_TEXT_LENGTH 4096
#define MAX_TEXT_COUNT 100

struct SharedData {
    DWORD textCount;
    HWND targetHwnd;
    wchar_t texts[MAX_TEXT_COUNT][MAX_TEXT_LENGTH];
    DWORD ready;
};

// 全局变量
static HMODULE g_hModule = NULL;
static HANDLE g_hMapFile = NULL;
static SharedData* g_pSharedData = NULL;
static std::mutex g_mutex;

// 原始函数指针
typedef BOOL (WINAPI *TextOutW_t)(HDC, int, int, LPCWSTR, int);
typedef BOOL (WINAPI *TextOutA_t)(HDC, int, int, LPCSTR, int);
typedef int (WINAPI *DrawTextW_t)(HDC, LPCWSTR, int, LPRECT, UINT);
typedef int (WINAPI *DrawTextA_t)(HDC, LPCSTR, int, LPRECT, UINT);
typedef int (WINAPI *DrawTextExW_t)(HDC, LPWSTR, int, LPRECT, UINT, LPDRAWTEXTPARAMS);
typedef int (WINAPI *DrawTextExA_t)(HDC, LPSTR, int, LPRECT, UINT, LPDRAWTEXTPARAMS);

static TextOutW_t Real_TextOutW = NULL;
static TextOutA_t Real_TextOutA = NULL;
static DrawTextW_t Real_DrawTextW = NULL;
static DrawTextA_t Real_DrawTextA = NULL;
static DrawTextExW_t Real_DrawTextExW = NULL;
static DrawTextExA_t Real_DrawTextExA = NULL;

// IAT Hook 辅助函数
void* HookIAT(HMODULE hModule, const char* dllName, const char* funcName, void* newFunc) {
    ULONG size;
    PIMAGE_IMPORT_DESCRIPTOR pImportDesc = (PIMAGE_IMPORT_DESCRIPTOR)
        ImageDirectoryEntryToDataEx(hModule, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &size, NULL);
    
    if (!pImportDesc) return NULL;
    
    while (pImportDesc->Name) {
        char* moduleName = (char*)((BYTE*)hModule + pImportDesc->Name);
        if (_stricmp(moduleName, dllName) == 0) {
            PIMAGE_THUNK_DATA pThunk = (PIMAGE_THUNK_DATA)((BYTE*)hModule + pImportDesc->FirstThunk);
            PIMAGE_THUNK_DATA pOrigThunk = (PIMAGE_THUNK_DATA)((BYTE*)hModule + pImportDesc->OriginalFirstThunk);
            
            while (pThunk->u1.Function) {
                if (!(pOrigThunk->u1.Ordinal & IMAGE_ORDINAL_FLAG)) {
                    PIMAGE_IMPORT_BY_NAME pImport = (PIMAGE_IMPORT_BY_NAME)((BYTE*)hModule + pOrigThunk->u1.AddressOfData);
                    if (strcmp((char*)pImport->Name, funcName) == 0) {
                        DWORD oldProtect;
                        VirtualProtect(&pThunk->u1.Function, sizeof(void*), PAGE_READWRITE, &oldProtect);
                        void* oldFunc = (void*)pThunk->u1.Function;
                        pThunk->u1.Function = (ULONG_PTR)newFunc;
                        VirtualProtect(&pThunk->u1.Function, sizeof(void*), oldProtect, &oldProtect);
                        return oldFunc;
                    }
                }
                pThunk++;
                pOrigThunk++;
            }
        }
        pImportDesc++;
    }
    return NULL;
}

// 保存捕获的文本
void SaveCapturedText(const wchar_t* text, int len) {
    if (!g_pSharedData || !text || len <= 0) return;
    
    std::lock_guard<std::mutex> lock(g_mutex);
    
    if (g_pSharedData->textCount < MAX_TEXT_COUNT) {
        int copyLen = min(len, MAX_TEXT_LENGTH - 1);
        wcsncpy_s(g_pSharedData->texts[g_pSharedData->textCount], MAX_TEXT_LENGTH, text, copyLen);
        g_pSharedData->texts[g_pSharedData->textCount][copyLen] = L'\0';
        g_pSharedData->textCount++;
    }
}

// Hook 函数实现
BOOL WINAPI Hook_TextOutW(HDC hdc, int x, int y, LPCWSTR lpString, int c) {
    if (lpString && c > 0) {
        SaveCapturedText(lpString, c);
    }
    return Real_TextOutW(hdc, x, y, lpString, c);
}

BOOL WINAPI Hook_TextOutA(HDC hdc, int x, int y, LPCSTR lpString, int c) {
    if (lpString && c > 0) {
        // 转换为 Unicode
        int wlen = MultiByteToWideChar(CP_ACP, 0, lpString, c, NULL, 0);
        if (wlen > 0) {
            wchar_t* wstr = new wchar_t[wlen + 1];
            MultiByteToWideChar(CP_ACP, 0, lpString, c, wstr, wlen);
            wstr[wlen] = L'\0';
            SaveCapturedText(wstr, wlen);
            delete[] wstr;
        }
    }
    return Real_TextOutA(hdc, x, y, lpString, c);
}

int WINAPI Hook_DrawTextW(HDC hdc, LPCWSTR lpchText, int cchText, LPRECT lprc, UINT format) {
    if (lpchText) {
        int len = (cchText == -1) ? wcslen(lpchText) : cchText;
        if (len > 0) {
            SaveCapturedText(lpchText, len);
        }
    }
    return Real_DrawTextW(hdc, lpchText, cchText, lprc, format);
}

int WINAPI Hook_DrawTextA(HDC hdc, LPCSTR lpchText, int cchText, LPRECT lprc, UINT format) {
    if (lpchText) {
        int len = (cchText == -1) ? strlen(lpchText) : cchText;
        if (len > 0) {
            int wlen = MultiByteToWideChar(CP_ACP, 0, lpchText, len, NULL, 0);
            if (wlen > 0) {
                wchar_t* wstr = new wchar_t[wlen + 1];
                MultiByteToWideChar(CP_ACP, 0, lpchText, len, wstr, wlen);
                wstr[wlen] = L'\0';
                SaveCapturedText(wstr, wlen);
                delete[] wstr;
            }
        }
    }
    return Real_DrawTextA(hdc, lpchText, cchText, lprc, format);
}

int WINAPI Hook_DrawTextExW(HDC hdc, LPWSTR lpchText, int cchText, LPRECT lprc, UINT format, LPDRAWTEXTPARAMS lpdtp) {
    if (lpchText) {
        int len = (cchText == -1) ? wcslen(lpchText) : cchText;
        if (len > 0) {
            SaveCapturedText(lpchText, len);
        }
    }
    return Real_DrawTextExW(hdc, lpchText, cchText, lprc, format, lpdtp);
}

int WINAPI Hook_DrawTextExA(HDC hdc, LPSTR lpchText, int cchText, LPRECT lprc, UINT format, LPDRAWTEXTPARAMS lpdtp) {
    if (lpchText) {
        int len = (cchText == -1) ? strlen(lpchText) : cchText;
        if (len > 0) {
            int wlen = MultiByteToWideChar(CP_ACP, 0, lpchText, len, NULL, 0);
            if (wlen > 0) {
                wchar_t* wstr = new wchar_t[wlen + 1];
                MultiByteToWideChar(CP_ACP, 0, lpchText, len, wstr, wlen);
                wstr[wlen] = L'\0';
                SaveCapturedText(wstr, wlen);
                delete[] wstr;
            }
        }
    }
    return Real_DrawTextExA(hdc, lpchText, cchText, lprc, format, lpdtp);
}

// 初始化共享内存
BOOL InitSharedMemory() {
    g_hMapFile = CreateFileMappingW(
        INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE, 0, sizeof(SharedData), SHARED_MEM_NAME);
    
    if (!g_hMapFile) {
        g_hMapFile = OpenFileMappingW(FILE_MAP_ALL_ACCESS, FALSE, SHARED_MEM_NAME);
    }
    
    if (!g_hMapFile) return FALSE;
    
    g_pSharedData = (SharedData*)MapViewOfFile(g_hMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(SharedData));
    if (!g_pSharedData) {
        CloseHandle(g_hMapFile);
        g_hMapFile = NULL;
        return FALSE;
    }
    
    return TRUE;
}

// 安装 Hook
void InstallHooks() {
    HMODULE hGdi32 = GetModuleHandleW(L"gdi32.dll");
    if (!hGdi32) return;
    
    // 获取原始函数地址
    Real_TextOutW = (TextOutW_t)GetProcAddress(hGdi32, "TextOutW");
    Real_TextOutA = (TextOutA_t)GetProcAddress(hGdi32, "TextOutA");
    Real_DrawTextW = (DrawTextW_t)GetProcAddress(GetModuleHandleW(L"user32.dll"), "DrawTextW");
    Real_DrawTextA = (DrawTextA_t)GetProcAddress(GetModuleHandleW(L"user32.dll"), "DrawTextA");
    Real_DrawTextExW = (DrawTextExW_t)GetProcAddress(GetModuleHandleW(L"user32.dll"), "DrawTextExW");
    Real_DrawTextExA = (DrawTextExA_t)GetProcAddress(GetModuleHandleW(L"user32.dll"), "DrawTextExA");
    
    // Hook 当前进程的所有模块
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId());
    if (hSnapshot != INVALID_HANDLE_VALUE) {
        MODULEENTRY32W me;
        me.dwSize = sizeof(me);
        if (Module32FirstW(hSnapshot, &me)) {
            do {
                HMODULE hMod = me.hModule;
                HookIAT(hMod, "gdi32.dll", "TextOutW", Hook_TextOutW);
                HookIAT(hMod, "gdi32.dll", "TextOutA", Hook_TextOutA);
                HookIAT(hMod, "user32.dll", "DrawTextW", Hook_DrawTextW);
                HookIAT(hMod, "user32.dll", "DrawTextA", Hook_DrawTextA);
                HookIAT(hMod, "user32.dll", "DrawTextExW", Hook_DrawTextExW);
                HookIAT(hMod, "user32.dll", "DrawTextExA", Hook_DrawTextExA);
            } while (Module32NextW(hSnapshot, &me));
        }
        CloseHandle(hSnapshot);
    }
}

// DLL 入口点
BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID reserved) {
    switch (reason) {
        case DLL_PROCESS_ATTACH:
            g_hModule = hModule;
            DisableThreadLibraryCalls(hModule);
            if (InitSharedMemory()) {
                InstallHooks();
            }
            break;
            
        case DLL_PROCESS_DETACH:
            if (g_pSharedData) {
                UnmapViewOfFile(g_pSharedData);
                g_pSharedData = NULL;
            }
            if (g_hMapFile) {
                CloseHandle(g_hMapFile);
                g_hMapFile = NULL;
            }
            break;
    }
    return TRUE;
}

// 导出函数用于测试
extern "C" __declspec(dllexport) void ClearTexts() {
    if (g_pSharedData) {
        std::lock_guard<std::mutex> lock(g_mutex);
        g_pSharedData->textCount = 0;
    }
}

extern "C" __declspec(dllexport) DWORD GetTextCount() {
    return g_pSharedData ? g_pSharedData->textCount : 0;
}
