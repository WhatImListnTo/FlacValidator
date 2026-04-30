// FlacValidatorGUI.cpp
// FLAC Integrity Validator v1.0 — GUI Edition
// Compile as x64 | Release | Subsystem: Windows

#define _CRT_SECURE_NO_WARNINGS
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <thread>
#include <mutex>
#include <atomic>
#include <filesystem>
#include <chrono>
#include <fstream>
#include <algorithm>
#include <shellapi.h>
#include <shobjidl.h>
#include "Resource.h"
#include "FLAC/stream_decoder.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "FLAC.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' " \
    "version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

namespace fs = std::filesystem;

// ============================================================================
//  Layout Constants
// ============================================================================
static constexpr int WIN_W = 550;
static constexpr int WIN_H = 430;
static constexpr int MARGIN = 15;
static constexpr int ROW_H = 25;
static constexpr int GAP = 8;

// Custom messages
static constexpr UINT WM_VAL_COMPLETE = WM_USER + 1;
static constexpr UINT WM_VAL_FAILURE = WM_USER + 2;
static constexpr UINT WM_SCAN_DONE = WM_USER + 3;

// Timer
static constexpr UINT_PTR TIMER_ID = 1;
static constexpr UINT     TIMER_MS = 200;

// Control IDs
enum CtlID {
    ID_PATH_EDIT = 101, ID_BROWSE_BTN, ID_THREADS_COMBO,
    ID_START_BTN, ID_CANCEL_BTN, ID_EXPORT_BTN,
    ID_PROGRESS_BAR, ID_FAILURES_LIST
};

// ============================================================================
//  Globals — UI
// ============================================================================
static HINSTANCE g_hInst;
static HWND g_hWnd;
static HWND g_hPathEdit, g_hBrowseBtn, g_hThreadsCombo;
static HWND g_hStartBtn, g_hCancelBtn, g_hExportBtn;
static HWND g_hProgressBar, g_hPctLabel;
static HWND g_hFilesLbl, g_hFailedLbl;
static HWND g_hSpeedLbl, g_hDataLbl;
static HWND g_hElapsedLbl, g_hEtaLbl;
static HWND g_hFailGrp, g_hFailList;

// ============================================================================
//  Globals — Engine
// ============================================================================
static std::atomic<bool>     g_cancel{ false };
static std::atomic<bool>     g_running{ false };
static std::vector<std::wstring> g_files;
static std::vector<std::wstring> g_failures;
static std::mutex            g_failMtx;
static std::atomic<uint32_t> g_done{ 0 };
static std::atomic<uint32_t> g_failed{ 0 };
static std::atomic<uint64_t> g_bytes{ 0 };
static uint32_t              g_total = 0;
static int                   g_nThreads = 8;
static std::wstring          g_rootPath;
static std::chrono::steady_clock::time_point g_t0;

// ============================================================================
//  Globals — Taskbar Progress
// ============================================================================
static ITaskbarList3* g_pTaskbar = nullptr;
static UINT           g_uTaskbarBtnCreated = 0;

// ============================================================================
//  FLAC Decoder Context
// ============================================================================
//  Single context struct passed as client_data to ALL decoder callbacks.
//  Using init_stream (not init_FILE) gives us exclusive ownership of the
//  FILE* — no ambiguity about whether finish() closes it.
// ============================================================================
struct FlacCtx {
    FILE* file = nullptr;
    bool        error = false;
    std::string detail;
};

// ============================================================================
//  Centered MessageBox — CBT Hook
// ============================================================================
//  Standard MessageBox centers on the screen/monitor, not on the
//  owner window.  This helper installs a short-lived CBT hook that
//  intercepts the MessageBox activation and repositions it to the
//  center of the owner window, clamped to the monitor work area.
// ============================================================================
static HHOOK g_hMsgBoxHook = nullptr;
static HWND  g_hMsgBoxOwner = nullptr;

static LRESULT CALLBACK MsgBoxCBTProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HCBT_ACTIVATE) {
        HWND hMsgBox = (HWND)wParam;
        
        if (g_hMsgBoxOwner) {
            RECT rcOwner, rcMsg;
            GetWindowRect(g_hMsgBoxOwner, &rcOwner);
            GetWindowRect(hMsgBox, &rcMsg);

            int msgW = rcMsg.right - rcMsg.left;
            int msgH = rcMsg.bottom - rcMsg.top;
            int ownerCX = (rcOwner.left + rcOwner.right) / 2;
            int ownerCY = (rcOwner.top + rcOwner.bottom) / 2;

            int x = ownerCX - msgW / 2;
            int y = ownerCY - msgH / 2;

            // Clamp to monitor work area so dialog stays on-screen
            HMONITOR hMon = MonitorFromWindow(g_hMsgBoxOwner, MONITOR_DEFAULTTONEAREST);
            MONITORINFO mi = { sizeof(mi) };

            if (GetMonitorInfo(hMon, &mi)) {
                const RECT& wa = mi.rcWork;
                if (x + msgW > wa.right)  x = wa.right - msgW;
                if (y + msgH > wa.bottom) y = wa.bottom - msgH;
                if (x < wa.left) x = wa.left;
                if (y < wa.top)  y = wa.top;
            }

            SetWindowPos(hMsgBox, nullptr, x, y, 0, 0,
                SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        UnhookWindowsHookEx(g_hMsgBoxHook);
        g_hMsgBoxHook = nullptr;
    }
    return CallNextHookEx(g_hMsgBoxHook, nCode, wParam, lParam);
}

static int CenteredMessageBox(HWND hOwner, LPCWSTR text, LPCWSTR caption, UINT type) {
    g_hMsgBoxOwner = hOwner;
    g_hMsgBoxHook = SetWindowsHookEx(WH_CBT, MsgBoxCBTProc, nullptr, GetCurrentThreadId());
    return MessageBox(hOwner, text, caption, type);
}

// ============================================================================
//  INI File Persistence
// ============================================================================
static WCHAR g_iniPath[MAX_PATH] = {};

void InitIniPath(HINSTANCE hInst) {
    GetModuleFileNameW(hInst, g_iniPath, MAX_PATH);
    WCHAR* slash = wcsrchr(g_iniPath, L'\\');

    if (slash) {
        *(slash + 1) = L'\0';
        wcscat_s(g_iniPath, MAX_PATH, L"FlacValidator.ini");
    }
}

void SaveLastPath(const WCHAR* path) {
    WritePrivateProfileStringW(L"Settings", L"LastPath", path, g_iniPath);
}

void LoadLastPath(WCHAR* path, int maxLen) {
    GetPrivateProfileStringW(L"Settings", L"LastPath", L"", path, maxLen, g_iniPath);
}

void SaveThreads(int threads) {
    WCHAR buf[16];
    swprintf_s(buf, L"%d", threads);
    WritePrivateProfileStringW(L"Settings", L"Threads", buf, g_iniPath);
}

int LoadThreads(int defaultVal) {
    return GetPrivateProfileIntW(L"Settings", L"Threads", defaultVal, g_iniPath);
}

// ============================================================================
//  Helpers
// ============================================================================
static std::string toUtf8(const std::wstring& w) {
    if (w.empty())
        return {};

    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, s.data(), n, nullptr, nullptr);
    return s;
}

static std::wstring fmtTime(int secs) {
    wchar_t b[32];
    swprintf_s(b, L"%02d:%02d:%02d", secs / 3600, (secs % 3600) / 60, secs % 60);
    return b;
}

static std::wstring fmtSize(uint64_t bytes) {
    wchar_t b[64];

    if (bytes >= (1ULL << 30))
        swprintf_s(b, L"%.2f GB", (double)bytes / (1ULL << 30));
    else if (bytes >= (1ULL << 20))
        swprintf_s(b, L"%.1f MB", (double)bytes / (1ULL << 20));
    else
        swprintf_s(b, L"%llu bytes", (unsigned long long)bytes);

    return b;
}

static std::wstring fmtSpeed(double bps) {
    wchar_t b[64];

    if (bps >= 1024.0 * 1024.0)
        swprintf_s(b, L"%.1f MB/s", bps / (1024.0 * 1024.0));
    else if (bps >= 1024.0)
        swprintf_s(b, L"%.1f KB/s", bps / 1024.0);
    else
        swprintf_s(b, L"%.0f B/s", bps);

    return b;
}

// ============================================================================
//  Clipboard Copy for Listbox
// ============================================================================
static void CopyListboxItemToClipboard(HWND hList, int index) {
    int len = (int)SendMessage(hList, LB_GETTEXTLEN, index, 0);
    
    if (len <= 0) 
        return;

    std::wstring text(len, L'\0');
    SendMessage(hList, LB_GETTEXT, index, (LPARAM)text.data());

    if (OpenClipboard(g_hWnd)) {
        EmptyClipboard();
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (text.size() + 1) * sizeof(WCHAR));

        if (hMem) {
            WCHAR* p = (WCHAR*)GlobalLock(hMem);
            wcscpy(p, text.c_str());
            GlobalUnlock(hMem);
            SetClipboardData(CF_UNICODETEXT, hMem);
        }

        CloseClipboard();
    }
}

static void CopyAllListboxItemsToClipboard(HWND hList) {
    int count = (int)SendMessage(hList, LB_GETCOUNT, 0, 0);
    
    if (count <= 0) 
        return;

    std::wstring all;

    for (int i = 0; i < count; i++) {
        int len = (int)SendMessage(hList, LB_GETTEXTLEN, i, 0);

        if (len <= 0) 
            continue;

        std::wstring line(len, L'\0');
        SendMessage(hList, LB_GETTEXT, i, (LPARAM)line.data());

        if (!all.empty()) 
            all += L"\r\n";

        all += line;
    }

    if (all.empty())
        return;

    if (OpenClipboard(g_hWnd)) {
        EmptyClipboard();
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (all.size() + 1) * sizeof(WCHAR));

        if (hMem) {
            WCHAR* p = (WCHAR*)GlobalLock(hMem);
            wcscpy(p, all.c_str());
            GlobalUnlock(hMem);
            SetClipboardData(CF_UNICODETEXT, hMem);
        }

        CloseClipboard();
    }
}

// ============================================================================
//  Failures Listbox Subclass — Ctrl+C and Right-Click Copy
// ============================================================================
static LRESULT CALLBACK FailListSubclassProc(
    HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR /*dwRefData*/) {

    switch (uMsg) {

    case WM_KEYDOWN:
        if (wParam == 'C' && GetKeyState(VK_CONTROL) < 0) {
            CopyAllListboxItemsToClipboard(hWnd);
            return 0;
        }

        break;

    case WM_CONTEXTMENU: {
        int count = (int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);

        if (count <= 0)
            return 0;

        POINT pt = { LOWORD(lParam), HIWORD(lParam) };
        POINT clientPt = pt;
        ScreenToClient(hWnd, &clientPt);

        // Select item under cursor
        LRESULT item = SendMessage(hWnd, LB_ITEMFROMPOINT, 0,
            MAKELPARAM(clientPt.x, clientPt.y));

        if (HIWORD(item) == 0)
            SendMessage(hWnd, LB_SETCURSEL, LOWORD(item), 0);

        int sel = (int)SendMessage(hWnd, LB_GETCURSEL, 0, 0);

        HMENU hMenu = CreatePopupMenu();

        if (sel != LB_ERR)
            AppendMenuW(hMenu, MF_STRING, 1, L"Copy Selected");

        if (count > 1)
            AppendMenuW(hMenu, MF_STRING, 2, L"Copy All");
        else if (sel == LB_ERR)
            AppendMenuW(hMenu, MF_STRING, 2, L"Copy");

        int cmd = TrackPopupMenu(hMenu,
            TPM_RETURNCMD | TPM_RIGHTBUTTON,
            pt.x, pt.y, 0, hWnd, nullptr);

        if (cmd == 1)
            CopyListboxItemToClipboard(hWnd, sel);
        else if (cmd == 2)
            CopyAllListboxItemsToClipboard(hWnd);

        DestroyMenu(hMenu);
        return 0;
    }

    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, FailListSubclassProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// ============================================================================
//  Tab Stop Helper — Enable/Disable with WS_TABSTOP Toggle
// ============================================================================
//  Disabled controls must not receive tab focus.  This pairs
//  EnableWindow() with WS_TABSTOP management so the tab key
//  skips any control that is currently disabled.
// ============================================================================
static void enableWithTabStop(HWND hwnd, bool enable) {
    EnableWindow(hwnd, enable);
    LONG_PTR style = GetWindowLongPtr(hwnd, GWL_STYLE);

    if (enable)
        style |= WS_TABSTOP;
    else
        style &= ~WS_TABSTOP;

    SetWindowLongPtr(hwnd, GWL_STYLE, style);
}

// ============================================================================
//  Browse Dialog Helpers
// ============================================================================
static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM, LPARAM lpData) {
    if (uMsg == BFFM_INITIALIZED && lpData)
        SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);

    return 0;
}

static std::wstring FindNearestExistingDir(const std::wstring& path) {
    fs::path p(path);

    while (!p.empty() && !fs::exists(p)) {
        fs::path parent = p.parent_path();
        if (parent == p) break;
        p = parent;
    }

    if (!p.empty() && fs::exists(p) && fs::is_directory(p))
        return p.wstring();

    return L"";
}

// ============================================================================
//  FLAC Stream Callbacks
// ============================================================================
//  Using init_stream (not init_FILE) gives us exclusive ownership
//  of the FILE*.  The FLAC library never stores or closes our handle.
//  This eliminates a class of FILE* lifecycle bugs that manifest as
//  double-fclose under high-concurrency / high-throughput conditions.
// ============================================================================
static FLAC__StreamDecoderReadStatus flacReadCb(
    const FLAC__StreamDecoder*, FLAC__byte buffer[],
    size_t* bytes, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);

    // Honor user cancellation immediately — don't wait for the
    // current file to finish decoding.
    if (g_cancel.load(std::memory_order_relaxed))
        return FLAC__STREAM_DECODER_READ_STATUS_ABORT;

    if (*bytes == 0)
        return FLAC__STREAM_DECODER_READ_STATUS_ABORT;

    *bytes = fread(buffer, 1, *bytes, ctx->file);

    if (ferror(ctx->file))
        return FLAC__STREAM_DECODER_READ_STATUS_ABORT;

    if (*bytes == 0)
        return FLAC__STREAM_DECODER_READ_STATUS_END_OF_STREAM;

    return FLAC__STREAM_DECODER_READ_STATUS_CONTINUE;
}

static FLAC__StreamDecoderSeekStatus flacSeekCb(
    const FLAC__StreamDecoder*,
    FLAC__uint64 absolute_byte_offset, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);

    if (_fseeki64(ctx->file, (long long)absolute_byte_offset, SEEK_SET) < 0)
        return FLAC__STREAM_DECODER_SEEK_STATUS_ERROR;

    return FLAC__STREAM_DECODER_SEEK_STATUS_OK;
}

static FLAC__StreamDecoderTellStatus flacTellCb(
    const FLAC__StreamDecoder*,
    FLAC__uint64* absolute_byte_offset, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    long long pos = _ftelli64(ctx->file);

    if (pos < 0)
        return FLAC__STREAM_DECODER_TELL_STATUS_ERROR;

    *absolute_byte_offset = (FLAC__uint64)pos;
    return FLAC__STREAM_DECODER_TELL_STATUS_OK;
}

static FLAC__StreamDecoderLengthStatus flacLengthCb(
    const FLAC__StreamDecoder*,
    FLAC__uint64* stream_length, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    long long cur = _ftelli64(ctx->file);

    if (cur < 0)
        return FLAC__STREAM_DECODER_LENGTH_STATUS_ERROR;

    if (_fseeki64(ctx->file, 0, SEEK_END) < 0)
        return FLAC__STREAM_DECODER_LENGTH_STATUS_ERROR;

    long long end = _ftelli64(ctx->file);
    _fseeki64(ctx->file, cur, SEEK_SET);

    if (end < 0)
        return FLAC__STREAM_DECODER_LENGTH_STATUS_ERROR;

    *stream_length = (FLAC__uint64)end;
    return FLAC__STREAM_DECODER_LENGTH_STATUS_OK;
}

static FLAC__bool flacEofCb(
    const FLAC__StreamDecoder*, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    return feof(ctx->file) ? true : false;
}

static FLAC__StreamDecoderWriteStatus flacWriteCb(
    const FLAC__StreamDecoder*, const FLAC__Frame*,
    const FLAC__int32* const [], void*) {
    return FLAC__STREAM_DECODER_WRITE_STATUS_CONTINUE;
}

static void flacErrorCb(
    const FLAC__StreamDecoder*,
    FLAC__StreamDecoderErrorStatus status, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    ctx->error = true;
    ctx->detail = FLAC__StreamDecoderErrorStatusString[status];
}

// ============================================================================
//  FLAC Validation
// ============================================================================
static bool validateFile(const std::wstring& path, std::wstring& reasonOut) {
    FlacCtx ctx;

    FLAC__StreamDecoder* dec = FLAC__stream_decoder_new();
    
    if (!dec) { 
        reasonOut = L"decoder alloc failed"; return false; 
    }

    FLAC__stream_decoder_set_md5_checking(dec, true);

    ctx.file = _wfopen(path.c_str(), L"rb");

    if (!ctx.file) {
        FLAC__stream_decoder_delete(dec);
        reasonOut = L"cannot open file";
        return false;
    }

    setvbuf(ctx.file, nullptr, _IOFBF, 1 << 20);

    auto st = FLAC__stream_decoder_init_stream(
        dec,
        flacReadCb, flacSeekCb, flacTellCb, flacLengthCb, flacEofCb,
        flacWriteCb, nullptr, flacErrorCb,
        &ctx);

    if (st != FLAC__STREAM_DECODER_INIT_STATUS_OK) {
        fclose(ctx.file);
        FLAC__stream_decoder_delete(dec);
        reasonOut = L"init failed";
        return false;
    }

    bool ok = FLAC__stream_decoder_process_until_end_of_stream(dec);
    bool md5 = FLAC__stream_decoder_finish(dec);

    // finish() does NOT touch our FILE* when using init_stream —
    // we have exclusive ownership and close it ourselves.
    fclose(ctx.file);
    FLAC__stream_decoder_delete(dec);

    if (!ok || ctx.error) {
        if (!ctx.detail.empty())
            reasonOut.assign(ctx.detail.begin(), ctx.detail.end());
        else
            reasonOut = L"decode error";

        return false;
    }

    if (!md5) {
        reasonOut = L"MD5 checksum mismatch";
        return false;
    }

    return true;
}

// ============================================================================
//  Worker Thread
// ============================================================================
static void workerMain(HWND hWnd) {
    // --- Scan phase ---
    g_files.clear();
    g_failures.clear();
    g_done = 0;  g_failed = 0;  g_bytes = 0;  g_total = 0;
    fs::path root(g_rootPath);

    try {
        for (const auto& e : fs::recursive_directory_iterator(root,
            fs::directory_options::skip_permission_denied)) {

            if (g_cancel) { 
                PostMessage(hWnd, WM_VAL_COMPLETE, 0, 0); return; 
            }

            if (e.is_regular_file()) {
                auto ext = e.path().extension().wstring();
                std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);

                if (ext == L".flac")
                    g_files.push_back(e.path().wstring());
            }
        }
    } catch (...) {}

    g_total = (uint32_t)g_files.size();
    PostMessage(hWnd, WM_SCAN_DONE, 0, 0);

    if (g_total == 0 || g_cancel) {
        PostMessage(hWnd, WM_VAL_COMPLETE, 0, 0);
        return;
    }

    // --- Validation phase ---
    g_t0 = std::chrono::steady_clock::now();
    std::atomic<uint32_t> idx{ 0 };

    auto work = [&]() {
        try {
            while (!g_cancel) {
                uint32_t i = idx.fetch_add(1);
                if (i >= g_total) break;

                const auto& fp = g_files[i];
                uint64_t sz = 0;
                try { 
                    sz = fs::file_size(fp); 
                } catch (...) {}

                std::wstring reason;
                bool ok = validateFile(fp, reason);

                if (g_cancel.load(std::memory_order_relaxed))
                    break;

                g_bytes += sz;
                g_done++;

                if (!ok) {
                    g_failed++;
                    std::wstring entry = fp + L"  [" + reason + L"]";
                    { std::lock_guard<std::mutex> lk(g_failMtx); g_failures.push_back(entry); }
                    PostMessage(hWnd, WM_VAL_FAILURE, 0, (LPARAM)_wcsdup(entry.c_str()));
                }
            }
        } catch (...) {
            // Prevent uncaught exceptions from calling std::terminate(),
            // which raises STATUS_STACK_BUFFER_OVERRUN (0xc0000409).
            g_failed++;
            std::wstring entry = L"(thread exception during validation)";
            { std::lock_guard<std::mutex> lk(g_failMtx); g_failures.push_back(entry); }
            PostMessage(hWnd, WM_VAL_FAILURE, 0, (LPARAM)_wcsdup(entry.c_str()));
        }
    };

    int n = std::min(g_nThreads, (int)g_total);
    std::vector<std::thread> threads;
    threads.reserve(n);
    
    for (int i = 0; i < n; i++) 
        threads.emplace_back(work);
    
    for (auto& t : threads) 
        t.join();

    PostMessage(hWnd, WM_VAL_COMPLETE, 0, 0);
}

// ============================================================================
//  UI Update
// ============================================================================
static void refreshStats() {
    uint32_t d = g_done.load();
    uint32_t f = g_failed.load();
    uint64_t b = g_bytes.load();
    double elapsed = std::chrono::duration<double>(
        std::chrono::steady_clock::now() - g_t0).count();
    double pct = g_total > 0 ? 100.0 * d / g_total : 0.0;
    double spd = elapsed > 0 ? (double)b / elapsed : 0.0;

    SendMessage(g_hProgressBar, PBM_SETPOS, (int)(pct * 10), 0);

    wchar_t buf[256];
    swprintf_s(buf, L"%.1f%%", pct);                 SetWindowText(g_hPctLabel, buf);
    swprintf_s(buf, L"Files:  %u / %u", d, g_total); SetWindowText(g_hFilesLbl, buf);
    swprintf_s(buf, L"Failed:  %u", f);              SetWindowText(g_hFailedLbl, buf);

    SetWindowText(g_hSpeedLbl, (L"Speed:  " + fmtSpeed(spd)).c_str());
    SetWindowText(g_hDataLbl, (L"Data:  " + fmtSize(b)).c_str());
    SetWindowText(g_hElapsedLbl, (L"Elapsed:  " + fmtTime((int)elapsed)).c_str());

    if (d > 0 && d < g_total) {
        double rem = elapsed * (g_total - d) / d;
        SetWindowText(g_hEtaLbl, (L"ETA:  " + fmtTime((int)rem)).c_str());
    } else if (d >= g_total) {
        SetWindowText(g_hEtaLbl, L"ETA:  00:00:00");
    } else {
        SetWindowText(g_hEtaLbl, L"ETA:  --:--:--");
    }

    // Update taskbar progress
    if (g_pTaskbar && g_total > 0) {
        g_pTaskbar->SetProgressValue(g_hWnd, d, g_total);
        
        if (g_failed.load() > 0)
            g_pTaskbar->SetProgressState(g_hWnd, TBPF_ERROR);
    }
}

static void setControlsEnabled(bool idle) {
    enableWithTabStop(g_hPathEdit, idle);
    enableWithTabStop(g_hBrowseBtn, idle);
    enableWithTabStop(g_hThreadsCombo, idle);
    enableWithTabStop(g_hStartBtn, idle);
    enableWithTabStop(g_hCancelBtn, !idle);
    enableWithTabStop(g_hExportBtn, idle);
    // g_hFailList always retains WS_TABSTOP
}

// ============================================================================
//  Actions — User-Initiated Operations
// ============================================================================
static void doBrowse(HWND hWnd) {
    wchar_t currentPath[MAX_PATH * 4] = {};
    GetWindowText(g_hPathEdit, currentPath, _countof(currentPath));

    std::wstring initialDir = FindNearestExistingDir(currentPath);

    BROWSEINFOW bi = {};
    bi.hwndOwner = hWnd;
    bi.lpszTitle = L"Select folder containing FLAC files";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    if (!initialDir.empty()) {
        bi.lpfn = BrowseCallbackProc;
        bi.lParam = (LPARAM)initialDir.c_str();
    }

    if (auto pidl = SHBrowseForFolder(&bi)) {
        wchar_t path[MAX_PATH];
        SHGetPathFromIDList(pidl, path);
        SetWindowText(g_hPathEdit, path);
        CoTaskMemFree(pidl);
        SaveLastPath(path);
    }
}

static void doStart() {
    wchar_t p[MAX_PATH * 4];
    GetWindowText(g_hPathEdit, p, _countof(p));

    if (!*p) {
        CenteredMessageBox(g_hWnd, L"Please select a directory.",
            L"FLAC Validator", MB_OK | MB_ICONWARNING);
        return;
    }

    if (!fs::exists(p)) {
        CenteredMessageBox(g_hWnd, L"Directory does not exist.",
            L"FLAC Validator", MB_OK | MB_ICONWARNING);
        return;
    }

    // Read settings from UI
    g_rootPath = p;
    int sel = (int)SendMessage(g_hThreadsCombo, CB_GETCURSEL, 0, 0);

    if (sel != CB_ERR) {
        wchar_t b[16];
        SendMessage(g_hThreadsCombo, CB_GETLBTEXT, sel, (LPARAM)b);
        g_nThreads = _wtoi(b);
    }

    g_cancel = false;
    g_running = true;
    g_t0 = std::chrono::steady_clock::now();

    // Reset UI
    SendMessage(g_hProgressBar, PBM_SETPOS, 0, 0);
    SetWindowText(g_hPctLabel, L"0.0%");
    SetWindowText(g_hFilesLbl, L"Files:  0 / 0");
    SetWindowText(g_hFailedLbl, L"Failed:  0");
    SetWindowText(g_hSpeedLbl, L"Speed:  --");
    SetWindowText(g_hDataLbl, L"Data:  --");
    SetWindowText(g_hElapsedLbl, L"Elapsed:  00:00:00");
    SetWindowText(g_hEtaLbl, L"ETA:  --:--:--");
    SendMessage(g_hFailList, LB_RESETCONTENT, 0, 0);
    SendMessage(g_hFailList, LB_ADDSTRING, 0, (LPARAM)L"Scanning for FLAC files...");

    setControlsEnabled(false);
    SetTimer(g_hWnd, TIMER_ID, TIMER_MS, nullptr);

    // Set taskbar to normal green progress
    if (g_pTaskbar)
        g_pTaskbar->SetProgressState(g_hWnd, TBPF_NORMAL);

    std::thread(workerMain, g_hWnd).detach();

    // Persist settings
    SaveLastPath(p);
    SaveThreads(g_nThreads);
}

static void doCancel() {
    g_cancel = true;
    enableWithTabStop(g_hCancelBtn, false);
}

static void doExport() {
    wchar_t fn[MAX_PATH] = L"flac_validation_results.txt";
    OPENFILENAME ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hWnd;
    ofn.lpstrFilter = L"Text Files\0*.txt\0All Files\0*.*\0";
    ofn.lpstrFile = fn;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_OVERWRITEPROMPT;
    ofn.lpstrDefExt = L"txt";

    if (GetSaveFileName(&ofn)) {
        std::wofstream f(fn);

        if (f.is_open()) {
            f << L"FLAC Validation Results\r\n";
            f << L"=======================\r\n\r\n";
            f << L"Path:       " << g_rootPath << L"\r\n";
            f << L"Total:      " << g_total << L"\r\n";
            f << L"Passed:     " << (g_total - g_failed.load()) << L"\r\n";
            f << L"Failed:     " << g_failed.load() << L"\r\n";
            f << L"Data:       " << fmtSize(g_bytes.load()) << L"\r\n\r\n";

            if (g_failures.empty()) {
                f << L"All files passed validation!\r\n";
            } else {
                f << L"Failed Files:\r\n";
                f << L"-------------\r\n";
                for (const auto& p : g_failures)
                    f << p << L"\r\n";
            }

            f.close();
            CenteredMessageBox(g_hWnd, L"Log exported successfully.",
                L"FLAC Validator", MB_OK | MB_ICONINFORMATION);
        }
    }
}

static void onThreadsChanged() {
    int sel = (int)SendMessage(g_hThreadsCombo, CB_GETCURSEL, 0, 0);

    if (sel != CB_ERR) {
        wchar_t b[16];
        SendMessage(g_hThreadsCombo, CB_GETLBTEXT, sel, (LPARAM)b);
        SaveThreads(_wtoi(b));
    }
}

// ============================================================================
//  Window Message Handlers
// ============================================================================
//  One free function per WM_* the dispatcher routes.  Keeps WndProc itself
//  reduced to a flat dispatch table that any reader can absorb at a glance.
// ============================================================================

static void onTaskbarButtonCreated() {
    if (g_pTaskbar) { 
        g_pTaskbar->Release(); 
        g_pTaskbar = nullptr; 
    }

    if (SUCCEEDED(CoCreateInstance(__uuidof(TaskbarList), nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&g_pTaskbar)))) {
        g_pTaskbar->HrInit();
    }
}

static void onCommand(HWND hWnd, WPARAM wParam) {
    switch (LOWORD(wParam)) {
        case ID_BROWSE_BTN:    
            doBrowse(hWnd); 
            break;
        case ID_START_BTN:     
            doStart();      
            break;
        case ID_CANCEL_BTN:    
            doCancel();     
            break;
        case ID_EXPORT_BTN:    
            doExport();     
            break;
        case ID_THREADS_COMBO:
            if (HIWORD(wParam) == CBN_SELCHANGE)
                onThreadsChanged();

            break;
    }
}

static void onTimer(WPARAM wParam) {
    if (wParam == TIMER_ID)
        refreshStats();
}

static void onScanDone() {
    SendMessage(g_hFailList, LB_RESETCONTENT, 0, 0);

    if (g_total == 0)
        SendMessage(g_hFailList, LB_ADDSTRING, 0,
            (LPARAM)L"No FLAC files found.");
    else
        SendMessage(g_hFailList, LB_ADDSTRING, 0,
            (LPARAM)L"No failures detected.");

    wchar_t b[128];
    swprintf_s(b, L"Files:  0 / %u", g_total);
    SetWindowText(g_hFilesLbl, b);
}

static void onValFailure(LPARAM lParam) {
    wchar_t* fp = (wchar_t*)lParam;

    // Replace the placeholder "No failures detected." on first real failure.
    // Use LB_GETTEXTLEN to avoid overflowing a fixed stack buffer.
    if (SendMessage(g_hFailList, LB_GETCOUNT, 0, 0) == 1) {
        int len = (int)SendMessage(g_hFailList, LB_GETTEXTLEN, 0, 0);

        if (len > 0 && len < 256) {
            wchar_t b[256] = {};
            SendMessage(g_hFailList, LB_GETTEXT, 0, (LPARAM)b);

            if (wcsstr(b, L"No failures"))
                SendMessage(g_hFailList, LB_RESETCONTENT, 0, 0);
        }
    }

    SendMessage(g_hFailList, LB_ADDSTRING, 0, (LPARAM)fp);
    free(fp);
}

static void onValComplete(HWND hWnd) {
    KillTimer(hWnd, TIMER_ID);
    refreshStats();
    g_running = false;

    if (g_pTaskbar) {
        if (g_cancel)
            g_pTaskbar->SetProgressState(g_hWnd, TBPF_PAUSED);     // Yellow
        else if (g_failed.load() > 0)
            g_pTaskbar->SetProgressState(g_hWnd, TBPF_ERROR);      // Red at 100%
        else
            g_pTaskbar->SetProgressState(g_hWnd, TBPF_NOPROGRESS); // Clear
    }

    if (g_failed.load() == 0 && g_total > 0) {
        SendMessage(g_hFailList, LB_RESETCONTENT, 0, 0);
        SendMessage(g_hFailList, LB_ADDSTRING, 0,
            g_cancel ? (LPARAM)L"Validation cancelled."
            : (LPARAM)L"No failures detected.");
    }

    setControlsEnabled(true);

    if (g_total == 0)
        enableWithTabStop(g_hExportBtn, false);

    if (!g_cancel && g_total > 0) {
        wchar_t m[256];
        swprintf_s(m, L"Validation complete!\n\n"
            L"Files:   %u\nPassed:  %u\nFailed:  %u",
            g_total, g_total - g_failed.load(), g_failed.load());
        CenteredMessageBox(hWnd, m, L"FLAC Validator",
            MB_OK | (g_failed.load() ? MB_ICONWARNING : MB_ICONINFORMATION));
    }
}

static LRESULT onClose(HWND hWnd) {
    if (g_running) {
        if (CenteredMessageBox(hWnd, L"Validation in progress. Cancel and exit?",
            L"FLAC Validator", MB_YESNO | MB_ICONQUESTION) != IDYES)
            return 0;

        g_cancel = true;
    }

    DestroyWindow(hWnd);
    return 0;
}

static void onDestroy() {
    g_cancel = true;
    PostQuitMessage(0);
}

// ============================================================================
//  Window Procedure — Dispatcher
// ============================================================================
static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == g_uTaskbarBtnCreated) {
        onTaskbarButtonCreated();
        return 0;
    }

    switch (msg) {
        case WM_COMMAND:      
            onCommand(hWnd, wParam);  
            return 0;
        case WM_TIMER:        
            onTimer(wParam);          
            return 0;
        case WM_SCAN_DONE:    
            onScanDone();             
            return 0;
        case WM_VAL_FAILURE:  
            onValFailure(lParam);     
            return 0;
        case WM_VAL_COMPLETE: 
            onValComplete(hWnd);      
            return 0;
        case WM_CLOSE:        
            return onClose(hWnd);
        case WM_DESTROY:      
            onDestroy();              
            return 0;
        default:              
            return DefWindowProc(hWnd, msg, wParam, lParam);
    }
}

// ============================================================================
//  Layout Metrics
// ============================================================================
struct LayoutMetrics {
    int x, cw;
    int lblW, btnW, editW, halfW;
    int yRow1, yRow2, yRow3, yRow4, yRow5, yRow6;
    int yFailGrp, yButtons;
    int pbW, grpH;
    int bw, bg, bx;
};

static LayoutMetrics computeLayout() {
    LayoutMetrics L{};
    L.x = MARGIN;
    L.cw = WIN_W - 2 * MARGIN;
    L.lblW = 55;
    L.btnW = 75;
    L.editW = L.cw - L.lblW - L.btnW - 10;
    L.halfW = L.cw / 2;

    L.yRow1 = MARGIN;
    L.yRow2 = L.yRow1 + ROW_H + GAP + 5;
    L.yRow3 = L.yRow2 + ROW_H + GAP + 10;
    L.pbW = L.cw - 60;
    L.yRow4 = L.yRow3 + ROW_H + GAP + 5;
    L.yRow5 = L.yRow4 + ROW_H + GAP;
    L.yRow6 = L.yRow5 + ROW_H + GAP;
    L.yFailGrp = L.yRow6 + ROW_H + GAP + 5;
    L.grpH = 120;
    L.yButtons = L.yFailGrp + L.grpH + GAP + 5;
    L.bw = 90;
    L.bg = 10;
    L.bx = L.x + (L.cw - 3 * L.bw - 2 * L.bg) / 2;
    return L;
}

// ============================================================================
//  Control Creation
// ============================================================================
//  Controls are created in TAB ORDER.  Win32 tab navigation follows
//  child z-order, which matches creation order.  All interactive
//  controls receive WS_TABSTOP; initially disabled controls omit it.
//  Non-interactive controls (labels, progress bar, group box) are
//  created after all tabbable controls so they never receive focus.
//
//  Tab Order:
//    1. Path Edit          (always tabbable when idle)        — createInputControls
//    2. Browse Button      (always tabbable when idle)        — createInputControls
//    3. Threads Combo      (always tabbable when idle)        — createInputControls
//    4. Start Button       (always tabbable when idle)        — createActionButtons
//    5. Cancel Button      (tabbable only when running)       — createActionButtons
//    6. Export Log Button  (tabbable only when idle + results) — createActionButtons
//    7. Failures Listbox   (always tabbable)                  — createFailuresGroup
// ============================================================================

static void createInputControls(HWND hWnd, const LayoutMetrics& L) {
    // Tab 1: Path Edit
    g_hPathEdit = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
        L.x + L.lblW, L.yRow1, L.editW, ROW_H,
        hWnd, (HMENU)ID_PATH_EDIT, g_hInst, nullptr);

    // Tab 2: Browse Button
    g_hBrowseBtn = CreateWindow(L"BUTTON", L"Browse",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        L.x + L.lblW + L.editW + 5, L.yRow1, L.btnW, ROW_H,
        hWnd, (HMENU)ID_BROWSE_BTN, g_hInst, nullptr);

    // Tab 3: Threads Combo
    g_hThreadsCombo = CreateWindow(L"COMBOBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
        L.x + L.lblW, L.yRow2, 60, 200,
        hWnd, (HMENU)ID_THREADS_COMBO, g_hInst, nullptr);
    for (auto s : { L"1", L"2", L"4", L"8", L"16", L"32" })
        SendMessage(g_hThreadsCombo, CB_ADDSTRING, 0, (LPARAM)s);
    SendMessage(g_hThreadsCombo, CB_SETCURSEL, 3, 0);  // default: 8
}

static void createActionButtons(HWND hWnd, const LayoutMetrics& L) {
    // Tab 4: Start Button
    g_hStartBtn = CreateWindow(L"BUTTON", L"Start",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        L.bx, L.yButtons, L.bw, 30,
        hWnd, (HMENU)ID_START_BTN, g_hInst, nullptr);

    // Tab 5: Cancel Button (initially disabled — no WS_TABSTOP)
    g_hCancelBtn = CreateWindow(L"BUTTON", L"Cancel",
        WS_CHILD | WS_VISIBLE | WS_DISABLED,
        L.bx + L.bw + L.bg, L.yButtons, L.bw, 30,
        hWnd, (HMENU)ID_CANCEL_BTN, g_hInst, nullptr);

    // Tab 6: Export Button (initially disabled — no WS_TABSTOP)
    g_hExportBtn = CreateWindow(L"BUTTON", L"Export Log",
        WS_CHILD | WS_VISIBLE | WS_DISABLED,
        L.bx + 2 * (L.bw + L.bg), L.yButtons, L.bw, 30,
        hWnd, (HMENU)ID_EXPORT_BTN, g_hInst, nullptr);
}

static void createFailuresGroup(HWND hWnd, const LayoutMetrics& L) {
    // Group box (not tabbable, must precede listbox for painting)
    g_hFailGrp = CreateWindow(L"BUTTON", L" Failures ",
        WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
        L.x, L.yFailGrp, L.cw, L.grpH,
        hWnd, nullptr, g_hInst, nullptr);

    // Tab 7: Failures Listbox (always tabbable)
    g_hFailList = CreateWindowEx(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT,
        L.x + 10, L.yFailGrp + 18, L.cw - 20, L.grpH - 28,
        hWnd, (HMENU)ID_FAILURES_LIST, g_hInst, nullptr);
    SendMessage(g_hFailList, LB_SETHORIZONTALEXTENT, 2000, 0);
    SendMessage(g_hFailList, LB_ADDSTRING, 0, (LPARAM)L"No failures detected.");
    SetWindowSubclass(g_hFailList, FailListSubclassProc, 0, 0);
}

static void createDecorativeControls(HWND hWnd, const LayoutMetrics& L) {
    // Path label
    CreateWindow(L"STATIC", L"Path:", WS_CHILD | WS_VISIBLE,
        L.x, L.yRow1 + 3, L.lblW, ROW_H, hWnd, nullptr, g_hInst, nullptr);

    // Threads label
    CreateWindow(L"STATIC", L"Threads:", WS_CHILD | WS_VISIBLE,
        L.x, L.yRow2 + 3, L.lblW, ROW_H, hWnd, nullptr, g_hInst, nullptr);

    // Progress bar
    g_hProgressBar = CreateWindowEx(0, PROGRESS_CLASS, L"",
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        L.x, L.yRow3, L.pbW, ROW_H,
        hWnd, (HMENU)ID_PROGRESS_BAR, g_hInst, nullptr);
    SendMessage(g_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 1000));

    g_hPctLabel = CreateWindow(L"STATIC", L"0.0%",
        WS_CHILD | WS_VISIBLE | SS_RIGHT,
        L.x + L.pbW + 5, L.yRow3 + 3, 50, ROW_H, hWnd, nullptr, g_hInst, nullptr);

    // Row 4 — Files / Failed
    g_hFilesLbl = CreateWindow(L"STATIC", L"Files:  0 / 0",
        WS_CHILD | WS_VISIBLE, L.x, L.yRow4, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);
    g_hFailedLbl = CreateWindow(L"STATIC", L"Failed:  0",
        WS_CHILD | WS_VISIBLE, L.x + L.halfW, L.yRow4, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);

    // Row 5 — Speed / Data
    g_hSpeedLbl = CreateWindow(L"STATIC", L"Speed:  --",
        WS_CHILD | WS_VISIBLE, L.x, L.yRow5, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);
    g_hDataLbl = CreateWindow(L"STATIC", L"Data:  --",
        WS_CHILD | WS_VISIBLE, L.x + L.halfW, L.yRow5, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);

    // Row 6 — Elapsed / ETA
    g_hElapsedLbl = CreateWindow(L"STATIC", L"Elapsed:  00:00:00",
        WS_CHILD | WS_VISIBLE, L.x, L.yRow6, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);
    g_hEtaLbl = CreateWindow(L"STATIC", L"ETA:  --:--:--",
        WS_CHILD | WS_VISIBLE, L.x + L.halfW, L.yRow6, L.halfW, ROW_H,
        hWnd, nullptr, g_hInst, nullptr);
}

static void applyControlFont(HWND hWnd) {
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    EnumChildWindows(hWnd, [](HWND ch, LPARAM lp) -> BOOL {
        SendMessage(ch, WM_SETFONT, lp, TRUE);
        return TRUE;
        }, (LPARAM)hFont);
}

static void createControls(HWND hWnd) {
    LayoutMetrics L = computeLayout();
    createInputControls(hWnd, L);       // Tabs 1–3
    createActionButtons(hWnd, L);       // Tabs 4–6
    createFailuresGroup(hWnd, L);       // Tab 7
    createDecorativeControls(hWnd, L);  // Non-tabbable
    applyControlFont(hWnd);
}

// ============================================================================
//  Application Lifecycle
// ============================================================================

static void initApp(HINSTANCE hInst) {
    InitIniPath(hInst);
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
}

static void registerMainWindowClass(HINSTANCE hInst) {
    WNDCLASSEX wc = { sizeof(wc) };
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.lpszClassName = L"FlacValClass";
    wc.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_FLACVALIDATORGUI));
    wc.hIconSm = LoadIcon(hInst, MAKEINTRESOURCE(IDI_FLACVALIDATORGUI));
    RegisterClassEx(&wc);
}

static HWND createMainWindow(HINSTANCE hInst) {
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    RECT rc = { 0, 0, WIN_W, WIN_H };
    AdjustWindowRectEx(&rc, style, FALSE, 0);
    int ww = rc.right - rc.left;
    int wh = rc.bottom - rc.top;
    int sx = GetSystemMetrics(SM_CXSCREEN);
    int sy = GetSystemMetrics(SM_CYSCREEN);

    return CreateWindowEx(0, L"FlacValClass", L"FLAC Validator v1.0", style,
        (sx - ww) / 2, (sy - wh) / 2, ww, wh,
        nullptr, nullptr, hInst, nullptr);
}

static void initTaskbarMessage(HWND hWnd) {
    // Registered message arrives once Explorer has built the taskbar
    // button for our HWND — only then is ITaskbarList3 effective.
    // ChangeWindowMessageFilterEx allows the broadcast through UIPI
    // even if this process is launched elevated.
    g_uTaskbarBtnCreated = RegisterWindowMessageW(L"TaskbarButtonCreated");
    ChangeWindowMessageFilterEx(hWnd, g_uTaskbarBtnCreated, MSGFLT_ALLOW, nullptr);
}

static bool parseCommandLine() {
    bool autoStart = false;
    int  argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    if (argv) {
        if (argc >= 2 && fs::exists(argv[1]) && fs::is_directory(argv[1])) {
            SetWindowTextW(g_hPathEdit, argv[1]);
            autoStart = true;
        }

        LocalFree(argv);
    }

    return autoStart;
}

static void restoreSettings(bool autoStart) {
    // Path: only restored when not driven by command line.
    if (!autoStart) {
        WCHAR lastPath[MAX_PATH];
        LoadLastPath(lastPath, MAX_PATH);
        if (lastPath[0])
            SetWindowTextW(g_hPathEdit, lastPath);
    }

    // Threads: always restored.
    int savedThreads = LoadThreads(8);
    WCHAR tb[16];
    swprintf_s(tb, L"%d", savedThreads);
    int idx = (int)SendMessage(g_hThreadsCombo, CB_FINDSTRINGEXACT, -1, (LPARAM)tb);

    if (idx != CB_ERR)
        SendMessage(g_hThreadsCombo, CB_SETCURSEL, idx, 0);
}

static int runMessageLoop() {
    MSG m;

    while (GetMessage(&m, nullptr, 0, 0)) {
        if (IsDialogMessage(g_hWnd, &m))
            continue;

        TranslateMessage(&m);
        DispatchMessage(&m);
    }

    return (int)m.wParam;
}

static void shutdownApp() {
    if (g_pTaskbar) {
        g_pTaskbar->Release();
        g_pTaskbar = nullptr;
    }

    CoUninitialize();
}

// ============================================================================
//  Entry Point
// ============================================================================
int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nShow) {
    g_hInst = hInst;

    initApp(hInst);
    registerMainWindowClass(hInst);
    g_hWnd = createMainWindow(hInst);
    initTaskbarMessage(g_hWnd);
    createControls(g_hWnd);

    ShowWindow(g_hWnd, nShow);
    UpdateWindow(g_hWnd);

    bool autoStart = parseCommandLine();
    restoreSettings(autoStart);

    if (autoStart)
        PostMessage(g_hWnd, WM_COMMAND, MAKEWPARAM(ID_START_BTN, BN_CLICKED), 0);

    int rc = runMessageLoop();
    shutdownApp();
    return rc;
}