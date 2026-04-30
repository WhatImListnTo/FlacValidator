/*
 * FlacValidator - High-Performance Parallel FLAC Integrity Validator
 *
 * Uses libFLAC directly (no process spawning) with multi-threaded
 * validation for maximum throughput on modern hardware.
 *
 * Build:  Visual Studio 2022 | C++17 | x64 Release
 * Deps:   libFLAC 1.5.0 via vcpkg
 * Usage:  FlacValidator.exe "E:\Music" --threads 16
 */

#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <FLAC/stream_decoder.h>
#include <filesystem>
#include <vector>
#include <string>
#include <thread>
#include <atomic>
#include <mutex>
#include <chrono>
#include <algorithm>
#include <fstream>
#include <cstdio>
#include <cstdint>

namespace fs = std::filesystem;

// ================================================================
//  Shared State
// ================================================================
static std::vector<fs::path>   g_files;
static std::atomic<size_t>     g_nextIdx{ 0 };
static std::atomic<uint64_t>   g_processed{ 0 };
static std::atomic<uint64_t>   g_passed{ 0 };
static std::atomic<uint64_t>   g_failed{ 0 };
static std::atomic<uint64_t>   g_bytes{ 0 };
static std::atomic<bool>       g_done{ false };
static uint64_t                g_totalFiles = 0;

struct Failure {
    std::string path;
    std::string reason;
};

static std::mutex              g_failMtx;
static std::vector<Failure>    g_failures;

// ================================================================
//  FLAC Decoder Context
// ================================================================
//  Single context struct passed as client_data to ALL decoder
//  callbacks.  Using init_stream (not init_FILE) gives us exclusive
//  ownership of the FILE* — no ambiguity about whether finish()
//  closes it.
// ================================================================
struct FlacCtx {
    FILE* file = nullptr;
    bool        error = false;
    std::string detail;
};

// ================================================================
//  FLAC Stream Callbacks
// ================================================================
static FLAC__StreamDecoderReadStatus flac_read_cb(
    const FLAC__StreamDecoder*, FLAC__byte buffer[],
    size_t* bytes, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);

    if (*bytes == 0)
        return FLAC__STREAM_DECODER_READ_STATUS_ABORT;

    *bytes = fread(buffer, 1, *bytes, ctx->file);

    if (ferror(ctx->file))
        return FLAC__STREAM_DECODER_READ_STATUS_ABORT;

    if (*bytes == 0)
        return FLAC__STREAM_DECODER_READ_STATUS_END_OF_STREAM;

    return FLAC__STREAM_DECODER_READ_STATUS_CONTINUE;
}

static FLAC__StreamDecoderSeekStatus flac_seek_cb(
    const FLAC__StreamDecoder*,
    FLAC__uint64 absolute_byte_offset, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);

    if (_fseeki64(ctx->file, (long long)absolute_byte_offset, SEEK_SET) < 0)
        return FLAC__STREAM_DECODER_SEEK_STATUS_ERROR;

    return FLAC__STREAM_DECODER_SEEK_STATUS_OK;
}

static FLAC__StreamDecoderTellStatus flac_tell_cb(
    const FLAC__StreamDecoder*,
    FLAC__uint64* absolute_byte_offset, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    long long pos = _ftelli64(ctx->file);

    if (pos < 0)
        return FLAC__STREAM_DECODER_TELL_STATUS_ERROR;

    *absolute_byte_offset = (FLAC__uint64)pos;
    return FLAC__STREAM_DECODER_TELL_STATUS_OK;
}

static FLAC__StreamDecoderLengthStatus flac_length_cb(
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

static FLAC__bool flac_eof_cb(
    const FLAC__StreamDecoder*, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    return feof(ctx->file) ? true : false;
}

static FLAC__StreamDecoderWriteStatus flac_write_cb(
    const FLAC__StreamDecoder*, const FLAC__Frame*,
    const FLAC__int32* const [], void*) {
    return FLAC__STREAM_DECODER_WRITE_STATUS_CONTINUE;
}

static void flac_error_cb(
    const FLAC__StreamDecoder*,
    FLAC__StreamDecoderErrorStatus status, void* client_data) {
    auto* ctx = static_cast<FlacCtx*>(client_data);
    ctx->error = true;
    ctx->detail = FLAC__StreamDecoderErrorStatusString[status];
}

// ================================================================
//  INI File Persistence (shared with GUI)
// ================================================================
static wchar_t g_iniPath[MAX_PATH] = {};

static void initIniPath() {
    GetModuleFileNameW(nullptr, g_iniPath, MAX_PATH);
    wchar_t* slash = wcsrchr(g_iniPath, L'\\');

    if (slash) {
        *(slash + 1) = L'\0';
        wcscat_s(g_iniPath, MAX_PATH, L"FlacValidator.ini");
    }
}

static std::wstring loadIniPathW() {
    wchar_t buf[MAX_PATH];
    GetPrivateProfileStringW(L"Settings", L"LastPath", L"", buf, MAX_PATH, g_iniPath);
    return std::wstring(buf);
}

static int loadIniThreads(int defaultVal) {
    return GetPrivateProfileIntW(L"Settings", L"Threads", defaultVal, g_iniPath);
}

// ================================================================
//  Raw Command Line Tokenizer
// ================================================================
//  Bypasses the MSVC CRT's argc/argv parser which treats \"
//  as an escaped quote.  This corrupts paths like "E:\" where
//  the trailing backslash + quote is misinterpreted.
//
//  Since the double-quote character is illegal in Windows file
//  paths, we treat " strictly as a delimiter — never escapable.
// ================================================================
static std::vector<std::wstring> tokenizeCommandLine(const wchar_t* cmdLine) {
    std::vector<std::wstring> tokens;

    while (*cmdLine)
    {
        // Skip whitespace
        while (*cmdLine == L' ' || *cmdLine == L'\t')
            cmdLine++;

        if (!*cmdLine)
            break;

        std::wstring token;
        bool inQuote = false;

        while (*cmdLine) {
            if (*cmdLine == L'"') {
                inQuote = !inQuote;
                cmdLine++;
            } else if ((*cmdLine == L' ' || *cmdLine == L'\t') && !inQuote) {
                break;
            } else {
                token += *cmdLine++;
            }
        }

        tokens.push_back(token);
    }

    return tokens;
}

// ================================================================
//  Wide-to-UTF8 Conversion
// ================================================================
static std::string wideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, s.data(), n, nullptr, nullptr);
    return s;
}

// ================================================================
//  FLAC Validation (per file)
// ================================================================
static bool validate_flac(const fs::path& filepath, std::string& errorOut) {
    FlacCtx ctx;

    FLAC__StreamDecoder* dec = FLAC__stream_decoder_new();

    if (!dec) {
        errorOut = "decoder alloc failed";
        return false;
    }

    FLAC__stream_decoder_set_md5_checking(dec, true);

    // Use wide-char fopen for full Unicode path support on Windows
    ctx.file = _wfopen(filepath.wstring().c_str(), L"rb");

    if (!ctx.file) {
        FLAC__stream_decoder_delete(dec);
        errorOut = "cannot open file";
        return false;
    }

    // Large read buffer for better sequential throughput
    setvbuf(ctx.file, nullptr, _IOFBF, 1 << 20);  // 1 MB

    FLAC__StreamDecoderInitStatus initSt =
        FLAC__stream_decoder_init_stream(
            dec,
            flac_read_cb, flac_seek_cb, flac_tell_cb,
            flac_length_cb, flac_eof_cb,
            flac_write_cb, nullptr, flac_error_cb,
            &ctx);

    if (initSt != FLAC__STREAM_DECODER_INIT_STATUS_OK) {
        fclose(ctx.file);
        FLAC__stream_decoder_delete(dec);
        errorOut = FLAC__StreamDecoderInitStatusString[initSt];
        return false;
    }

    // Full decode pass - validates frame CRCs during decode
    bool ok = FLAC__stream_decoder_process_until_end_of_stream(dec);

    // finish() returns false if MD5 checksum mismatch.
    // With init_stream, finish() does NOT touch our FILE*.
    bool md5 = FLAC__stream_decoder_finish(dec);

    // We have exclusive ownership — close it ourselves.
    fclose(ctx.file);
    FLAC__stream_decoder_delete(dec);

    if (!ok || ctx.error) {
        errorOut = ctx.detail.empty() ? "decode error" : ctx.detail;
        return false;
    }

    if (!md5) {
        errorOut = "MD5 checksum mismatch";
        return false;
    }

    return true;
}

// ================================================================
//  Log Path Helper
// ================================================================
static std::string getDefaultLogPath() {
    const char* profile = getenv("USERPROFILE");
    std::string dir = profile ? std::string(profile) + "\\Logs" : ".";
    fs::create_directories(dir);
    return dir + "\\flac_failures.log";
}

// ================================================================
//  Worker Thread
// ================================================================
static void worker_thread() {
    try {
        while (true) {
            size_t i = g_nextIdx.fetch_add(1, std::memory_order_relaxed);

            if (i >= g_files.size())
                break;

            uint64_t sz = 0;

            try { 
                sz = fs::file_size(g_files[i]); 
            } catch (...) {}

            std::string err;
            bool ok = validate_flac(g_files[i], err);

            g_bytes.fetch_add(sz, std::memory_order_relaxed);
            g_processed.fetch_add(1, std::memory_order_relaxed);

            if (ok) {
                g_passed.fetch_add(1, std::memory_order_relaxed);
            } else {
                g_failed.fetch_add(1, std::memory_order_relaxed);
                std::lock_guard<std::mutex> lk(g_failMtx);
                g_failures.push_back({ g_files[i].string(), err });
            }
        }
    } catch (const std::exception& ex) {
        g_failed.fetch_add(1, std::memory_order_relaxed);
        g_processed.fetch_add(1, std::memory_order_relaxed);
        std::lock_guard<std::mutex> lk(g_failMtx);
        g_failures.push_back({ "(thread exception)", ex.what() });
    } catch (...) {
        g_failed.fetch_add(1, std::memory_order_relaxed);
        g_processed.fetch_add(1, std::memory_order_relaxed);
        std::lock_guard<std::mutex> lk(g_failMtx);
        g_failures.push_back({ "(thread exception)", "unknown error" });
    }
}

// ================================================================
//  Progress Reporter Thread
// ================================================================
static void progress_thread(std::chrono::steady_clock::time_point t0) {
    while (!g_done.load(std::memory_order_relaxed)) {
        std::this_thread::sleep_for(std::chrono::seconds(2));

        double sec = std::chrono::duration<double>(
            std::chrono::steady_clock::now() - t0).count();
        uint64_t p = g_processed.load(std::memory_order_relaxed);
        uint64_t b = g_bytes.load(std::memory_order_relaxed);
        uint64_t fl = g_failed.load(std::memory_order_relaxed);

        double pct = g_totalFiles ? 100.0 * p / g_totalFiles : 0;
        double mbps = sec > 0 ? (b / 1048576.0) / sec : 0;
        double eta = (p > 0 && p < g_totalFiles)
            ? sec * (g_totalFiles - p) / p : 0;

        printf("\r  [%llu/%llu] %5.1f%%  | %7.1f MB/s | Failed: %llu | ETA: %d:%02d   ",
            (unsigned long long)p,
            (unsigned long long)g_totalFiles,
            pct, mbps,
            (unsigned long long)fl,
            (int)eta / 60, (int)eta % 60);
        fflush(stdout);
    }
}

// ================================================================
//  Main
// ================================================================
int main(int argc, char* argv[]) {

    // Bypass argc/argv — the MSVC CRT parser treats \" as an escaped
    // quote, which corrupts paths like "E:\" into E:" .  Our custom
    // tokenizer treats " strictly as a delimiter since double-quote
    // is illegal in Windows file paths.
    (void)argc;
    (void)argv;

    initIniPath();

    auto args = tokenizeCommandLine(GetCommandLineW());

    fs::path     root;
    unsigned     threads = 0;
    bool         pathSet = false;
    bool         threadsSet = false;
    std::string  logPath = getDefaultLogPath();

    // Argument parsing (skip args[0] which is the executable)
    for (size_t i = 1; i < args.size(); i++) {
        const std::wstring& a = args[i];

        if ((a == L"-t" || a == L"--threads") && i + 1 < args.size()) {
            threads = (unsigned)std::stoi(wideToUtf8(args[++i]));
            threadsSet = true;
        } else if ((a == L"-l" || a == L"--log") && i + 1 < args.size()) {
            logPath = wideToUtf8(args[++i]);
        } else if (a == L"-h" || a == L"--help") {
            puts("Usage: FlacValidator [path] [--threads N] [--log file.log]");
            return 0;
        } else {
            root = args[i];   // fs::path from wide string preserves Unicode
            pathSet = true;
        }
    }

    // Fall back to INI values (written by the GUI), then to defaults
    if (!pathSet) {
        std::wstring iniPath = loadIniPathW();

        if (!iniPath.empty() && fs::exists(iniPath))
            root = iniPath;
        else
            root = fs::current_path();
    }

    if (!threadsSet) {
        int iniThreads = loadIniThreads(0);

        if (iniThreads > 0)
            threads = (unsigned)iniThreads;
        else
            threads = std::thread::hardware_concurrency();
    }

    if (!threads) {
        threads = 8;
    }

    // Validate root path
    if (!fs::exists(root) || !fs::is_directory(root)) {
        printf("Error: path does not exist or is not a directory:\n  %s\n",
            root.string().c_str());
        return 1;
    }

    // Header
    printf("========================================\n");
    printf("  FLAC Integrity Validator v1.0\n");
    printf("  libFLAC %s\n", FLAC__VERSION_STRING);
    printf("========================================\n");
    printf("  Path:    %s\n", root.string().c_str());
    printf("  Threads: %u\n", threads);
    printf("  Log:     %s\n", logPath.c_str());
    printf("========================================\n\n");

    // ── Scan for FLAC files ──────────────────────────────────
    printf("  Scanning for FLAC files...");
    fflush(stdout);

    auto scanT0 = std::chrono::steady_clock::now();

    for (auto& entry : fs::recursive_directory_iterator(
        root, fs::directory_options::skip_permission_denied)) {

        if (!entry.is_regular_file())
            continue;

        auto ext = entry.path().extension().string();
        std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

        if (ext == ".flac")
            g_files.push_back(entry.path());
    }

    // Sort for sequential disk access patterns
    std::sort(g_files.begin(), g_files.end());
    g_totalFiles = g_files.size();

    double scanSec = std::chrono::duration<double>(
        std::chrono::steady_clock::now() - scanT0).count();
    printf(" found %llu files (%.1fs)\n\n", (unsigned long long)g_totalFiles, scanSec);

    if (g_totalFiles == 0) {
        puts("  No FLAC files found.");
        return 0;
    }

    // ── Validate ─────────────────────────────────────────────
    printf("  Validating...\n\n");

    auto t0 = std::chrono::steady_clock::now();
    std::thread progThr(progress_thread, t0);
    std::vector<std::thread> workers;
    workers.reserve(threads);

    for (unsigned i = 0; i < threads; i++)
        workers.emplace_back(worker_thread);

    for (auto& w : workers)
        w.join();

    g_done.store(true, std::memory_order_relaxed);
    progThr.join();

    // ── Summary ──────────────────────────────────────────────
    double elapsed = std::chrono::duration<double>(
        std::chrono::steady_clock::now() - t0).count();
    double totalGB = g_bytes.load() / (1024.0 * 1024.0 * 1024.0);
    double mbps = elapsed > 0 ? (g_bytes.load() / 1048576.0) / elapsed : 0;
    int eH = (int)elapsed / 3600;
    int eM = ((int)elapsed % 3600) / 60;
    int eS = (int)elapsed % 60;

    printf("\n\n========================================\n");
    printf("  SUMMARY\n");
    printf("========================================\n");
    printf("  Total files:  %llu\n", (unsigned long long)g_totalFiles);
    printf("  Passed:       %llu\n", (unsigned long long)g_passed.load());
    printf("  Failed:       %llu\n", (unsigned long long)g_failed.load());
    printf("  Data:         %.2f GB\n", totalGB);
    printf("  Throughput:   %.1f MB/s\n", mbps);
    printf("  Elapsed:      %02d:%02d:%02d\n", eH, eM, eS);
    printf("  Threads:      %u\n", threads);
    printf("========================================\n\n");

    // ── Failure report ───────────────────────────────────────
    if (!g_failures.empty()) {
        std::ofstream log(logPath);
        printf("  FAILURES:\n");

        for (auto& f : g_failures) {
            printf("    FAIL: %s [%s]\n", f.path.c_str(), f.reason.c_str());
            log << f.path << " | " << f.reason << "\n";
        }

        printf("\n  Failure log written to: %s\n", logPath.c_str());
    } else {
        printf("  All files passed validation!\n");
    }

    return g_failures.empty() ? 0 : 1;
}