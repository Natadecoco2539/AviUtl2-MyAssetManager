// ------------------------------------------------------------
// MyAssetManager.cpp
// v1.5
// ------------------------------------------------------------

#define NOMINMAX
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <vector>
#include <string>
#include <algorithm>
#include <sstream>
#include <set>
#include <map>
#include <cctype>
#include <climits>
#include <fstream>
#include <functional>

// GDI+ Headers
#include <objidl.h>
#include <gdiplus.h>
#include "../sdk/output2.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "uxtheme.lib")

// 関数名を強制エクスポート
#if defined(_M_IX86)
  #pragma comment(linker, "/EXPORT:RegisterPlugin=_RegisterPlugin")
#elif defined(_M_AMD64)
  #pragma comment(linker, "/EXPORT:RegisterPlugin")
#endif

using namespace Gdiplus;

// --- ID定義 ---
#define WM_REFRESH_ASSETS       (WM_USER + 100)
#define ID_TIMER_HOVER          998
#define ID_TIMER_ADD_PREVIEW    997 
#define ID_TIMER_TOOLTIP        996 
#define ID_TIMER_GIF_HINT       995
#define ID_TIMER_GIF_HINT_POLL  994

#define ID_COMBO_CATEGORY 200
#define ID_BTN_FAV_SORT   201
#define ID_BTN_REFRESH    202
#define ID_EDIT_SEARCH    203
#define ID_BTN_SNIP       204 
#define ID_BTN_GIF_EXPORT 205

#define IDM_FAVORITE      300
#define IDM_EDIT          301
#define IDM_DELETE        302
#define IDM_SETTINGS      303
#define IDM_TOGGLE_FIXED  304 
#define IDM_OPEN_FOLDER   305

#define IDC_EDIT_NAME     101
#define IDC_COMBO_CAT     102
#define ID_BTN_SAVE       103
#define ID_BTN_CANCEL     104
#define IDC_SLIDER_SPEED  105
#define IDC_EDIT_SPEED    106 
#define IDC_CHK_GIF_ORIGINAL 107

#define ID_BTN_MSG_OK     401
#define ID_BTN_MSG_YES    402
#define ID_BTN_MSG_NO     403
#define ID_BTN_GUIDE_OK   404
#define ID_BTN_GUIDE_CANCEL 405
#define IDC_CHK_GUIDE_HIDE 406

// --- UI定数 ---
static constexpr int TITLE_H        = 32;
static constexpr int FOOTER_H       = 40;
static constexpr int ITEM_HEIGHT    = 90; 
static constexpr int THUMB_W        = 120;
static constexpr int THUMB_H        = 68;
static constexpr int SCROLL_SPD     = 30;
static constexpr int MIN_ITEM_WIDTH = 240;
static constexpr int RESIZE_MARGIN  = 6;

static constexpr COLORREF COL_BG        = RGB(30, 30, 30);
static constexpr COLORREF COL_TITLE_BG  = RGB(45, 45, 45);
static constexpr COLORREF COL_BORDER    = RGB(60, 60, 60);
static constexpr COLORREF COL_ITEM_BG   = RGB(40, 40, 40);
static constexpr COLORREF COL_ITEM_SEL  = RGB(60, 70, 90);
static constexpr COLORREF COL_TEXT      = RGB(220, 220, 220);
static constexpr COLORREF COL_SUBTEXT   = RGB(160, 160, 160);
static constexpr COLORREF COL_FOOTER    = RGB(40, 40, 40);
static constexpr COLORREF COL_BTN_BG    = RGB(60, 60, 60);
static constexpr COLORREF COL_BTN_PUSH  = RGB(100, 100, 100);
static constexpr COLORREF COL_BTN_ACT   = RGB(100, 120, 200);
static constexpr COLORREF COL_INPUT_BG  = RGB(20, 20, 20);
static constexpr COLORREF COL_FIXED_TEXT = RGB(31, 205, 219); 

static constexpr COLORREF COL_TIP_BG     = RGB(50, 50, 50);
static constexpr COLORREF COL_TIP_BORDER = RGB(100, 100, 100);
static constexpr COLORREF COL_TIP_TEXT   = RGB(230, 230, 230);

// --- 構造体 ---
struct Asset {
    std::wstring name; 
    std::wstring path; 
    std::wstring imagePath; 
    std::wstring category;
    bool isFavorite; 
    bool isFixedFrame; 
    bool isMulti;      
    Image* pImage; 
    UINT frameCount; 
    UINT* frameDelays; 
    int currentFrame; 
    PropertyItem* pPropertyItem;
};

struct EDIT_SECTION_SAFE {
    void* info; void* f1; void* f2; void* f3; void* f4;
    LPCSTR (*get_object_alias)(void* object);
    void* f6; void* f7; void* f8; void* f9;
    void* f10; void* f11; void* f12;
    void* (*get_selected_object)(int index);
    int (*get_selected_object_num)();
};

struct OBJECT_LAYER_FRAME_SAFE {
    int layer;
    int start;
    int end;
};

struct EDIT_SECTION_OUT_SAFE {
    void* info;
    void* create_object_from_alias;
    void* find_object;
    void* count_object_effect;
    OBJECT_LAYER_FRAME_SAFE (*get_object_layer_frame)(void* object);
    LPCSTR (*get_object_alias)(void* object);
    LPCSTR (*get_object_item_value)(void* object, LPCWSTR effect, LPCWSTR item);
    bool (*set_object_item_value)(void* object, LPCWSTR effect, LPCWSTR item, LPCSTR value);
    bool (*move_object)(void* object, int layer, int frame);
    void (*delete_object)(void* object);
    void* (*get_focus_object)();
    void (*set_focus_object)(void* object);
    void* (*get_project_file)(void* edit);
    void* (*get_selected_object)(int index);
    int (*get_selected_object_num)();
    bool (*get_mouse_layer_frame)(int* layer, int* frame);
    bool (*pos_to_layer_frame)(int x, int y, int* layer, int* frame);
    bool (*is_support_media_file)(LPCWSTR file, bool strict);
    bool (*get_media_info)(LPCWSTR file, void* info, int info_size);
    void* (*create_object_from_media_file)(LPCWSTR file, int layer, int frame, int length);
    void* (*create_object)(LPCWSTR effect, int layer, int frame, int length);
    void (*set_cursor_layer_frame)(int layer, int frame);
    void (*set_display_layer_frame)(int layer, int frame);
    void (*set_select_range)(int start, int end);
};

struct EDIT_HANDLE_SAFE {
    bool (*call_edit_section)(void (*func_proc_edit)(EDIT_SECTION_OUT_SAFE* edit));
    bool (*call_edit_section_param)(void* param, void (*func_proc_edit)(void* param, EDIT_SECTION_OUT_SAFE* edit));
};

struct HOST_APP_TABLE {
    void (*set_plugin_information)(LPCWSTR information);
    void (*f1)(void*); void (*f2)(void*); void (*f3)(void*); void (*f4)(void*);
    void (*register_import_menu)(LPCWSTR, void (*)(void*));
    void (*register_export_menu)(LPCWSTR, void (*)(void*));
    void (*register_window_client)(LPCWSTR name, HWND hwnd);
    void* (*create_edit_handle)();
    void (*f9)(void (*)(void*)); void (*f10)(void (*)(void*));
    void (*register_layer_menu)(LPCWSTR name, void (*func)(EDIT_SECTION_SAFE*));
    void (*register_object_menu)(LPCWSTR name, void (*func)(EDIT_SECTION_SAFE*));
    void (*register_config_menu)(LPCWSTR, void (*)(HWND, HINSTANCE));
    void (*register_edit_menu)(LPCWSTR, void (*)(EDIT_SECTION_SAFE*));
    void (*f15)(void (*)(EDIT_SECTION_SAFE*)); void (*f16)(void (*)(EDIT_SECTION_SAFE*));
};

// --- グローバル変数 ---
static HOST_APP_TABLE* g_host = nullptr;
static EDIT_HANDLE_SAFE* g_editHandle = nullptr;
static bool g_cachedOutputRangeValid = false;
static int g_cachedOutputRangeStart = 0;
static int g_cachedOutputRangeEnd = 0;
static HINSTANCE g_hInst = nullptr;
static HWND g_hwnd = nullptr;
static HWND g_hDlg = nullptr;
static HWND g_hSettingDlg = nullptr;
static HWND g_hSnipWnd = nullptr;
static HWND g_hInfoWnd = nullptr;
static HWND g_hCombo = nullptr;
static HWND g_hSearch = nullptr;
static HWND g_hTooltip = nullptr; 

static ULONG_PTR g_gdiplusToken;

static HFONT g_hFontUI = nullptr, g_hFontList = nullptr, g_hFontListSub = nullptr, g_hFontType = nullptr;
static HBRUSH g_hBrInputBg = nullptr, g_hBrBg = nullptr;

static std::wstring g_baseDir;
static std::vector<Asset> g_assets; 
static std::vector<Asset*> g_displayAssets;
static std::vector<std::wstring> g_categories; 
static std::wstring g_currentCategory = L"ALL", g_searchQuery = L""; 
static std::set<std::wstring> g_favPaths; 
static std::set<std::wstring> g_fixedPaths; 
static bool g_sortFavFirst = false;
static int g_gifSpeedPercent = 100;
static bool g_showGifExportGuide = true;
static bool g_gifExportKeepOriginal = false;

static int g_winX = 100, g_winY = 100, g_winW = 360, g_winH = 550;
static int g_dlgX = -1, g_dlgY = -1;

static int g_scrollY = 0, g_selectedIndex = -1, g_contextTargetIndex = -1, g_hoverIndex = -1;
static std::string g_tempAliasData; 
static std::wstring g_editOrgPath, g_editName, g_editCat, g_tempImgPath, g_msgTitle, g_msgText;
static bool g_isDragCheck = false, g_isImageRemoved = false, g_isSnipping = false;
static POINT g_dragStartPt = {0}, g_snipStart = {0}, g_snipEnd = {0};
static int g_msgResult = 0;
static int g_guideResult = IDCANCEL;
static bool g_guideHideNext = false;
static bool g_isMouseTracking = false;

static Image* g_pAddPreviewImage = nullptr;
static IStream* g_pAddPreviewStream = nullptr; 
static UINT g_addFrameCount = 0;
static UINT* g_addFrameDelays = nullptr;
static int g_addCurrentFrame = 0;
static PropertyItem* g_pAddPropertyItem = nullptr;

static std::wstring g_lastTempPath = L"";

static int g_tooltipTargetIndex = -1; 
static bool g_tooltipShown = false;   
static std::wstring g_tooltipTextMain;
static std::wstring g_tooltipTextSub;
static bool g_addGifHintTimerStarted = false;
static bool g_addDlgMouseTracking = false;
static ULONGLONG g_addGifHoverStartTick = 0;

// ============================================================
// 前方宣言
// ============================================================
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
static LRESULT CALLBACK AddDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
static int GetObjectLayerIndex(void* obj);
static int GetObjectGroupIndex(void* obj);
void OpenAddDialog(const std::string& data, bool isEdit);
int ShowDarkMsg(HWND parent, LPCWSTR text, LPCWSTR title, UINT type);
void RefreshAssets(bool reloadFav);

// ============================================================
// 1. ユーティリティ & 設定ファイル処理
// ============================================================
static std::wstring ToLower(const std::wstring& s) {
    std::wstring ret = s;
    std::transform(ret.begin(), ret.end(), ret.begin(), ::towlower);
    return ret;
}

static int GetMaxScrollY(HWND hwnd) {
    RECT rc = {0};
    GetClientRect(hwnd, &rc);
    int clientW = (int)(rc.right - rc.left);
    int clientH = (int)(rc.bottom - rc.top);
    int viewH = clientH - TITLE_H - FOOTER_H;
    if (viewH <= 0) return 0;

    int cols = (std::max)(1, clientW / MIN_ITEM_WIDTH);
    int rows = ((int)g_displayAssets.size() + cols - 1) / cols;
    int contentH = rows * ITEM_HEIGHT;
    return (std::max)(0, contentH - viewH);
}

static void UpdateScrollBar(HWND hwnd) {
    if (!hwnd) return;

    RECT rc = {0};
    GetClientRect(hwnd, &rc);
    int clientH = (int)(rc.bottom - rc.top);
    int viewH = clientH - TITLE_H - FOOTER_H;
    int maxScroll = GetMaxScrollY(hwnd);
    g_scrollY = (std::max)(0, (std::min)(g_scrollY, maxScroll));

    SCROLLINFO si = {0};
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = (maxScroll > 0) ? (maxScroll + (std::max)(1, viewH) - 1) : 0;
    si.nPage = (UINT)(viewH > 0 ? viewH : 0);
    si.nPos = g_scrollY;
    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

static void SetScrollY(HWND hwnd, int newY) {
    int maxScroll = GetMaxScrollY(hwnd);
    int clamped = (std::max)(0, (std::min)(newY, maxScroll));
    if (clamped == g_scrollY) return;

    g_scrollY = clamped;
    SetScrollPos(hwnd, SB_VERT, g_scrollY, TRUE);
    InvalidateRect(hwnd, nullptr, FALSE);
}

static std::string ReadFileContent(const std::wstring& path) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return "";
    DWORD fileSize = GetFileSize(hFile, NULL);
    if (fileSize == 0 || fileSize == INVALID_FILE_SIZE) { CloseHandle(hFile); return ""; }
    std::vector<char> buffer(fileSize); DWORD bytesRead = 0;
    ReadFile(hFile, buffer.data(), fileSize, &bytesRead, NULL);
    CloseHandle(hFile);
    return std::string(buffer.begin(), buffer.begin() + bytesRead);
}

static std::string ReadFileHead(const std::wstring& path, int size) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return "";
    std::vector<char> buffer(size + 1); DWORD bytesRead = 0;
    ReadFile(hFile, buffer.data(), size, &bytesRead, NULL);
    CloseHandle(hFile);
    buffer[bytesRead] = '\0';
    return std::string(buffer.data());
}

static void WriteFileContent(const std::wstring& path, const std::string& content) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    DWORD bytesWritten = 0;
    WriteFile(hFile, content.c_str(), (DWORD)content.size(), &bytesWritten, NULL);
    CloseHandle(hFile);
}

static void RemoveLinesStartingWith(std::string& str, const std::string& prefix) {
    std::stringstream ss(str); std::string line, output;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (line.find(prefix) == 0) continue;
        output += line + "\r\n";
    }
    str = output;
}

static void ReplaceStringAll(std::string& s, const std::string& from, const std::string& to) {
    if (from.empty()) return;
    size_t start_pos = 0;
    while ((start_pos = s.find(from, start_pos)) != std::string::npos) {
        s.replace(start_pos, from.length(), to);
        start_pos += to.length();
    }
}

static int ExtractHeaderIntValue(const std::string& src, const std::string& key) {
    std::stringstream ss(src);
    std::string line;
    bool inHeader = false;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();

        if (!inHeader) {
            if (line == "[Object]" || line == "[0]") inHeader = true;
            continue;
        }

        if (!line.empty() && line[0] == '[') break;
        if (line.rfind(key, 0) == 0) {
            try { return std::stoi(line.substr(key.length())); }
            catch (...) { return 0; }
        }
    }
    return 0;
}

// パターンA修復
static std::string RepairLegacyObjectContent(const std::string& src) {
    std::vector<size_t> headers;
    size_t pos = 0;
    while ((pos = src.find("[Object]", pos)) != std::string::npos) {
        char next = (pos + 8 < src.length()) ? src[pos + 8] : 0;
        if (next == '\r' || next == '\n') headers.push_back(pos);
        pos += 8;
    }
    if (headers.empty()) return src;
    
    std::string result = "";
    if (headers[0] > 0) result += src.substr(0, headers[0]);
    
    for (size_t i = 0; i < headers.size(); i++) {
        size_t start = headers[i];
        size_t end = (i + 1 < headers.size()) ? headers[i+1] : src.length();
        std::string block = src.substr(start, end - start);
        
        block.replace(0, 8, "[" + std::to_string(i) + "]");
        
        size_t eol = block.find_first_of("\r\n");
        if (eol != std::string::npos) {
            size_t nextLine = eol;
            while(nextLine < block.length() && (block[nextLine] == '\r' || block[nextLine] == '\n')) nextLine++;
            block.insert(nextLine, "layer=" + std::to_string(i + 1) + "\r\n");
        }
        
        std::string oldSub = "[Object.";
        std::string newSub = "[" + std::to_string(i) + ".";
        ReplaceStringAll(block, oldSub, newSub);
        
        result += block;
    }
    return result;
}

// パターンB最適化
static std::string NormalizeSingleObjectContent(const std::string& src) {
    std::string s = src;
    ReplaceStringAll(s, "[0]", "[Object]");
    ReplaceStringAll(s, "[0.", "[Object.");
    RemoveLinesStartingWith(s, "layer=");
    return s;
}

static int GetObjectLayerIndex(void* obj) {
    if (!obj) return 0;
    return *(int*)((char*)obj + 0x60); 
}

static int GetObjectGroupIndex(void* obj) {
    if (!obj) return 0;
    int group44 = *(int*)((char*)obj + 0x44);
    if (group44 > 0 && group44 < 100000000) return group44;

    int group64 = *(int*)((char*)obj + 0x64);
    if (group64 > 0 && group64 < 100000000) return group64;

    return 0;
}

static std::wstring GetFavFilePath() { 
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH); 
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetFavorites.txt"); 
    return path; 
}
static std::wstring GetFixedFilePath() { 
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH); 
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetFixed.txt"); 
    return path; 
}
static std::wstring GetConfigPath() { 
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH); 
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetConfig.ini"); 
    return path; 
}
static void InitBaseDir() {
    if (!g_baseDir.empty()) return;
    wchar_t path[MAX_PATH] = {0}; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAsset");
    g_baseDir = path; CreateDirectoryW(g_baseDir.c_str(), nullptr);
}

void SaveFavorites() { 
    std::string data; for(const auto& p : g_favPaths) { 
        int sz = WideCharToMultiByte(CP_UTF8, 0, p.c_str(), -1, NULL, 0, NULL, NULL);
        if(sz > 0) { std::vector<char> buf(sz); WideCharToMultiByte(CP_UTF8, 0, p.c_str(), -1, buf.data(), sz, NULL, NULL); data += buf.data(); data += "\r\n"; }
    }
    WriteFileContent(GetFavFilePath(), data);
}
void LoadFavorites() { 
    g_favPaths.clear(); std::string data = ReadFileContent(GetFavFilePath());
    std::stringstream ss(data); std::string line;
    while(std::getline(ss, line)) {
        if(!line.empty() && line.back() == '\r') line.pop_back();
        if(line.empty()) continue;
        int sz = MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, NULL, 0);
        if(sz > 0) { std::vector<wchar_t> buf(sz); MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, buf.data(), sz); g_favPaths.insert(std::wstring(buf.data())); }
    }
}
static void ToggleFavorite(const std::wstring& path) { 
    if (g_favPaths.count(path)) g_favPaths.erase(path); else g_favPaths.insert(path); 
    SaveFavorites(); 
}

void SaveFixedFrames() { 
    std::string data; for(const auto& p : g_fixedPaths) { 
        int sz = WideCharToMultiByte(CP_UTF8, 0, p.c_str(), -1, NULL, 0, NULL, NULL);
        if(sz > 0) { std::vector<char> buf(sz); WideCharToMultiByte(CP_UTF8, 0, p.c_str(), -1, buf.data(), sz, NULL, NULL); data += buf.data(); data += "\r\n"; }
    }
    WriteFileContent(GetFixedFilePath(), data);
}
void LoadFixedFrames() { 
    g_fixedPaths.clear(); std::string data = ReadFileContent(GetFixedFilePath());
    std::stringstream ss(data); std::string line;
    while(std::getline(ss, line)) {
        if(!line.empty() && line.back() == '\r') line.pop_back();
        if(line.empty()) continue;
        int sz = MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, NULL, 0);
        if(sz > 0) { std::vector<wchar_t> buf(sz); MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, buf.data(), sz); g_fixedPaths.insert(std::wstring(buf.data())); }
    } 
}
static void ToggleFixedFrame(const std::wstring& path) { 
    if (g_fixedPaths.count(path)) g_fixedPaths.erase(path); else g_fixedPaths.insert(path); 
    SaveFixedFrames(); 
}

static void SaveConfig() { 
    WritePrivateProfileStringW(L"Settings", L"GifSpeed", std::to_wstring(g_gifSpeedPercent).c_str(), GetConfigPath().c_str()); 
    WritePrivateProfileStringW(L"Settings", L"ShowGifExportGuide", g_showGifExportGuide ? L"1" : L"0", GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"GifExportKeepOriginal", g_gifExportKeepOriginal ? L"1" : L"0", GetConfigPath().c_str());
}
static void LoadConfig() { 
    g_gifSpeedPercent = GetPrivateProfileIntW(L"Settings", L"GifSpeed", 100, GetConfigPath().c_str()); 
    if (g_gifSpeedPercent < 10) g_gifSpeedPercent = 100;
    g_showGifExportGuide = (GetPrivateProfileIntW(L"Settings", L"ShowGifExportGuide", 1, GetConfigPath().c_str()) != 0);
    g_gifExportKeepOriginal = (GetPrivateProfileIntW(L"Settings", L"GifExportKeepOriginal", 0, GetConfigPath().c_str()) != 0);
}

static void SaveWindowPos(HWND hwnd, bool isDialog) {
    if (!hwnd) return;
    RECT rc; if (!GetWindowRect(hwnd, &rc)) return;
    std::wstring path = GetConfigPath();
    if (!isDialog) {
        g_winX = rc.left; g_winY = rc.top; g_winW = rc.right - rc.left; g_winH = rc.bottom - rc.top;
        WritePrivateProfileStringW(L"Window", L"X", std::to_wstring(g_winX).c_str(), path.c_str());
        WritePrivateProfileStringW(L"Window", L"Y", std::to_wstring(g_winY).c_str(), path.c_str());
        WritePrivateProfileStringW(L"Window", L"W", std::to_wstring(g_winW).c_str(), path.c_str());
        WritePrivateProfileStringW(L"Window", L"H", std::to_wstring(g_winH).c_str(), path.c_str());
    } else {
        g_dlgX = rc.left; g_dlgY = rc.top;
        WritePrivateProfileStringW(L"Dialog", L"X", std::to_wstring(g_dlgX).c_str(), path.c_str());
        WritePrivateProfileStringW(L"Dialog", L"Y", std::to_wstring(g_dlgY).c_str(), path.c_str());
    }
}

static void LoadWindowConfig() {
    std::wstring path = GetConfigPath();
    g_winX = GetPrivateProfileIntW(L"Window", L"X", 100, path.c_str());
    g_winY = GetPrivateProfileIntW(L"Window", L"Y", 100, path.c_str());
    g_winW = GetPrivateProfileIntW(L"Window", L"W", 360, path.c_str());
    g_winH = GetPrivateProfileIntW(L"Window", L"H", 550, path.c_str());
    g_dlgX = GetPrivateProfileIntW(L"Dialog", L"X", -1, path.c_str());
    g_dlgY = GetPrivateProfileIntW(L"Dialog", L"Y", -1, path.c_str());
    int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    if (g_winW < 100) g_winW = 360; if (g_winH < 100) g_winH = 550;
    if (g_winX > sw - 50) g_winX = 0; if (g_winY > sh - 50) g_winY = 0;
}

static std::wstring GetUniqueFileName(const std::wstring& dir, const std::wstring& name) {
    std::wstring baseName = name; std::wstring finalN = baseName; int counter = 2;
    while (PathFileExistsW((dir + L"\\" + finalN + L".object").c_str())) { finalN = baseName + L" (" + std::to_wstring(counter++) + L")"; }
    return finalN;
}

static std::wstring GetTempPreviewPath(const std::wstring& ext) {
    wchar_t tempPath[MAX_PATH]; GetTempPathW(MAX_PATH, tempPath);
    return std::wstring(tempPath) + L"myasset_temp" + ext;
}

static bool IsGifFile(const std::wstring& path) {
    std::wstring ext = path.substr(path.find_last_of(L"."));
    return (ToLower(ext) == L".gif");
}

static int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
    UINT num = 0, size = 0; GetImageEncodersSize(&num, &size);
    if(size == 0) return -1;
    ImageCodecInfo* pici = (ImageCodecInfo*)(malloc(size));
    GetImageEncoders(num, size, pici);
    for(UINT j = 0; j < num; ++j) {
        if(wcscmp(pici[j].MimeType, format) == 0) {
            *pClsid = pici[j].Clsid; free(pici); return j;
        }
    }
    free(pici); return -1;
}

struct OutputRangeContext {
    int frameStart;
    int frameEnd;
    bool valid;
};

static void CollectSelectedRange(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (OutputRangeContext*)param;
    if (!ctx || !edit || !edit->get_selected_object_num || !edit->get_selected_object || !edit->get_object_layer_frame) return;

    int count = edit->get_selected_object_num();
    if (count <= 0) return;

    int start = INT_MAX;
    int end = INT_MIN;
    for (int i = 0; i < count; i++) {
        void* obj = edit->get_selected_object(i);
        if (!obj) continue;
        OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(obj);
        start = (std::min)(start, lf.start);
        end = (std::max)(end, lf.end);
    }
    if (start <= end) {
        ctx->frameStart = start;
        ctx->frameEnd = end;
        ctx->valid = true;
    }
}

static void ApplyRangeFromContextProc(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (OutputRangeContext*)param;
    if (!ctx || !ctx->valid || !edit || !edit->set_select_range) return;
    edit->set_select_range(ctx->frameStart, ctx->frameEnd);
}

static bool TryParseFrameRangeFromAliasData(const std::string& alias, int& outStart, int& outEnd) {
    std::stringstream ss(alias);
    std::string line;
    bool inHeader = false;
    int start = INT_MAX;
    int end = INT_MIN;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (line.empty()) continue;

        if (line[0] == '[') {
            inHeader = false;
            if (line == "[Object]") inHeader = true;
            else if (line.size() >= 3 && line[0] == '[' && isdigit((unsigned char)line[1]) && line.find('.') == std::string::npos) inHeader = true;
            continue;
        }

        if (!inHeader) continue;
        if (line.rfind("frame=", 0) == 0) {
            std::string v = line.substr(6);
            size_t comma = v.find(',');
            if (comma == std::string::npos) continue;
            try {
                int s = std::stoi(v.substr(0, comma));
                int e = std::stoi(v.substr(comma + 1));
                start = (std::min)(start, s);
                end = (std::max)(end, e);
            } catch (...) {
            }
        }
    }
    if (start <= end) {
        outStart = start;
        outEnd = end;
        return true;
    }
    return false;
}

static bool CacheRangeFromCurrentAliasData() {
    g_cachedOutputRangeValid = false;
    int s = 0, e = 0;
    if (!TryParseFrameRangeFromAliasData(g_tempAliasData, s, e)) return false;
    g_cachedOutputRangeStart = s;
    g_cachedOutputRangeEnd = e;
    g_cachedOutputRangeValid = true;
    return true;
}

static bool ApplyCachedRangeToTimeline() {
    if (!g_cachedOutputRangeValid) return false;
    if (!g_editHandle || !g_editHandle->call_edit_section_param) return false;
    OutputRangeContext ctx = { g_cachedOutputRangeStart, g_cachedOutputRangeEnd, true };
    return g_editHandle->call_edit_section_param(&ctx, ApplyRangeFromContextProc) ? true : false;
}

static std::wstring GetForcedPreviewOutputPath() {
    InitBaseDir();
    return g_baseDir + L"\\_preview.gif";
}

static bool BuildBitmapFromOutputFrame(OUTPUT_INFO* oip, int frame, Bitmap* bmp) {
    if (!oip || !bmp || !oip->func_get_video) return false;
    auto* src = (const unsigned char*)oip->func_get_video(frame, BI_RGB);
    if (!src) return false;

    Rect rect(0, 0, oip->w, oip->h);
    BitmapData bits = {};
    if (bmp->LockBits(&rect, ImageLockModeWrite, PixelFormat24bppRGB, &bits) != Ok) return false;

    for (int y = 0; y < oip->h; ++y) {
        const unsigned char* srcLine = src + (size_t)(oip->h - 1 - y) * (size_t)oip->w * 3;
        unsigned char* dstLine = (unsigned char*)bits.Scan0 + (size_t)y * (size_t)bits.Stride;
        memcpy(dstLine, srcLine, (size_t)oip->w * 3);
    }
    bmp->UnlockBits(&bits);
    return true;
}

static void GetPreviewOutputSize(int srcW, int srcH, int& outW, int& outH) {
    outW = srcW;
    outH = srcH;
    if (g_gifExportKeepOriginal) return;
    if (srcW <= 480) return;
    outW = 480;
    outH = (int)(((long long)srcH * 480LL + (srcW / 2)) / srcW);
    if (outH < 1) outH = 1;
}

static Bitmap* CreateResizedFrameForPreview(Bitmap* src, int dstW, int dstH) {
    if (!src) return nullptr;
    int srcW = (int)src->GetWidth();
    int srcH = (int)src->GetHeight();
    if (srcW == dstW && srcH == dstH) return src->Clone(0, 0, srcW, srcH, PixelFormat24bppRGB);

    Bitmap* dst = new Bitmap(dstW, dstH, PixelFormat24bppRGB);
    if (!dst || dst->GetLastStatus() != Ok) {
        delete dst;
        return nullptr;
    }

    Graphics g(dst);
    g.SetInterpolationMode(InterpolationModeNearestNeighbor);
    g.SetPixelOffsetMode(PixelOffsetModeHighSpeed);
    g.SetSmoothingMode(SmoothingModeNone);
    g.SetCompositingQuality(CompositingQualityHighSpeed);
    g.SetCompositingMode(CompositingModeSourceCopy);
    g.DrawImage(src, Rect(0, 0, dstW, dstH), 0, 0, srcW, srcH, UnitPixel);
    return dst;
}

static bool SaveAnimatedGifToPath(const std::wstring& filePath, const std::vector<Bitmap*>& frames, int delayCentisec) {
    if (frames.empty()) return false;

    CLSID gifClsid = {};
    if (GetEncoderClsid(L"image/gif", &gifClsid) < 0) return false;

    Bitmap* first = frames[0];
    int frameCount = (int)frames.size();

    std::vector<unsigned char> delayBuf((size_t)frameCount * sizeof(unsigned int));
    auto* delayValues = (unsigned int*)delayBuf.data();
    for (int i = 0; i < frameCount; ++i) delayValues[i] = (unsigned int)(std::max)(1, delayCentisec);

    std::vector<unsigned char> delayPropBuf(sizeof(PropertyItem) + delayBuf.size());
    auto* delayProp = (PropertyItem*)delayPropBuf.data();
    delayProp->id = PropertyTagFrameDelay;
    delayProp->length = (ULONG)delayBuf.size();
    delayProp->type = PropertyTagTypeLong;
    delayProp->value = delayPropBuf.data() + sizeof(PropertyItem);
    memcpy(delayProp->value, delayBuf.data(), delayBuf.size());
    first->SetPropertyItem(delayProp);

    unsigned short loopCount = 0;
    std::vector<unsigned char> loopPropBuf(sizeof(PropertyItem) + sizeof(loopCount));
    auto* loopProp = (PropertyItem*)loopPropBuf.data();
    loopProp->id = PropertyTagLoopCount;
    loopProp->length = sizeof(loopCount);
    loopProp->type = PropertyTagTypeShort;
    loopProp->value = loopPropBuf.data() + sizeof(PropertyItem);
    memcpy(loopProp->value, &loopCount, sizeof(loopCount));
    first->SetPropertyItem(loopProp);

    EncoderParameters ep = {};
    ep.Count = 1;
    ep.Parameter[0].Guid = EncoderSaveFlag;
    ep.Parameter[0].Type = EncoderParameterValueTypeLong;
    ep.Parameter[0].NumberOfValues = 1;

    ULONG saveFlag = EncoderValueMultiFrame;
    ep.Parameter[0].Value = &saveFlag;
    if (first->Save(filePath.c_str(), &gifClsid, &ep) != Ok) return false;

    for (size_t i = 1; i < frames.size(); ++i) {
        saveFlag = EncoderValueFrameDimensionTime;
        ep.Parameter[0].Value = &saveFlag;
        if (first->SaveAdd(frames[i], &ep) != Ok) return false;
    }

    saveFlag = EncoderValueFlush;
    ep.Parameter[0].Value = &saveFlag;
    return first->SaveAdd(&ep) == Ok;
}

static bool OutputSelectedRangeGif(OUTPUT_INFO* oip) {
    if (!oip || !oip->func_get_video || oip->w <= 0 || oip->h <= 0 || oip->n <= 0) return false;

    int start = 0;
    int end = oip->n - 1;
    if (g_cachedOutputRangeValid) {
        int s = g_cachedOutputRangeStart;
        int e = g_cachedOutputRangeEnd;
        if (s >= 0 && e >= s && e < oip->n) {
            start = s;
            end = e;
        }
    }
    g_cachedOutputRangeValid = false;
    if (start > end) return false;

    int outW = oip->w;
    int outH = oip->h;
    GetPreviewOutputSize(oip->w, oip->h, outW, outH);

    std::vector<Bitmap*> frames;
    frames.reserve((size_t)(end - start + 1));

    bool ok = true;
    for (int f = start; f <= end; ++f) {
        if (oip->func_is_abort && oip->func_is_abort()) { ok = false; break; }
        Bitmap* srcBmp = new Bitmap(oip->w, oip->h, PixelFormat24bppRGB);
        if (!BuildBitmapFromOutputFrame(oip, f, srcBmp)) {
            delete srcBmp;
            ok = false;
            break;
        }

        Bitmap* outBmp = CreateResizedFrameForPreview(srcBmp, outW, outH);
        delete srcBmp;
        if (!outBmp) {
            ok = false;
            break;
        }

        frames.push_back(outBmp);
        if (oip->func_rest_time_disp) oip->func_rest_time_disp(f - start + 1, end - start + 1);
    }

    if (ok) {
        int delay = 3;
        if (oip->rate > 0 && oip->scale > 0) {
            double cs = (100.0 * (double)oip->scale) / (double)oip->rate;
            delay = (std::max)(1, (int)(cs + 0.5));
        }
        ok = SaveAnimatedGifToPath(GetForcedPreviewOutputPath(), frames, delay);
    }

    for (auto* bmp : frames) delete bmp;
    return ok;
}

static OUTPUT_PLUGIN_TABLE g_outputPluginTable = {
    OUTPUT_PLUGIN_TABLE::FLAG_VIDEO,
    L"MyAsset Preview GIF",
    L"GIF File (*.gif)\0*.gif\0\0",
    L"選択範囲をGIFとして出力（保存先はMyAssetに固定）",
    OutputSelectedRangeGif,
    nullptr,
    nullptr
};

static void CaptureRect(const RECT& rc, const std::wstring& savePath) {
    int w = rc.right - rc.left, h = rc.bottom - rc.top;
    if (w <= 0 || h <= 0) return;
    HDC hdcS = GetDC(NULL), hdcM = CreateCompatibleDC(hdcS);
    HBITMAP hB = CreateCompatibleBitmap(hdcS, w, h); SelectObject(hdcM, hB);
    BitBlt(hdcM, 0, 0, w, h, hdcS, rc.left, rc.top, SRCCOPY);
    Bitmap* bmp = new Bitmap(hB, NULL); CLSID clsid;
    GetEncoderClsid(L"image/png", &clsid); bmp->Save(savePath.c_str(), &clsid, NULL);
    delete bmp; DeleteObject(hB); DeleteDC(hdcM); ReleaseDC(NULL, hdcS);
}

static bool GenerateThumbnail(const std::wstring& srcPath, const std::wstring& destPath) {
    IShellItemImageFactory* pF = nullptr;
    if (SUCCEEDED(SHCreateItemFromParsingName(srcPath.c_str(), NULL, IID_PPV_ARGS(&pF)))) {
        HBITMAP hB = NULL; SIZE sz = { 256, 256 };
        if (SUCCEEDED(pF->GetImage(sz, SIIGBF_RESIZETOFIT, &hB)) && hB) {
            Bitmap* bmp = new Bitmap(hB, NULL); CLSID clsid;
            GetEncoderClsid(L"image/png", &clsid); bmp->Save(destPath.c_str(), &clsid, NULL);
            delete bmp; DeleteObject(hB); pF->Release(); return true;
        }
        pF->Release();
    }
    return false;
}

// 起動時一括チェック＆最適化関数
static void CheckAndOptimizeLegacyAssets(HWND hwnd) {
    std::wstring cfg = GetConfigPath();
    int done = GetPrivateProfileIntW(L"Settings", L"LegacyCheckDone", 0, cfg.c_str());
    if (done == 1) return; // 既にチェック済みなら即リターン

    InitBaseDir();
    
    // スキャン開始 (カウント)
    int needRepair = 0;   // Pattern A: [Object]重複シングルObject
    int needOptimize = 0; // Pattern B: 旧マルチObject
    
    // 再帰探索で全ファイルをチェック
    std::function<void(std::wstring)> scanDir = [&](std::wstring d) {
        std::wstring sp = d + L"\\*";
        WIN32_FIND_DATAW fd;
        HANDLE h = FindFirstFileW(sp.c_str(), &fd);
        if (h != INVALID_HANDLE_VALUE) {
            do {
                if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                    scanDir(d + L"\\" + fd.cFileName);
                } else {
                    std::wstring name = fd.cFileName;
                    if (name.size() > 7 && name.substr(name.size()-7) == L".object") {
                        std::wstring path = d + L"\\" + name;
                        std::string content = ReadFileContent(path);
                        
                        int hCount = 0;
                        size_t pos = 0;
                        while ((pos = content.find("[Object]", pos)) != std::string::npos) {
                            char next = (pos + 8 < content.length()) ? content[pos + 8] : 0;
                            if (next == '\r' || next == '\n') hCount++;
                            pos += 8;
                        }
                        bool hasIdx = (content.find("[0]") != std::string::npos);

                        if (hCount >= 2 && !hasIdx) needRepair++;
                        else if (hasIdx && content.find("[1]") == std::string::npos) needOptimize++;
                    }
                }
            } while (FindNextFileW(h, &fd));
            FindClose(h);
        }
    };
    scanDir(g_baseDir);

    // 修正が必要なファイルがなければ、何もせずフラグを立てて終了
    if (needRepair == 0 && needOptimize == 0) {
        WritePrivateProfileStringW(L"Settings", L"LegacyCheckDone", L"1", cfg.c_str());
        return;
    }

    // ユーザーに確認
    std::wstring msg = L"【My Asset Manager】\nアセット形式の確認\n\n";
    msg += L"形式の最適化が必要なファイルが " + std::to_wstring(needRepair + needOptimize) + L" 個見つかりました。\n";
    if (needRepair > 0) msg += L"・複数オブジェクト(旧仕様)の修復: " + std::to_wstring(needRepair) + L" 個\n";
    if (needOptimize > 0) msg += L"・単体オブジェクトの最適化: " + std::to_wstring(needOptimize) + L" 個\n";
    msg += L"\n自動的に修正・最適化しますか？";

    int res = ShowDarkMsg(hwnd, msg.c_str(), L"確認", MB_YESNO);

    if (res != IDYES) {
        // いいえなら次回以降スキップ
        WritePrivateProfileStringW(L"Settings", L"LegacyCheckDone", L"1", cfg.c_str());
        return;
    }

    // 書き換え実行
    int fixCount = 0;
    std::function<void(std::wstring)> fixDir = [&](std::wstring d) {
        std::wstring sp = d + L"\\*";
        WIN32_FIND_DATAW fd;
        HANDLE h = FindFirstFileW(sp.c_str(), &fd);
        if (h != INVALID_HANDLE_VALUE) {
            do {
                if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                    fixDir(d + L"\\" + fd.cFileName);
                } else {
                    std::wstring name = fd.cFileName;
                    if (name.size() > 7 && name.substr(name.size()-7) == L".object") {
                        std::wstring path = d + L"\\" + name;
                        std::string content = ReadFileContent(path);
                        bool modified = false;

                        // Pattern A check
                        int hCount = 0;
                        size_t pos = 0;
                        while ((pos = content.find("[Object]", pos)) != std::string::npos) {
                            char next = (pos + 8 < content.length()) ? content[pos + 8] : 0;
                            if (next == '\r' || next == '\n') hCount++;
                            pos += 8;
                        }
                        bool hasIdx = (content.find("[0]") != std::string::npos);

                        if (hCount >= 2 && !hasIdx) {
                            content = RepairLegacyObjectContent(content);
                            modified = true;
                        }
                        // Pattern B check
                        else if (hasIdx && content.find("[1]") == std::string::npos) {
                            content = NormalizeSingleObjectContent(content);
                            modified = true;
                        }

                        if (modified) {
                            WriteFileContent(path, content);
                            fixCount++;
                        }
                    }
                }
            } while (FindNextFileW(h, &fd));
            FindClose(h);
        }
    };
    fixDir(g_baseDir);

    WritePrivateProfileStringW(L"Settings", L"LegacyCheckDone", L"1", cfg.c_str());
    
    std::wstring doneMsg = std::to_wstring(fixCount) + L" 個のファイルを最適化しました。";
    ShowDarkMsg(hwnd, doneMsg.c_str(), L"完了", MB_OK);
    
    RefreshAssets(false);
}

// ============================================================
// 2. ドラッグ＆ドロップ実装 (COM)
// ============================================================
class CEnumFormatEtc : public IEnumFORMATETC {
    long m_cRef; int m_nIndex;
public:
    CEnumFormatEtc() : m_cRef(1), m_nIndex(0) {}
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) { if (riid == IID_IUnknown || riid == IID_IEnumFORMATETC) { *ppv = this; AddRef(); return S_OK; } return E_NOINTERFACE; }
    STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_cRef); }
    STDMETHODIMP_(ULONG) Release() { long v = InterlockedDecrement(&m_cRef); if (!v) delete this; return v; }
    STDMETHODIMP Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched) { if (pceltFetched) *pceltFetched = 0; if (m_nIndex == 0 && celt > 0) { rgelt[0] = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL }; if (pceltFetched) *pceltFetched = 1; m_nIndex++; return S_OK; } return S_FALSE; }
    STDMETHODIMP Skip(ULONG celt) { m_nIndex += (int)celt; return S_OK; }
    STDMETHODIMP Reset() { m_nIndex = 0; return S_OK; }
    STDMETHODIMP Clone(IEnumFORMATETC** ppEnum) { *ppEnum = new CEnumFormatEtc(); return S_OK; }
};

class CDataObject : public IDataObject {
    long m_cRef; FORMATETC m_fmt; STGMEDIUM m_med;
public:
    CDataObject(FORMATETC* f, STGMEDIUM* m) : m_cRef(1), m_fmt(*f), m_med(*m) {}
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) { if (riid == IID_IUnknown || riid == IID_IDataObject) { *ppv = this; AddRef(); return S_OK; } return E_NOINTERFACE; }
    STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_cRef); }
    STDMETHODIMP_(ULONG) Release() { long v = InterlockedDecrement(&m_cRef); if (!v) { ReleaseStgMedium(&m_med); delete this; } return v; }
    STDMETHODIMP GetData(FORMATETC* pf, STGMEDIUM* pm) { if (pf->cfFormat == m_fmt.cfFormat) { pm->tymed = TYMED_HGLOBAL; pm->hGlobal = GlobalAlloc(GHND, GlobalSize(m_med.hGlobal)); void* d = GlobalLock(pm->hGlobal); void* s = GlobalLock(m_med.hGlobal); memcpy(d, s, GlobalSize(m_med.hGlobal)); GlobalUnlock(pm->hGlobal); GlobalUnlock(m_med.hGlobal); pm->pUnkForRelease = nullptr; return S_OK; } return DV_E_FORMATETC; }
    STDMETHODIMP QueryGetData(FORMATETC* pf) { return (pf->cfFormat == m_fmt.cfFormat) ? S_OK : S_FALSE; }
    STDMETHODIMP EnumFormatEtc(DWORD dwDir, IEnumFORMATETC** ppenum) { if (dwDir == DATADIR_GET) { *ppenum = new CEnumFormatEtc(); return S_OK; } return E_NOTIMPL; }
    STDMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) { return E_NOTIMPL; }
    STDMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) { return E_NOTIMPL; }
    STDMETHODIMP SetData(FORMATETC*, STGMEDIUM*, BOOL) { return E_NOTIMPL; }
    STDMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) { return E_NOTIMPL; }
    STDMETHODIMP DUnadvise(DWORD) { return E_NOTIMPL; }
    STDMETHODIMP EnumDAdvise(IEnumSTATDATA**) { return E_NOTIMPL; }
};

class CDropSource : public IDropSource {
public:
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) { if (riid == IID_IUnknown || riid == IID_IDropSource) { *ppv = this; return S_OK; } return E_NOINTERFACE; }
    STDMETHODIMP_(ULONG) AddRef() { return 1; }
    STDMETHODIMP_(ULONG) Release() { return 1; }
    STDMETHODIMP QueryContinueDrag(BOOL fEsc, DWORD dwKeyState) { if (fEsc) return DRAGDROP_S_CANCEL; if (!(dwKeyState & MK_LBUTTON)) return DRAGDROP_S_DROP; return S_OK; }
    STDMETHODIMP GiveFeedback(DWORD) { return DRAGDROP_S_USEDEFAULTCURSORS; }
};

static HRESULT CreateFileDropDataObject(const std::wstring& path, IDataObject** ppObj) {
    size_t size = sizeof(DROPFILES) + (path.size() + 2) * sizeof(wchar_t);
    HGLOBAL hG = GlobalAlloc(GHND, size); if (!hG) return E_OUTOFMEMORY;
    DROPFILES* df = (DROPFILES*)GlobalLock(hG); df->pFiles = sizeof(DROPFILES); df->fWide = TRUE;
    wcscpy((wchar_t*)((BYTE*)df + sizeof(DROPFILES)), path.c_str()); GlobalUnlock(hG);
    FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM med = {0}; med.tymed = TYMED_HGLOBAL; med.hGlobal = hG; med.pUnkForRelease = NULL;
    *ppObj = new CDataObject(&fmt, &med); return S_OK;
}

// ============================================================
// 3. UI 描画補助
// ============================================================
static void DrawDarkButton(LPDRAWITEMSTRUCT dis) { 
    RECT rc = dis->rcItem; HBRUSH br = CreateSolidBrush((dis->itemState & ODS_SELECTED) ? COL_BTN_PUSH : COL_BTN_BG); 
    FillRect(dis->hDC, &rc, br); DeleteObject(br); FrameRect(dis->hDC, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH)); 
    wchar_t buf[256]; GetWindowTextW(dis->hwndItem, buf, 256); SetBkMode(dis->hDC, TRANSPARENT); SetTextColor(dis->hDC, COL_TEXT); SelectObject(dis->hDC, g_hFontUI); 
    DrawTextW(dis->hDC, buf, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE); 
}

static void DrawTitleBar(HDC hdc, int w, LPCWSTR title) { 
    RECT rt = {0, 0, w, TITLE_H}; HBRUSH br = CreateSolidBrush(COL_TITLE_BG); FillRect(hdc, &rt, br); DeleteObject(br); 
    SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontUI); 
    RECT rtx = {10, 0, w - 40, TITLE_H}; DrawTextW(hdc, title, -1, &rtx, DT_LEFT | DT_VCENTER | DT_SINGLELINE); 
    RECT rc = {w - 40, 0, w, TITLE_H}; DrawTextW(hdc, L"✕", -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE); 
}

static void DrawWindowBorder(HWND hwnd) { 
    HDC hdc = GetWindowDC(hwnd); RECT rc; GetWindowRect(hwnd, &rc); int w = rc.right - rc.left, h = rc.bottom - rc.top; 
    RECT rb = {0, 0, w, h}; HBRUSH br = CreateSolidBrush(COL_BORDER); FrameRect(hdc, &rb, br); DeleteObject(br); ReleaseDC(hwnd, hdc); 
}

static void OpenAssetFolderInExplorer(HWND hwnd) {
    InitBaseDir();
    if (!PathFileExistsW(g_baseDir.c_str())) {
        CreateDirectoryW(g_baseDir.c_str(), nullptr);
    }
    HINSTANCE h = ShellExecuteW(hwnd, L"open", g_baseDir.c_str(), NULL, NULL, SW_SHOWNORMAL);
    if ((INT_PTR)h <= 32) {
        ShowDarkMsg(hwnd, L"MyAsset フォルダを開けませんでした。", L"Error", MB_OK);
    }
}

static void ShowSettingsRefreshMenu(HWND hwnd, int xScreen, int yScreen) {
    HMENU hm = CreatePopupMenu();
    AppendMenuW(hm, MF_STRING, IDM_SETTINGS, L"設定...");
    AppendMenuW(hm, MF_STRING, ID_BTN_REFRESH, L"更新");
    AppendMenuW(hm, MF_STRING, IDM_OPEN_FOLDER, L"MyAssetフォルダを開く");
    TrackPopupMenu(hm, TPM_RIGHTBUTTON, xScreen, yScreen, 0, hwnd, NULL);
    DestroyMenu(hm);
}

// ============================================================
// 4. カスタムダイアログ
// ============================================================
static LRESULT CALLBACK DarkMsgDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp);
    case WM_NCACTIVATE: return TRUE;
    case WM_PAINT: { PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, g_msgTitle.c_str()); RECT rcText = {20, TITLE_H + 20, rc.right - 20, rc.bottom - 50}; SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontUI); DrawTextW(hdc, g_msgText.c_str(), -1, &rcText, DT_LEFT | DT_WORDBREAK); DrawWindowBorder(hwnd); EndPaint(hwnd, &ps); return 0; }
    case WM_NCHITTEST: { 
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; ScreenToClient(hwnd, &pt); 
        RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H && pt.x > rc.right - 40) return HTCLIENT; 
        if (pt.y < TITLE_H) return HTCAPTION; 
        return HTCLIENT; 
    }
    case WM_DRAWITEM: DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_COMMAND: { int id = LOWORD(wp); if (id == ID_BTN_MSG_OK || id == ID_BTN_MSG_YES) { g_msgResult = IDYES; DestroyWindow(hwnd); } else if (id == ID_BTN_MSG_NO || id == IDCANCEL) { g_msgResult = IDNO; DestroyWindow(hwnd); } return 0; }
    case WM_LBUTTONUP: { POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; RECT rc; GetClientRect(hwnd, &rc); if (pt.y < TITLE_H && pt.x > rc.right - 40) { g_msgResult = IDNO; DestroyWindow(hwnd); } return 0; }
    } return DefWindowProc(hwnd, msg, wp, lp);
}

void RegisterCustomDialogs() { 
    WNDCLASSW wc = {0}; wc.lpfnWndProc = DarkMsgDlgProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Msg"; wc.hCursor = LoadCursor(nullptr, IDC_ARROW); wc.hbrBackground = g_hBrBg; RegisterClassW(&wc); 
}

int ShowDarkMsg(HWND parent, LPCWSTR text, LPCWSTR title, UINT type) {
    g_msgTitle = title; g_msgText = text; g_msgResult = IDNO; int w = 400, h = 220; int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN); 
    HWND hMsg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"MyAsset_Msg", title, WS_POPUP | WS_VISIBLE | WS_THICKFRAME, (sw-w)/2, (sh-h)/2, w, h, parent, NULL, g_hInst, NULL);
    g_hInfoWnd = hMsg;
    int btnY = h - 45; 
    if (type == MB_YESNO) { CreateWindowW(L"BUTTON", L"はい", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, w - 180, btnY, 80, 26, hMsg, (HMENU)ID_BTN_MSG_YES, g_hInst, NULL); CreateWindowW(L"BUTTON", L"いいえ", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, w - 90, btnY, 80, 26, hMsg, (HMENU)ID_BTN_MSG_NO, g_hInst, NULL); } 
    else { CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, w - 100, btnY, 80, 26, hMsg, (HMENU)ID_BTN_MSG_OK, g_hInst, NULL); }
    EnableWindow(parent, FALSE); 
    MSG msg; 
    while (IsWindow(hMsg)) { 
        BOOL res = GetMessage(&msg, NULL, 0, 0);
        if (res == 0 || res == -1) { DestroyWindow(hMsg); if(res == 0) PostQuitMessage(msg.wParam); break; }
        TranslateMessage(&msg); DispatchMessage(&msg); 
    }
    EnableWindow(parent, TRUE); SetForegroundWindow(parent); g_hInfoWnd = nullptr; return g_msgResult;
}

static LRESULT CALLBACK SettingsDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp); 
    case WM_NCACTIVATE: return TRUE;
    case WM_CREATE: {
        HWND h = CreateWindowW(L"STATIC", L"GIF再生速度倍率 (50% - 1000%)", WS_VISIBLE|WS_CHILD, 20, TITLE_H+20, 280, 20, hwnd, NULL, g_hInst, NULL); SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hs = CreateWindowExW(0, TRACKBAR_CLASSW, L"", WS_VISIBLE|WS_CHILD|TBS_HORZ|TBS_AUTOTICKS, 20, TITLE_H+48, 210, 30, hwnd, (HMENU)IDC_SLIDER_SPEED, g_hInst, NULL);
        SendMessage(hs, TBM_SETRANGE, TRUE, MAKELONG(50, 1000)); SendMessage(hs, TBM_SETPOS, TRUE, g_gifSpeedPercent);
        HWND he = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", (std::to_wstring(g_gifSpeedPercent)).c_str(), WS_VISIBLE|WS_CHILD|ES_NUMBER|ES_CENTER|ES_AUTOHSCROLL, 240, TITLE_H+53, 60, 22, hwnd, (HMENU)IDC_EDIT_SPEED, g_hInst, NULL); SendMessageW(he, WM_SETFONT, (WPARAM)g_hFontUI, 0);

        HWND h2 = CreateWindowW(L"STATIC", L"GIF出力解像度", WS_VISIBLE|WS_CHILD, 20, TITLE_H+92, 280, 20, hwnd, NULL, g_hInst, NULL); SendMessageW(h2, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hChk = CreateWindowW(L"BUTTON", L"元解像度で出力（OFF: 最大幅480px）", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 20, TITLE_H+116, 280, 24, hwnd, (HMENU)IDC_CHK_GIF_ORIGINAL, g_hInst, NULL);
        SendMessageW(hChk, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChk, L"", L"");
        SendMessageW(hChk, BM_SETCHECK, g_gifExportKeepOriginal ? BST_CHECKED : BST_UNCHECKED, 0);
        CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 120, 178, 80, 30, hwnd, (HMENU)ID_BTN_MSG_OK, g_hInst, NULL);
        return 0;
    }
    
    case WM_PAINT: { PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, L"設定"); DrawWindowBorder(hwnd); EndPaint(hwnd, &ps); return 0; }
    case WM_NCHITTEST: { 
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; ScreenToClient(hwnd, &pt); 
        RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H) { 
            if (pt.x > rc.right - 40) return HTCLIENT; 
            return HTCAPTION; 
        } 
        return HTCLIENT; 
    }
    case WM_DRAWITEM: DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_HSCROLL: { 
        g_gifSpeedPercent = (int)SendMessage((HWND)lp, TBM_GETPOS, 0, 0); 
        SetDlgItemInt(hwnd, IDC_EDIT_SPEED, g_gifSpeedPercent, FALSE); 
        SaveConfig(); return 0; 
    }
    case WM_COMMAND: {
        if (LOWORD(wp) == IDC_EDIT_SPEED && HIWORD(wp) == EN_CHANGE) {
            if (GetFocus() == GetDlgItem(hwnd, IDC_EDIT_SPEED)) {
                int val = GetDlgItemInt(hwnd, IDC_EDIT_SPEED, NULL, FALSE);
                if (val < 50) val = 50; if (val > 1000) val = 1000;
                g_gifSpeedPercent = val;
                SendMessage(GetDlgItem(hwnd, IDC_SLIDER_SPEED), TBM_SETPOS, TRUE, val);
                SaveConfig();
            }
        }
        else if (LOWORD(wp) == IDC_CHK_GIF_ORIGINAL && HIWORD(wp) == BN_CLICKED) {
            g_gifExportKeepOriginal = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_GIF_ORIGINAL), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
        }
        else if (LOWORD(wp) == ID_BTN_MSG_OK) DestroyWindow(hwnd); 
        return 0;
    }
    case WM_LBUTTONUP: { POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; RECT rc; GetClientRect(hwnd, &rc); if (pt.y < TITLE_H && pt.x > rc.right - 40) { DestroyWindow(hwnd); } return 0; }
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkMode(hdc, TRANSPARENT); return (LRESULT)g_hBrBg; }
    case WM_DESTROY: g_hSettingDlg = nullptr; return 0;
    } return DefWindowProc(hwnd, msg, wp, lp);
}
void OpenSettings() { if (g_hSettingDlg) { SetForegroundWindow(g_hSettingDlg); return; } WNDCLASSW wc = {0}; wc.lpfnWndProc = SettingsDlgProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Settings"; wc.hbrBackground = g_hBrBg; RegisterClassW(&wc); int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN); g_hSettingDlg = CreateWindowExW(WS_EX_TOPMOST|WS_EX_TOOLWINDOW, L"MyAsset_Settings", L"設定", WS_POPUP|WS_VISIBLE|WS_CLIPCHILDREN|WS_THICKFRAME, (sw-320)/2, (sh-240)/2, 320, 240, g_hwnd, nullptr, g_hInst, nullptr); }

// ============================================================
// 5. アセット管理
// ============================================================
void ClearAssets() { for (auto& a : g_assets) { if (a.pImage) delete a.pImage; if (a.frameDelays) delete[] a.frameDelays; if (a.pPropertyItem) free(a.pPropertyItem); } g_assets.clear(); g_displayAssets.clear(); }

void ScanDirectory(const std::wstring& dir, const std::wstring& category) {
    std::wstring sp = dir + L"\\*"; WIN32_FIND_DATAW fd; HANDLE hF = FindFirstFileW(sp.c_str(), &fd);
    if (hF != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
            
            if (wcsncmp(fd.cFileName, L"_auto_", 6) == 0) continue;

            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) { ScanDirectory(dir + L"\\" + fd.cFileName, fd.cFileName); }
            else {
                std::wstring fn = fd.cFileName; size_t ep = fn.find_last_of(L".");
                if (ep != std::wstring::npos && fn.substr(ep) == L".object") {
                    Asset a; a.name = fn.substr(0, ep); a.path = dir + L"\\" + fn; a.category = category; 
                    a.isFavorite = (g_favPaths.count(a.path) > 0);
                    a.isFixedFrame = (g_fixedPaths.count(a.path) > 0);
                    
                    std::string head = ReadFileHead(a.path, 100);
                    a.isMulti = (head.find("[0]") != std::string::npos);

                    a.pImage = nullptr; a.frameCount = 0; a.frameDelays = nullptr; a.currentFrame = 0; a.pPropertyItem = nullptr;
                    std::wstring p1 = dir + L"\\" + a.name + L".png", p2 = dir + L"\\" + a.name + L".gif", lp = L"";
                    if (PathFileExistsW(p2.c_str())) lp = p2; else if (PathFileExistsW(p1.c_str())) lp = p1;
                    if (!lp.empty()) { a.imagePath = lp; a.pImage = Bitmap::FromFile(lp.c_str());
                        if (a.pImage && a.pImage->GetLastStatus() == Ok) {
                            UINT dc = a.pImage->GetFrameDimensionsCount();
                            if (dc > 0) { GUID* dg = new GUID[dc]; a.pImage->GetFrameDimensionsList(dg, dc); GUID gt = FrameDimensionTime;
                                for (UINT i = 0; i < dc; ++i) { if (IsEqualGUID(dg[i], gt)) { a.frameCount = a.pImage->GetFrameCount(&dg[i]); break; } }
                                if (a.frameCount == 0) a.frameCount = a.pImage->GetFrameCount(&dg[0]); delete[] dg;
                                if (a.frameCount > 1) { int ps = a.pImage->GetPropertyItemSize(PropertyTagFrameDelay);
                                    if (ps > 0) { a.pPropertyItem = (PropertyItem*)malloc(ps); a.pImage->GetPropertyItem(PropertyTagFrameDelay, ps, a.pPropertyItem);
                                        a.frameDelays = new UINT[a.frameCount]; for(UINT i=0; i<a.frameCount; i++) { long d = ((long*)a.pPropertyItem->value)[i]; a.frameDelays[i] = (d < 5) ? 100 : d * 10; } } } } } }
                    g_assets.push_back(a);
                }
            }
        } while (FindNextFileW(hF, &fd)); FindClose(hF);
    }
}
void UpdateDisplayList() { 
    g_displayAssets.clear(); 
    g_selectedIndex = -1; g_hoverIndex = -1;
    std::wstring sl = ToLower(g_searchQuery); 
    for (auto& a : g_assets) { 
        if ((g_currentCategory == L"ALL" || a.category == g_currentCategory) && (sl.empty() || ToLower(a.name).find(sl) != std::wstring::npos)) 
            g_displayAssets.push_back(&a); 
    } 
    std::sort(g_displayAssets.begin(), g_displayAssets.end(), [](Asset* a, Asset* b) { 
        if (g_sortFavFirst && a->isFavorite != b->isFavorite) return a->isFavorite > b->isFavorite; 
        return a->name < b->name; 
    }); 
    if (g_hwnd) {
        UpdateScrollBar(g_hwnd);
        InvalidateRect(g_hwnd, nullptr, FALSE);
    }
}
void RefreshAssets(bool reloadFav) { 
    InitBaseDir(); if (reloadFav) LoadFavorites(); 
    LoadFixedFrames(); 
    LoadConfig(); 
    if (!g_lastTempPath.empty()) { DeleteFileW(g_lastTempPath.c_str()); g_lastTempPath = L""; }
    ClearAssets(); ScanDirectory(g_baseDir, L"Main"); 
    g_categories.clear(); std::set<std::wstring> cs; for(auto& a : g_assets) cs.insert(a.category);
    if(g_hCombo) { SendMessage(g_hCombo, CB_RESETCONTENT, 0, 0); SendMessage(g_hCombo, CB_ADDSTRING, 0, (LPARAM)L"すべて (ALL)"); g_categories.push_back(L"ALL"); for(auto& c : cs) { SendMessage(g_hCombo, CB_ADDSTRING, 0, (LPARAM)c.c_str()); g_categories.push_back(c); } SendMessage(g_hCombo, CB_SETCURSEL, 0, 0); }
    g_currentCategory = L"ALL"; UpdateDisplayList(); 
}

// ============================================================
// 6. スニッピング & 登録ダイアログ & ツールチップ
// ============================================================

static LRESULT CALLBACK TooltipWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCHITTEST: return HTTRANSPARENT; 
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc; GetClientRect(hwnd, &rc);
        HBRUSH brBg = CreateSolidBrush(COL_TIP_BG); FillRect(hdc, &rc, brBg); DeleteObject(brBg);
        HBRUSH brBorder = CreateSolidBrush(COL_TIP_BORDER); FrameRect(hdc, &rc, brBorder); DeleteObject(brBorder);
        SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TIP_TEXT); SelectObject(hdc, g_hFontUI);
        RECT rcMain = {5, 5, rc.right - 5, 20};
        DrawTextW(hdc, g_tooltipTextMain.c_str(), -1, &rcMain, DT_LEFT | DT_TOP | DT_SINGLELINE);
        RECT rcSub = {5, 25, rc.right - 5, rc.bottom - 5};
        SetTextColor(hdc, RGB(180, 180, 180));
        DrawTextW(hdc, g_tooltipTextSub.c_str(), -1, &rcSub, DT_LEFT | DT_TOP | DT_WORDBREAK);
        EndPaint(hwnd, &ps); return 0;
    }
    } return DefWindowProc(hwnd, msg, wp, lp);
}

static void ShowMyTooltip(int x, int y, const std::wstring& mainText, const std::wstring& subText) {
    if (g_tooltipShown) return;
    if (!g_hTooltip) {
        WNDCLASSW wc = {0}; wc.lpfnWndProc = TooltipWndProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Tooltip";
        wc.hbrBackground = (HBRUSH)GetStockObject(NULL_BRUSH);
        RegisterClassW(&wc);
        g_hTooltip = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, L"MyAsset_Tooltip", NULL, WS_POPUP, 0, 0, 0, 0, NULL, NULL, g_hInst, NULL);
    }
    g_tooltipTextMain = mainText; g_tooltipTextSub = subText;
    int w = 220; int h = 60; 
    SetWindowPos(g_hTooltip, HWND_TOPMOST, x + 15, y + 15, w, h, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    InvalidateRect(g_hTooltip, NULL, FALSE);
    g_tooltipShown = true;
}

static void HideMyTooltip() {
    if (g_hTooltip && g_tooltipShown) { ShowWindow(g_hTooltip, SW_HIDE); g_tooltipShown = false; }
}

static Bitmap* LoadBitmapNoLock(const std::wstring& path, IStream** ppStream) {
    *ppStream = nullptr;
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return nullptr;
    DWORD fileSize = GetFileSize(hFile, NULL);
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, fileSize);
    if (!hGlobal) { CloseHandle(hFile); return nullptr; }
    void* pData = GlobalLock(hGlobal);
    DWORD readBytes; ReadFile(hFile, pData, fileSize, &readBytes, NULL);
    GlobalUnlock(hGlobal); CloseHandle(hFile);
    CreateStreamOnHGlobal(hGlobal, TRUE, ppStream);
    return Bitmap::FromStream(*ppStream);
}

void UpdateAddPreviewImage() {
    if (g_pAddPreviewImage) { delete g_pAddPreviewImage; g_pAddPreviewImage = nullptr; }
    if (g_pAddPreviewStream) { g_pAddPreviewStream->Release(); g_pAddPreviewStream = nullptr; } 
    if (g_addFrameDelays) { delete[] g_addFrameDelays; g_addFrameDelays = nullptr; }
    if (g_pAddPropertyItem) { free(g_pAddPropertyItem); g_pAddPropertyItem = nullptr; }
    g_addFrameCount = 0; g_addCurrentFrame = 0; if (g_hDlg) KillTimer(g_hDlg, ID_TIMER_ADD_PREVIEW);
    if (g_isImageRemoved) { if (g_hDlg) InvalidateRect(g_hDlg, NULL, FALSE); return; }
    std::wstring lp = L""; if (!g_tempImgPath.empty() && PathFileExistsW(g_tempImgPath.c_str())) lp = g_tempImgPath;
    else if (!g_editOrgPath.empty()) { std::wstring b = g_editOrgPath.substr(0, g_editOrgPath.find_last_of(L".")); std::wstring p1 = b + L".png", p2 = b + L".gif"; if (PathFileExistsW(p2.c_str())) lp = p2; else if (PathFileExistsW(p1.c_str())) lp = p1; }
    if (!lp.empty()) {
        g_pAddPreviewImage = LoadBitmapNoLock(lp, &g_pAddPreviewStream); 
        if (g_pAddPreviewImage && g_pAddPreviewImage->GetLastStatus() == Ok) {
            GUID gt = FrameDimensionTime; g_addFrameCount = g_pAddPreviewImage->GetFrameCount(&gt);
            if (g_addFrameCount > 1) { 
                int ps = g_pAddPreviewImage->GetPropertyItemSize(PropertyTagFrameDelay); 
                if (ps > 0) { g_pAddPropertyItem = (PropertyItem*)malloc(ps); g_pAddPreviewImage->GetPropertyItem(PropertyTagFrameDelay, ps, g_pAddPropertyItem); 
                g_addFrameDelays = new UINT[g_addFrameCount]; for(UINT i=0; i<g_addFrameCount; i++) g_addFrameDelays[i] = ((long*)g_pAddPropertyItem->value)[i] * 10; }
                if (g_hDlg) SetTimer(g_hDlg, ID_TIMER_ADD_PREVIEW, 100, NULL);
            }
        }
    }
    if (g_hDlg) InvalidateRect(g_hDlg, NULL, FALSE);
}

static LRESULT CALLBACK SnipWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_KEYDOWN: if (wp == VK_ESCAPE) { ReleaseCapture(); DestroyWindow(hwnd); if (g_hDlg) { ShowWindow(g_hDlg, SW_SHOW); SetForegroundWindow(g_hDlg); } else if (g_hwnd) { ShowWindow(g_hwnd, SW_SHOW); } } return 0;
    case WM_LBUTTONDOWN: g_isSnipping = true; g_snipStart.x = GET_X_LPARAM(lp); g_snipStart.y = GET_Y_LPARAM(lp); g_snipEnd = g_snipStart; SetCapture(hwnd); return 0;
    case WM_MOUSEMOVE: if (g_isSnipping) { g_snipEnd.x = GET_X_LPARAM(lp); g_snipEnd.y = GET_Y_LPARAM(lp); InvalidateRect(hwnd, NULL, FALSE); } return 0;
    case WM_LBUTTONUP: if (g_isSnipping) {
            g_isSnipping = false; ReleaseCapture(); ShowWindow(hwnd, SW_HIDE);
            RECT rc; rc.left = (std::min)(g_snipStart.x, g_snipEnd.x); rc.top = (std::min)(g_snipStart.y, g_snipEnd.y); rc.right = (std::max)(g_snipStart.x, g_snipEnd.x); rc.bottom = (std::max)(g_snipStart.y, g_snipEnd.y);
            POINT ptL = {rc.left, rc.top}, ptR = {rc.right, rc.bottom}; ClientToScreen(hwnd, &ptL); ClientToScreen(hwnd, &ptR); rc.left = ptL.x; rc.top = ptL.y; rc.right = ptR.x; rc.bottom = ptR.y;
            g_tempImgPath = GetTempPreviewPath(L".png"); CaptureRect(rc, g_tempImgPath); DestroyWindow(hwnd);
            g_isImageRemoved = false; UpdateAddPreviewImage(); 
            if (g_hwnd) ShowWindow(g_hwnd, SW_SHOW);
            if (g_hDlg) { ShowWindow(g_hDlg, SW_SHOW); SetForegroundWindow(g_hDlg); }
        } return 0;
    case WM_PAINT: { PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); if (g_isSnipping) { RECT rc = {(std::min)(g_snipStart.x, g_snipEnd.x), (std::min)(g_snipStart.y, g_snipEnd.y), (std::max)(g_snipStart.x, g_snipEnd.x), (std::max)(g_snipStart.y, g_snipEnd.y)}; HBRUSH br = CreateSolidBrush(RGB(255, 0, 0)); FrameRect(hdc, &rc, br); DeleteObject(br); } EndPaint(hwnd, &ps); return 0; }
    case WM_ERASEBKGND: return 1;
    } return DefWindowProc(hwnd, msg, wp, lp);
}

void StartSnipping() {
    if (g_hDlg) ShowWindow(g_hDlg, SW_HIDE); if (g_hwnd) ShowWindow(g_hwnd, SW_HIDE);
    WNDCLASSW wc = {0}; wc.lpfnWndProc = SnipWndProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Snipper"; wc.hCursor = LoadCursor(NULL, IDC_CROSS); RegisterClassW(&wc);
    int x = GetSystemMetrics(SM_XVIRTUALSCREEN), y = GetSystemMetrics(SM_YVIRTUALSCREEN), w = GetSystemMetrics(SM_CXVIRTUALSCREEN), h = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    g_hSnipWnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW, L"MyAsset_Snipper", L"", WS_POPUP | WS_VISIBLE, x, y, w, h, NULL, NULL, g_hInst, NULL);
    SetLayeredWindowAttributes(g_hSnipWnd, 0, 50, LWA_ALPHA); SetFocus(g_hSnipWnd);
}

void OpenImageFileExplorer(HWND parent) { wchar_t sz[MAX_PATH] = {0}; OPENFILENAMEW of = {0}; of.lStructSize = sizeof(of); of.hwndOwner = parent; of.lpstrFile = sz; of.nMaxFile = sizeof(sz); of.lpstrFilter = L"Image Files\0*.png;*.gif;*.jpg;*.jpeg;*.bmp\0All Files\0*.*\0"; of.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST; if (GetOpenFileNameW(&of)) { g_tempImgPath = sz; g_isImageRemoved = false; UpdateAddPreviewImage(); } }

static bool TriggerGifExportShortcut() {
    HWND hostWnd = g_hwnd ? GetAncestor(g_hwnd, GA_ROOT) : nullptr;
    if (hostWnd && IsWindow(hostWnd)) {
        ShowWindow(hostWnd, SW_RESTORE);
        SetForegroundWindow(hostWnd);
        SetActiveWindow(hostWnd);
    }
    Sleep(80);

    INPUT in[6] = {};
    in[0].type = INPUT_KEYBOARD; in[0].ki.wVk = VK_CONTROL;
    in[1].type = INPUT_KEYBOARD; in[1].ki.wVk = VK_SHIFT;
    in[2].type = INPUT_KEYBOARD; in[2].ki.wVk = 'M';
    in[3].type = INPUT_KEYBOARD; in[3].ki.wVk = 'M'; in[3].ki.dwFlags = KEYEVENTF_KEYUP;
    in[4].type = INPUT_KEYBOARD; in[4].ki.wVk = VK_SHIFT; in[4].ki.dwFlags = KEYEVENTF_KEYUP;
    in[5].type = INPUT_KEYBOARD; in[5].ki.wVk = VK_CONTROL; in[5].ki.dwFlags = KEYEVENTF_KEYUP;
    UINT sent = SendInput((UINT)(sizeof(in) / sizeof(in[0])), in, sizeof(INPUT));
    return sent == (UINT)(sizeof(in) / sizeof(in[0]));
}

static void ActivateHostForSelectionRead(HWND requester) {
    HWND hostWnd = g_hwnd ? GetAncestor(g_hwnd, GA_ROOT) : nullptr;
    if (hostWnd && IsWindow(hostWnd)) {
        ShowWindow(hostWnd, SW_RESTORE);
        SetForegroundWindow(hostWnd);
        SetActiveWindow(hostWnd);
        SetFocus(hostWnd);
    }
    Sleep(120);
}

static LRESULT CALLBACK GifGuideDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCCALCSIZE: if (wp) return 0; return DefWindowProc(hwnd, msg, wp, lp);
    case WM_NCACTIVATE: return TRUE;
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc; GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrBg);
        DrawTitleBar(hdc, rc.right, L"保存先の指定");
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, COL_TEXT);
        SelectObject(hdc, g_hFontUI);
        RECT rt = {20, TITLE_H + 20, rc.right - 20, rc.bottom - 90};
        DrawTextW(
            hdc,
            L"この後に表示される保存ウィンドウでは、保存先やファイル名は任意で問題ありません。\n"
            L"実際のGIFはプラグイン側で MyAsset フォルダへ自動保存されます。\n"
            L"そのまま「保存」を押してください。",
            -1, &rt, DT_LEFT | DT_WORDBREAK);
        DrawWindowBorder(hwnd);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_NCHITTEST: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        ScreenToClient(hwnd, &pt);
        RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H && pt.x > rc.right - 40) return HTCLIENT;
        if (pt.y < TITLE_H) return HTCAPTION;
        return HTCLIENT;
    }
    case WM_DRAWITEM:
        DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wp;
        HWND hCtrl = (HWND)lp;
        if (hCtrl && GetDlgCtrlID(hCtrl) == IDC_CHK_GUIDE_HIDE) SetTextColor(hdc, RGB(255, 255, 255));
        else SetTextColor(hdc, COL_TEXT);
        SetBkMode(hdc, TRANSPARENT);
        SetBkColor(hdc, COL_BG);
        return (LRESULT)g_hBrBg;
    }
    case WM_COMMAND: {
        int id = LOWORD(wp);
        if (id == ID_BTN_GUIDE_OK) {
            g_guideHideNext = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_GUIDE_HIDE), BM_GETCHECK, 0, 0) == BST_CHECKED);
            g_guideResult = IDOK;
            DestroyWindow(hwnd);
        } else if (id == ID_BTN_GUIDE_CANCEL || id == IDCANCEL) {
            g_guideResult = IDCANCEL;
            DestroyWindow(hwnd);
        }
        return 0;
    }
    case WM_LBUTTONUP: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H && pt.x > rc.right - 40) {
            g_guideResult = IDCANCEL;
            DestroyWindow(hwnd);
        }
        return 0;
    }
    }
    return DefWindowProc(hwnd, msg, wp, lp);
}

static bool ShowGifExportGuideDialog(HWND parent) {
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = GifGuideDlgProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"MyAsset_GifGuide";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = g_hBrBg;
    RegisterClassW(&wc);

    int w = 500, h = 270;
    int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    HWND hDlg = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"MyAsset_GifGuide",
        L"保存先の指定",
        WS_POPUP | WS_VISIBLE | WS_THICKFRAME,
        (sw - w) / 2, (sh - h) / 2, w, h,
        parent, nullptr, g_hInst, nullptr);
    if (!hDlg) return false;

    HWND hChk = CreateWindowW(L"BUTTON", L"今後表示しない", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
                              20, h - 78, 160, 24, hDlg, (HMENU)IDC_CHK_GUIDE_HIDE, g_hInst, NULL);
    CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
                  w - 190, h - 55, 80, 28, hDlg, (HMENU)ID_BTN_GUIDE_OK, g_hInst, NULL);
    CreateWindowW(L"BUTTON", L"キャンセル", WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
                  w - 100, h - 55, 80, 28, hDlg, (HMENU)ID_BTN_GUIDE_CANCEL, g_hInst, NULL);
    SendMessageW(hChk, WM_SETFONT, (WPARAM)g_hFontUI, 0);
    SetWindowTheme(hChk, L"", L"");

    g_guideResult = IDCANCEL;
    g_guideHideNext = false;
    EnableWindow(parent, FALSE);
    MSG msg;
    while (IsWindow(hDlg)) {
        BOOL res = GetMessage(&msg, NULL, 0, 0);
        if (res <= 0) break;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    EnableWindow(parent, TRUE);
    SetForegroundWindow(parent);
    return g_guideResult == IDOK;
}

static bool ShowGifExportGuideIfNeeded(HWND parent) {
    if (!g_showGifExportGuide) return true;

    bool ok = ShowGifExportGuideDialog(parent);
    if (ok && g_guideHideNext) {
        g_showGifExportGuide = false;
        SaveConfig();
    }
    return ok;
}

static LRESULT CALLBACK AddDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp); 
    case WM_NCACTIVATE: return TRUE;
    case WM_CREATE: {
        g_addGifHintTimerStarted = false;
        g_addDlgMouseTracking = false;
        g_addGifHoverStartTick = 0;
        SetTimer(hwnd, ID_TIMER_GIF_HINT_POLL, 100, NULL);
        HWND h = CreateWindowW(L"STATIC", L"名前:", WS_VISIBLE|WS_CHILD, 15, TITLE_H+15, 50, 20, hwnd, NULL, g_hInst, NULL); SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        h = CreateWindowW(L"EDIT", L"", WS_VISIBLE|WS_CHILD|WS_BORDER|ES_AUTOHSCROLL, 75, TITLE_H+12, 260, 24, hwnd, (HMENU)IDC_EDIT_NAME, g_hInst, NULL); SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hc = CreateWindowW(L"COMBOBOX", L"User", WS_VISIBLE|WS_CHILD|WS_VSCROLL|CBS_DROPDOWN, 75, TITLE_H+47, 260, 200, hwnd, (HMENU)IDC_COMBO_CAT, g_hInst, NULL); SendMessageW(hc, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hcat = CreateWindowW(L"STATIC", L"カテゴリ:", WS_VISIBLE|WS_CHILD, 15, TITLE_H+50, 60, 20, hwnd, NULL, g_hInst, NULL); SendMessageW(hcat, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        for(const auto& c : g_categories) if(c != L"ALL") SendMessage(hc, CB_ADDSTRING, 0, (LPARAM)c.c_str());
        CreateWindowW(L"BUTTON", L"キャプチャ", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, TITLE_H+80, 100, 26, hwnd, (HMENU)ID_BTN_SNIP, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"GIF生成", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 190, TITLE_H+80, 100, 26, hwnd, (HMENU)ID_BTN_GIF_EXPORT, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"保存", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, 340, 100, 30, hwnd, (HMENU)ID_BTN_SAVE, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"キャンセル", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 190, 340, 100, 30, hwnd, (HMENU)ID_BTN_CANCEL, g_hInst, NULL);
        if (!g_editOrgPath.empty()) { SetDlgItemTextW(hwnd, IDC_EDIT_NAME, g_editName.c_str()); SetDlgItemTextW(hwnd, IDC_COMBO_CAT, g_editCat.c_str()); }
        UpdateAddPreviewImage(); return 0;
    }
    case WM_TIMER:
        if (wp == ID_TIMER_ADD_PREVIEW && g_pAddPreviewImage && g_addFrameCount > 1) {
            g_addCurrentFrame = (g_addCurrentFrame + 1) % g_addFrameCount;
            RECT rp = {75, TITLE_H+120, 335, TITLE_H+266};
            InvalidateRect(hwnd, &rp, FALSE);
        } else if (wp == ID_TIMER_GIF_HINT_POLL) {
            HWND hGifBtn = GetDlgItem(hwnd, ID_BTN_GIF_EXPORT);
            if (hGifBtn && IsWindow(hGifBtn)) {
                RECT rb; GetWindowRect(hGifBtn, &rb);
                POINT cur = {0, 0}; GetCursorPos(&cur);
                bool onGifBtn = PtInRect(&rb, cur) ? true : false;
                if (onGifBtn) {
                    if (g_addGifHoverStartTick == 0) g_addGifHoverStartTick = GetTickCount64();
                    if (!g_tooltipShown && (GetTickCount64() - g_addGifHoverStartTick >= 1000)) {
                        int x = rb.left + 10;
                        int y = rb.top - 90;
                        if (y < 10) y = 10;
                        ShowMyTooltip(
                            x, y,
                            L"ショートカット設定の案内",
                            L"ファイル: MyAsset Preview GIF を Ctrl+Shift+M に設定してください");
                    }
                } else {
                    g_addGifHoverStartTick = 0;
                    HideMyTooltip();
                }
            }
        } else if (wp == ID_TIMER_GIF_HINT) {
            KillTimer(hwnd, ID_TIMER_GIF_HINT);
            g_addGifHintTimerStarted = false;
            HWND hGifBtn = GetDlgItem(hwnd, ID_BTN_GIF_EXPORT);
            if (hGifBtn && IsWindow(hGifBtn)) {
                RECT rb; GetWindowRect(hGifBtn, &rb);
                int x = rb.left + 10;
                int y = rb.top - 90;
                if (y < 10) y = 10;
                ShowMyTooltip(
                    x, y,
                    L"ショートカット設定の案内",
                    L"ファイル: MyAsset Preview GIF を Ctrl+Shift+M に設定してください");
            }
        }
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, L"アセット編集"); DrawWindowBorder(hwnd);
        Graphics graphics(hdc); RECT rp = {75, TITLE_H+120, 335, TITLE_H+266}; HBRUSH brP = CreateSolidBrush(RGB(20, 20, 20)); FillRect(hdc, &rp, brP); DeleteObject(brP);
        
        if (g_pAddPreviewImage) {
            if (g_addFrameCount > 1) { GUID gt = FrameDimensionTime; g_pAddPreviewImage->SelectActiveFrame(&gt, g_addCurrentFrame); } 
            
            
            UINT iw = g_pAddPreviewImage->GetWidth();
            UINT ih = g_pAddPreviewImage->GetHeight();
            float previewW = 260.0f;
            float previewH = 146.0f;
            
            float thumbRatio = previewW / previewH;
            float imgRatio = (float)iw / (float)ih;
            float srcX = 0, srcY = 0, srcW = (float)iw, srcH = (float)ih;

            if (imgRatio > thumbRatio) {
                srcW = ih * thumbRatio;
                srcX = (iw - srcW) / 2.0f;
            } else {
                srcH = iw / thumbRatio;
                srcY = (ih - srcH) / 2.0f;
            }

            graphics.DrawImage(g_pAddPreviewImage, 
                Rect((INT)rp.left, (INT)rp.top, (INT)previewW, (INT)previewH),
                (INT)srcX, (INT)srcY, (INT)srcW, (INT)srcH,
                UnitPixel);
            

            RECT rx = { rp.right - 24, rp.top, rp.right, rp.top + 24 }; HBRUSH brX = CreateSolidBrush(RGB(200, 50, 50)); FillRect(hdc, &rx, brX); DeleteObject(brX);
            SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, RGB(255, 255, 255)); DrawTextW(hdc, L"×", -1, &rx, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        } else { SetTextColor(hdc, COL_TEXT); DrawTextW(hdc, L"No Preview (Click to Add)", -1, &rp, DT_CENTER | DT_VCENTER | DT_SINGLELINE); }
        EndPaint(hwnd, &ps); return 0;
    }
    case WM_NCHITTEST: { POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; ScreenToClient(hwnd, &pt); RECT rc; GetClientRect(hwnd, &rc); if (pt.y < TITLE_H) { if (pt.x > rc.right - 40) return HTCLIENT; return HTCAPTION; } return HTCLIENT; }
    case WM_ACTIVATE:
        if (LOWORD(wp) != WA_INACTIVE && !g_isImageRemoved) {
            std::wstring forcedGif = GetForcedPreviewOutputPath();
            if (PathFileExistsW(forcedGif.c_str())) {
                std::wstring cur = ToLower(g_tempImgPath);
                std::wstring forced = ToLower(forcedGif);
                if (g_tempImgPath.empty() || cur == forced) {
                    g_tempImgPath = forcedGif;
                    UpdateAddPreviewImage();
                }
            }
        }
        return 0;
    case WM_LBUTTONUP: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; RECT rc; GetClientRect(hwnd, &rc); 
        if (pt.y < TITLE_H && pt.x > rc.right - 40) { DestroyWindow(hwnd); return 0; }
        RECT rp = {75, TITLE_H+120, 335, TITLE_H+266};
        if (g_pAddPreviewImage) { RECT rx = { rp.right - 24, rp.top, rp.right, rp.top + 24 }; if (PtInRect(&rx, pt)) { g_isImageRemoved = true; g_tempImgPath = L""; UpdateAddPreviewImage(); return 0; } }
        if (PtInRect(&rp, pt)) { wchar_t sz[MAX_PATH] = {0}; OPENFILENAMEW of = {0}; of.lStructSize = sizeof(of); of.hwndOwner = hwnd; of.lpstrFile = sz; of.nMaxFile = sizeof(sz); of.lpstrFilter = L"Image Files\0*.png;*.gif;*.jpg;*.jpeg;*.bmp\0All Files\0*.*\0"; of.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST; if (GetOpenFileNameW(&of)) { g_tempImgPath = sz; g_isImageRemoved = false; UpdateAddPreviewImage(); } }
        return 0;
    }
    case WM_DRAWITEM: DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_CTLCOLOREDIT: case WM_CTLCOLORLISTBOX: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkColor(hdc, COL_INPUT_BG); return (LRESULT)g_hBrInputBg; }
    case WM_CTLCOLORSTATIC: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkMode(hdc, TRANSPARENT); return (LRESULT)g_hBrBg; }
    case WM_COMMAND: {
        if (LOWORD(wp) == ID_BTN_SAVE) {
            wchar_t nb[256], cb[256]; 
            GetDlgItemTextW(hwnd, IDC_EDIT_NAME, nb, 256); 
            GetDlgItemTextW(hwnd, IDC_COMBO_CAT, cb, 256);
            
            if (wcslen(nb) == 0) { ShowDarkMsg(hwnd, L"名前を入力してください", L"確認", MB_OK); return 0; }
            
            InitBaseDir(); 
            std::wstring sd = g_baseDir + L"\\" + cb; 
            CreateDirectoryW(sd.c_str(), nullptr);
            
            std::wstring fn = g_editOrgPath.empty() ? GetUniqueFileName(sd, nb) : nb;
            std::wstring destBase = sd + L"\\" + fn;
            std::wstring sp = destBase + L".object";
            
            std::wstring oldBase = L"";
            if (!g_editOrgPath.empty()) oldBase = g_editOrgPath.substr(0, g_editOrgPath.find_last_of(L"."));

            if (!g_editOrgPath.empty()) {
                MoveFileW(g_editOrgPath.c_str(), sp.c_str());
            } else {
                WriteFileContent(sp, g_tempAliasData); 
            }
            
            // プレビューの解放
            if (g_pAddPreviewImage) { delete g_pAddPreviewImage; g_pAddPreviewImage = nullptr; }
            if (g_pAddPreviewStream) { g_pAddPreviewStream->Release(); g_pAddPreviewStream = nullptr; }

            // 旧画像の整理
            if (!oldBase.empty()) {
                bool nameChanged = (oldBase != destBase);
                bool imageUpdated = (!g_tempImgPath.empty());
                
                if (nameChanged) {
                    if (!g_isImageRemoved && !imageUpdated) {
                        MoveFileW((oldBase + L".png").c_str(), (destBase + L".png").c_str());
                        MoveFileW((oldBase + L".gif").c_str(), (destBase + L".gif").c_str());
                    } else {
                        DeleteFileW((oldBase + L".png").c_str());
                        DeleteFileW((oldBase + L".gif").c_str());
                    }
                } else {
                    if (g_isImageRemoved || imageUpdated) {
                        DeleteFileW((oldBase + L".png").c_str());
                        DeleteFileW((oldBase + L".gif").c_str());
                    }
                }
            }

            // 新しい画像の移動
            std::wstring srcImgPath = L"";
            std::wstring forcedGif = GetForcedPreviewOutputPath();
            if (PathFileExistsW(forcedGif.c_str())) srcImgPath = forcedGif;
            else if (!g_tempImgPath.empty() && PathFileExistsW(g_tempImgPath.c_str())) srcImgPath = g_tempImgPath;

            if (!srcImgPath.empty()) {
                size_t dot = srcImgPath.find_last_of(L".");
                std::wstring ext = (dot != std::wstring::npos) ? srcImgPath.substr(dot) : L".gif";
                std::wstring destImgPath = destBase + ext;
                MoveFileExW(srcImgPath.c_str(), destImgPath.c_str(), MOVEFILE_REPLACE_EXISTING);
            }

            RefreshAssets(false); 
            if(g_hwnd) { ShowWindow(g_hwnd, SW_SHOW); SetForegroundWindow(g_hwnd); InvalidateRect(g_hwnd, NULL, FALSE); }
            DestroyWindow(hwnd);
        } else if (LOWORD(wp) == ID_BTN_CANCEL) DestroyWindow(hwnd);
        else if (LOWORD(wp) == ID_BTN_SNIP) StartSnipping();
        else if (LOWORD(wp) == ID_BTN_GIF_EXPORT) {
            if (!ShowGifExportGuideIfNeeded(hwnd)) return 0;
            if (!CacheRangeFromCurrentAliasData()) {
                ShowDarkMsg(hwnd, L"書き出し範囲を取得できませんでした。\n先に「MyAsset: 追加」で対象オブジェクトを選択してください。", L"Info", MB_OK);
                return 0;
            }
            ActivateHostForSelectionRead(hwnd);
            if (!ApplyCachedRangeToTimeline()) {
                ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
                ShowDarkMsg(hwnd, L"タイムラインの範囲指定に失敗しました。", L"Error", MB_OK);
                return 0;
            }
            if (!TriggerGifExportShortcut()) {
                ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
                ShowDarkMsg(hwnd, L"ショートカット送信に失敗しました。", L"Error", MB_OK);
            }
        }
        return 0;
    }

    case WM_DESTROY: 
        SaveWindowPos(hwnd, true); 
        DeleteFileW(GetForcedPreviewOutputPath().c_str());
        KillTimer(hwnd, ID_TIMER_GIF_HINT_POLL);
        if (g_addGifHintTimerStarted) {
            KillTimer(hwnd, ID_TIMER_GIF_HINT);
            g_addGifHintTimerStarted = false;
        }
        g_addDlgMouseTracking = false;
        g_addGifHoverStartTick = 0;
        HideMyTooltip();
        if (g_pAddPreviewImage) { delete g_pAddPreviewImage; g_pAddPreviewImage = nullptr; }
        if (g_pAddPreviewStream) { g_pAddPreviewStream->Release(); g_pAddPreviewStream = nullptr; }
        g_hDlg = nullptr; return 0;
    } return DefWindowProc(hwnd, msg, wp, lp);
}

void OpenAddDialog(const std::string& data, bool isEdit) { 
    if (g_hDlg) { ShowWindow(g_hDlg, SW_SHOW); SetForegroundWindow(g_hDlg); return; } 
    g_tempAliasData = data; if (!isEdit) g_editOrgPath = L""; g_isImageRemoved = false; g_tempImgPath = L""; 
    LoadWindowConfig(); int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    int x = (g_dlgX == -1) ? (sw - 360) / 2 : g_dlgX; int y = (g_dlgY == -1) ? (sh - 400) / 2 : g_dlgY;
    WNDCLASSW wc = {0}; wc.lpfnWndProc = AddDlgProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Add"; wc.hbrBackground = g_hBrBg; RegisterClassW(&wc); 
    g_hDlg = CreateWindowExW(WS_EX_TOPMOST|WS_EX_TOOLWINDOW, L"MyAsset_Add", L"アセット", WS_POPUP|WS_VISIBLE|WS_THICKFRAME, x, y, 360, 400, g_hwnd, nullptr, g_hInst, nullptr); 
}

// ============================================================
// 8. メイン描画 & プロシージャ (WndProc)
// ============================================================

static void DrawContent(HDC hdc, int w, int h) {
    RECT rcA = {0,0,w,h}; FillRect(hdc, &rcA, g_hBrBg); 
    int yO = TITLE_H - g_scrollY; HRGN hr = CreateRectRgn(0, TITLE_H, w, h - FOOTER_H); SelectClipRgn(hdc, hr);
    Graphics g(hdc); int cols = (std::max)(1, w / MIN_ITEM_WIDTH); int cw = w / cols;
    
    SetBkMode(hdc, TRANSPARENT);

    for (int i = 0; i < (int)g_displayAssets.size(); i++) {
        int x = (i % cols) * cw, y = (i / cols) * ITEM_HEIGHT + yO;
        if (y > h - FOOTER_H || y + ITEM_HEIGHT < TITLE_H) continue;
        RECT ri = {x + 2, y + 2, x + cw - 2, y + ITEM_HEIGHT - 2}; if (!RectVisible(hdc, &ri)) continue;
        
        HBRUSH br = CreateSolidBrush((i == g_selectedIndex) ? COL_ITEM_SEL : COL_ITEM_BG); FillRect(hdc, &ri, br); DeleteObject(br);
        Asset* p = g_displayAssets[i];
        
        if (p->pImage) { 
            if (p->frameCount > 1) { 
                GUID gt = FrameDimensionTime; 
                p->pImage->SelectActiveFrame(&gt, p->currentFrame); 
            } 
            
            // 画像サイズと描画先(サムネ枠)サイズを取得
            UINT iw = p->pImage->GetWidth();
            UINT ih = p->pImage->GetHeight();
            float thumbRatio = (float)THUMB_W / (float)THUMB_H;
            float imgRatio = (float)iw / (float)ih;

            float srcX = 0, srcY = 0, srcW = (float)iw, srcH = (float)ih;

            if (imgRatio > thumbRatio) {
                srcW = ih * thumbRatio;
                srcX = (iw - srcW) / 2.0f;
            } else {
                srcH = iw / thumbRatio;
                srcY = (ih - srcH) / 2.0f;
            }

            g.DrawImage(p->pImage, 
                Rect(x + 7, y + 11, THUMB_W, THUMB_H),
                (INT)srcX, (INT)srcY, (INT)srcW, (INT)srcH,
                UnitPixel);
        }

        if (p->isFavorite) { SetTextColor(hdc, RGB(255, 215, 0)); SelectObject(hdc, g_hFontUI); RECT rs = {ri.right - 25, ri.top + 5, ri.right - 5, ri.top + 25}; DrawTextW(hdc, L"★", -1, &rs, DT_RIGHT); }
        
        g.SetSmoothingMode(SmoothingModeAntiAlias);
        if (p->isMulti) {
            RECT rm = {ri.right - 25, ri.bottom - 25, ri.right - 5, ri.bottom - 5};
            SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontType); 
            DrawTextW(hdc, L"M", -1, &rm, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        } else {
            RECT rs = {ri.right - 25, ri.bottom - 25, ri.right - 5, ri.bottom - 5};
            SetTextColor(hdc, p->isFixedFrame ? COL_FIXED_TEXT : COL_TEXT); 
            SelectObject(hdc, g_hFontType); 
            DrawTextW(hdc, L"S", -1, &rs, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
        g.SetSmoothingMode(SmoothingModeDefault);

        SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontList); RECT rn = {x + THUMB_W + 15, y + 15, ri.right - 30, y + 45}; DrawTextW(hdc, p->name.c_str(), -1, &rn, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
        SetTextColor(hdc, COL_SUBTEXT); SelectObject(hdc, g_hFontListSub); RECT rc = {x + THUMB_W + 15, y + 50, ri.right - 10, ri.bottom - 5}; DrawTextW(hdc, p->category.c_str(), -1, &rc, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);
    }
    SelectClipRgn(hdc, NULL); DeleteObject(hr); 
    DrawTitleBar(hdc, w, L"My Asset Manager");
    RECT rcF = {0, h - FOOTER_H, w, h}; HBRUSH brF = CreateSolidBrush(COL_FOOTER); FillRect(hdc, &rcF, brF); DeleteObject(brF);
    RECT rcRef = {w - 70, h - FOOTER_H + 8, w - 10, h - FOOTER_H + 32}; HBRUSH brRef = CreateSolidBrush(COL_BTN_BG); FillRect(hdc, &rcRef, brRef); DeleteObject(brRef);
    SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontUI); DrawTextW(hdc, L"更新", -1, &rcRef, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
    RECT rcFav = {w - 140, h - FOOTER_H + 8, w - 80, h - FOOTER_H + 32}; HBRUSH brFav = CreateSolidBrush(g_sortFavFirst ? COL_BTN_ACT : COL_BTN_BG); FillRect(hdc, &rcFav, brFav); DeleteObject(brFav);
    DrawTextW(hdc, L"★優先", -1, &rcFav, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
    DrawWindowBorder(g_hwnd);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        OleInitialize(NULL); GdiplusStartupInput si; GdiplusStartup(&g_gdiplusToken, &si, NULL);
        g_hFontUI = CreateFontW(15,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontList = CreateFontW(17,0,0,0,FW_BOLD,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontListSub = CreateFontW(13,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontType = CreateFontW(16,0,0,0,FW_BOLD,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Arial");

        g_hBrInputBg = CreateSolidBrush(COL_INPUT_BG); g_hBrBg = CreateSolidBrush(COL_BG);
        RegisterCustomDialogs();
        g_hCombo = CreateWindowW(L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST, 10, 0, 140, 300, hwnd, (HMENU)ID_COMBO_CATEGORY, g_hInst, NULL); SendMessageW(g_hCombo, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        g_hSearch = CreateWindowW(L"EDIT", L"", WS_CHILD|WS_VISIBLE|WS_BORDER|ES_AUTOHSCROLL, 160, 0, 100, 24, hwnd, (HMENU)ID_EDIT_SEARCH, g_hInst, NULL); SendMessageW(g_hSearch, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        
        CheckAndOptimizeLegacyAssets(hwnd);
        
        RefreshAssets(true); return 0;
    }

    case WM_NCHITTEST: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        ScreenToClient(hwnd, &pt);
        RECT rc; GetClientRect(hwnd, &rc);

        if (pt.y >= rc.bottom - RESIZE_MARGIN) {
            if (pt.x <= RESIZE_MARGIN) return HTBOTTOMLEFT;
            if (pt.x >= rc.right - RESIZE_MARGIN) return HTBOTTOMRIGHT;
            return HTBOTTOM;
        }
        if (pt.y <= RESIZE_MARGIN) {
            if (pt.x <= RESIZE_MARGIN) return HTTOPLEFT;
            if (pt.x >= rc.right - RESIZE_MARGIN) return HTTOPRIGHT;
            return HTTOP;
        }
        if (pt.x <= RESIZE_MARGIN) return HTLEFT;
        if (pt.x >= rc.right - RESIZE_MARGIN) return HTRIGHT;

        if (pt.y < TITLE_H) {
            if (pt.x > rc.right - 40) return HTCLIENT; 
            return HTCAPTION;
        }
        return HTCLIENT;
    }

    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp);
    case WM_NCACTIVATE: return DefWindowProc(hwnd, msg, wp, lp);
    case WM_MOUSEACTIVATE: return MA_ACTIVATE;

    case WM_SIZE: { 
        RECT rc; GetClientRect(hwnd, &rc); int y = rc.bottom - FOOTER_H + 8; 
        if (g_hCombo) MoveWindow(g_hCombo, 10, y, 140, 300, TRUE); 
        if (g_hSearch) { int sw = (rc.right - 150) - 160; if (sw < 50) sw = 50; MoveWindow(g_hSearch, 160, y, sw, 24, TRUE); } 
        UpdateScrollBar(hwnd);
        InvalidateRect(hwnd, nullptr, FALSE); return 0; 
    }
    case WM_EXITSIZEMOVE:
        SaveWindowPos(hwnd, false);
        return 0;
    
    case WM_TIMER: 
        if (wp == ID_TIMER_HOVER && g_hoverIndex != -1 && g_hoverIndex < (int)g_displayAssets.size()) {
            Asset* p = g_displayAssets[g_hoverIndex]; if (p->frameCount > 1) {
                p->currentFrame = (p->currentFrame + 1) % p->frameCount; InvalidateRect(hwnd, NULL, FALSE);
                UINT d = p->frameDelays ? p->frameDelays[p->currentFrame] : 100;
                UINT sc = (UINT)(d * 100.0f / (float)(std::max)(1, g_gifSpeedPercent));
                SetTimer(hwnd, ID_TIMER_HOVER, (std::max)(10u, sc), NULL);
            }
        }
        if (wp == ID_TIMER_TOOLTIP) {
            KillTimer(hwnd, ID_TIMER_TOOLTIP);
            if (g_tooltipTargetIndex != -1 && g_tooltipTargetIndex < (int)g_displayAssets.size()) {
                Asset* p = g_displayAssets[g_tooltipTargetIndex];
                std::wstring mainT, subT;
                
                if (p->isMulti) {
                    mainT = L"複数オブジェクト (Multi Object)";
                    subT = L"構造を維持するため、保存時の長さで固定されます";
                } else if (p->isFixedFrame) {
                    mainT = L"単体オブジェクト (Single Object)";
                    subT = L"保存時のフレーム数を維持して配置します";
                } else {
                    mainT = L"単体オブジェクト (Single Object)";
                    subT = L"現在のレイヤー表示の拡大率によって配置の長さが自動で伸縮します";
                }
                
                POINT pt; GetCursorPos(&pt);
                ShowMyTooltip(pt.x, pt.y, mainT, subT);
            }
        }
        return 0;

    case WM_MOUSEWHEEL: {
        int delta = GET_WHEEL_DELTA_WPARAM(wp);
        UINT lines = 3;
        SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        int step = (lines == WHEEL_PAGESCROLL) ? (ITEM_HEIGHT * 3) : ((int)lines * SCROLL_SPD);
        if (step <= 0) step = SCROLL_SPD;
        SetScrollY(hwnd, g_scrollY - (delta / WHEEL_DELTA) * step);
        return 0;
    }

    case WM_VSCROLL: {
        SCROLLINFO si = {0};
        si.cbSize = sizeof(si);
        si.fMask = SIF_ALL;
        GetScrollInfo(hwnd, SB_VERT, &si);

        int newY = g_scrollY;
        switch (LOWORD(wp)) {
        case SB_TOP: newY = 0; break;
        case SB_BOTTOM: newY = GetMaxScrollY(hwnd); break;
        case SB_LINEUP: newY -= SCROLL_SPD; break;
        case SB_LINEDOWN: newY += SCROLL_SPD; break;
        case SB_PAGEUP: newY -= (int)si.nPage; break;
        case SB_PAGEDOWN: newY += (int)si.nPage; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: newY = si.nTrackPos; break;
        default: break;
        }

        SetScrollY(hwnd, newY);
        return 0;
    }
    
    case WM_KILLFOCUS:
        ReleaseCapture();
        g_isDragCheck = false;
        g_hoverIndex = -1;
        HideMyTooltip(); 
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;

    case WM_CAPTURECHANGED:
        g_isDragCheck = false;
        return 0;

    case WM_RBUTTONDOWN:
    case WM_KEYDOWN:
        if (g_isDragCheck) {
            ReleaseCapture();
            g_isDragCheck = false;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        if (msg == WM_RBUTTONDOWN) goto RBUTTON_HANDLER; 
        return 0;

    case WM_NCRBUTTONUP:
        if (wp == HTCAPTION) {
            ShowSettingsRefreshMenu(hwnd, GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
            return 0;
        }
        return DefWindowProc(hwnd, msg, wp, lp);

    case WM_RBUTTONUP: {
RBUTTON_HANDLER:
        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        if (y >= 0 && y < TITLE_H) {
            POINT pt; GetCursorPos(&pt);
            ShowSettingsRefreshMenu(hwnd, pt.x, pt.y);
        }
        else if (y > TITLE_H && y < rc.bottom - FOOTER_H) {
            int cols = (std::max)(1, (int)rc.right / MIN_ITEM_WIDTH); int cw = (int)rc.right / cols;
            int idx = ((y - TITLE_H + g_scrollY) / ITEM_HEIGHT) * cols + (x / cw);
            HMENU hm = CreatePopupMenu();
            if (idx >= 0 && idx < (int)g_displayAssets.size()) { 
                g_contextTargetIndex = idx; Asset* p = g_displayAssets[idx]; 
                
                // Multiの場合はフレーム固定の切り替えを無効化
                UINT flags = p->isMulti ? MF_GRAYED : MF_STRING;
                if (p->isFixedFrame) flags |= MF_CHECKED;
                
                AppendMenuW(hm, MF_STRING, IDM_EDIT, L"編集"); 
                AppendMenuW(hm, MF_STRING, IDM_FAVORITE, p->isFavorite ? L"お気に入り解除" : L"お気に入り登録"); 
                AppendMenuW(hm, flags, IDM_TOGGLE_FIXED, L"フレーム数を固定する");
                AppendMenuW(hm, MF_STRING, IDM_DELETE, L"削除"); 
            }
            else {
                AppendMenuW(hm, MF_STRING, IDM_SETTINGS, L"設定...");
                AppendMenuW(hm, MF_STRING, ID_BTN_REFRESH, L"更新");
                AppendMenuW(hm, MF_STRING, IDM_OPEN_FOLDER, L"MyAssetフォルダを開く");
            }
            POINT pt; GetCursorPos(&pt); TrackPopupMenu(hm, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL); DestroyMenu(hm);
        } return 0;
    }
    case WM_COMMAND: {
        int id = LOWORD(wp); int code = HIWORD(wp);
        if (id == IDM_SETTINGS) OpenSettings();
        else if (id == ID_BTN_REFRESH) RefreshAssets(true);
        else if (id == IDM_OPEN_FOLDER) OpenAssetFolderInExplorer(hwnd);
        else if (id == IDM_EDIT && g_contextTargetIndex != -1) { Asset* p = g_displayAssets[g_contextTargetIndex]; if (p->pImage) { delete p->pImage; p->pImage = nullptr; } g_editOrgPath = p->path; g_editName = p->name; g_editCat = p->category; OpenAddDialog("", true); }
        else if (id == IDM_FAVORITE && g_contextTargetIndex != -1) { ToggleFavorite(g_displayAssets[g_contextTargetIndex]->path); RefreshAssets(false); }
        else if (id == IDM_TOGGLE_FIXED && g_contextTargetIndex != -1) { 
            // Multiでない場合のみトグル可能
            if (!g_displayAssets[g_contextTargetIndex]->isMulti) {
                ToggleFixedFrame(g_displayAssets[g_contextTargetIndex]->path); 
                RefreshAssets(false); 
            }
        }
        else if (id == IDM_DELETE && g_contextTargetIndex != -1) { 
            Asset* p = g_displayAssets[g_contextTargetIndex]; 
            if (ShowDarkMsg(hwnd, L"削除しますか？", L"確認", MB_YESNO) == IDYES) { 
                if (p->pImage) { delete p->pImage; p->pImage = nullptr; } 
                DeleteFileW(p->path.c_str()); 
                std::wstring base = p->path.substr(0, p->path.find_last_of(L"."));
                DeleteFileW((base + L".png").c_str()); DeleteFileW((base + L".gif").c_str());
                RefreshAssets(false); 
            } 
        }
        else if (id == ID_COMBO_CATEGORY && code == CBN_SELCHANGE) { int idx = SendMessage(g_hCombo, CB_GETCURSEL, 0, 0); if (idx >= 0 && idx < (int)g_categories.size()) { g_currentCategory = g_categories[idx]; UpdateDisplayList(); } }
        else if (id == ID_EDIT_SEARCH && code == EN_CHANGE) { int len = GetWindowTextLengthW(g_hSearch); std::vector<wchar_t> buf(len + 1); GetWindowTextW(g_hSearch, buf.data(), len + 1); g_searchQuery = buf.data(); UpdateDisplayList(); }
        return 0;
    }
    case WM_PAINT: { PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); HDC mc = CreateCompatibleDC(hdc); HBITMAP mb = CreateCompatibleBitmap(hdc, rc.right, rc.bottom); SelectObject(mc, mb); DrawContent(mc, rc.right, rc.bottom); BitBlt(hdc, 0, 0, rc.right, rc.bottom, mc, 0, 0, SRCCOPY); DeleteObject(mb); DeleteDC(mc); EndPaint(hwnd, &ps); return 0; }
    case WM_LBUTTONDOWN: {
        SetFocus(hwnd); 

        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        if (y > TITLE_H && y < rc.bottom - FOOTER_H) {
            SetCapture(hwnd); 
            int cols = (std::max)(1, (int)rc.right / MIN_ITEM_WIDTH); int cw = (int)rc.right / cols;
            int idx = ((y - TITLE_H + g_scrollY) / ITEM_HEIGHT) * cols + (x / cw);
            if (idx >= 0 && idx < (int)g_displayAssets.size()) { g_selectedIndex = idx; g_isDragCheck = true; g_dragStartPt = {x, y}; InvalidateRect(hwnd, nullptr, FALSE); }
        } 
        else if (y >= rc.bottom - FOOTER_H) {
            if (x >= rc.right - 70 && x <= rc.right - 10) RefreshAssets(true);
            else if (x >= rc.right - 140 && x <= rc.right - 80) { g_sortFavFirst = !g_sortFavFirst; UpdateDisplayList(); }
        }
        return 0;
    }
    case WM_MOUSELEAVE: { 
        if (g_hoverIndex != -1) { KillTimer(hwnd, ID_TIMER_HOVER); g_hoverIndex = -1; InvalidateRect(hwnd, NULL, FALSE); } 
        g_isMouseTracking = false; 
        HideMyTooltip(); 
        return 0; 
    }
    case WM_LBUTTONUP: { 
        ReleaseCapture(); 
        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc); 
        if (y < TITLE_H && x > rc.right - 40) { SaveWindowPos(hwnd, false); ShowWindow(hwnd, SW_HIDE); } 
        g_isDragCheck = false; return 0; 
    }
    case WM_MOUSEMOVE: {
        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        if (!g_isMouseTracking) { TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT), TME_LEAVE, hwnd, 0 }; TrackMouseEvent(&tme); g_isMouseTracking = true; }
        
        bool onMark = false;
        if (y > TITLE_H && y < rc.bottom - FOOTER_H) {
            int cols = (std::max)(1, (int)rc.right / MIN_ITEM_WIDTH); int cw = (int)rc.right / cols;
            int row = (y - TITLE_H + g_scrollY) / ITEM_HEIGHT; int col = x / cw;
            int idx = row * cols + col;
            
            if (idx >= 0 && idx < (int)g_displayAssets.size()) { 
                if (g_hoverIndex != idx) { g_hoverIndex = idx; SetTimer(hwnd, ID_TIMER_HOVER, 100, NULL); } 

                int itemX = col * cw; int itemY = row * ITEM_HEIGHT + (TITLE_H - g_scrollY);
                int right = itemX + cw - 2; int bottom = itemY + ITEM_HEIGHT - 2;
                RECT rcMark = { right - 30, bottom - 30, right, bottom };
                POINT pt = {x, y};
                
                if (PtInRect(&rcMark, pt)) {
                    onMark = true;
                    if (g_tooltipTargetIndex != idx) {
                        HideMyTooltip();
                        g_tooltipTargetIndex = idx;
                        SetTimer(hwnd, ID_TIMER_TOOLTIP, 1000, NULL);
                    }
                }
            } else { if (g_hoverIndex != -1) { KillTimer(hwnd, ID_TIMER_HOVER); g_hoverIndex = -1; } }
        } else { if (g_hoverIndex != -1) { KillTimer(hwnd, ID_TIMER_HOVER); g_hoverIndex = -1; } }
        
        if (!onMark) {
            KillTimer(hwnd, ID_TIMER_TOOLTIP);
            HideMyTooltip();
            g_tooltipTargetIndex = -1;
        }

        if (g_isDragCheck && g_selectedIndex != -1 && g_selectedIndex < (int)g_displayAssets.size()) { 
            if (abs(x - g_dragStartPt.x) > 5 || abs(y - g_dragStartPt.y) > 5) { 
                if (!g_lastTempPath.empty()) { DeleteFileW(g_lastTempPath.c_str()); g_lastTempPath = L""; }
                ReleaseCapture(); g_isDragCheck = false; 

                Asset* p = g_displayAssets[g_selectedIndex];
                std::string content = ReadFileContent(p->path);

                int headerCount = 0;
                size_t pos = 0;
                while ((pos = content.find("[Object]", pos)) != std::string::npos) {
                    char next = (pos + 8 < content.length()) ? content[pos + 8] : 0;
                    if (next == '\r' || next == '\n') headerCount++;
                    pos += 8;
                }
                bool hasIndex = (content.find("[0]") != std::string::npos);

                if (headerCount >= 2 && !hasIndex) {
                    int res = ShowDarkMsg(hwnd, 
                        L"古い形式のアセットを検知しました\n\nこのアセット（複数オブジェクト）は、以前のバージョンで保存されたため\nAviUtlで正常に読み込めない可能性があります。\n現在の仕様に合わせてデータを修復しますか？\n\n※元のレイヤー構造は復元できないため、階段状に配置されます。",
                        L"アセット修復の確認", MB_YESNO);
                        
                    if (res == IDYES) {
                        content = RepairLegacyObjectContent(content);
                        WriteFileContent(p->path, content);
                        p->isMulti = true;
                    }
                }
                else if (hasIndex && content.find("[1]") == std::string::npos) {
                    content = NormalizeSingleObjectContent(content);
                    WriteFileContent(p->path, content);
                    p->isMulti = false;
                }

                std::wstring dragPath = p->path;
                std::wstring tempPath;
                bool isModernFormat = (content.find("[Object]") != std::string::npos); 
                bool isMulti = (content.find("[0]") != std::string::npos);

                if (!p->isFixedFrame && !isMulti && isModernFormat) {
                    RemoveLinesStartingWith(content, "frame=");
                    std::wstring dir = p->path.substr(0, p->path.find_last_of(L"\\") + 1);
                    tempPath = dir + L"_auto_" + p->name + L".object"; 
                    WriteFileContent(tempPath, content);
                    dragPath = tempPath;
                    g_lastTempPath = tempPath; 
                }

                IDataObject* po; 
                if (SUCCEEDED(CreateFileDropDataObject(dragPath, &po))) { 
                    CDropSource* ps = new CDropSource(); 
                    DWORD de; 
                    DoDragDrop(po, ps, DROPEFFECT_COPY, &de); 
                    po->Release(); 
                } 
                g_selectedIndex = -1;
                InvalidateRect(hwnd, NULL, FALSE);
            } 
        }
        return 0;
    }
    case WM_DESTROY: 
        SaveWindowPos(hwnd, false); 
        if (!g_lastTempPath.empty()) DeleteFileW(g_lastTempPath.c_str());
        if(g_hDlg) DestroyWindow(g_hDlg);
        if(g_hSettingDlg) DestroyWindow(g_hSettingDlg);
        if(g_hSnipWnd) DestroyWindow(g_hSnipWnd);
        if(g_hInfoWnd) DestroyWindow(g_hInfoWnd);
        if(g_hTooltip) DestroyWindow(g_hTooltip); 
        GdiplusShutdown(g_gdiplusToken); 
        PostQuitMessage(0); 
        return 0;
    case WM_CTLCOLOREDIT: case WM_CTLCOLORLISTBOX: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkColor(hdc, COL_INPUT_BG); return (LRESULT)g_hBrInputBg; }
    default: return DefWindowProc(hwnd, msg, wp, lp);
    }
}

// ------------------------------------------------------------
// 9. AviUtl エントリポイント
// ------------------------------------------------------------
void OnAddAsset(EDIT_SECTION_SAFE* edit) {
    if (!edit || !edit->get_selected_object_num) return; 
    int count = edit->get_selected_object_num(); 
    if (count <= 0) { ShowDarkMsg(NULL, L"オブジェクトを選択してください", L"Info", MB_OK); return; }
    
    std::string bodyData;
    int minLayer = 99999;

    for(int i = 0; i < count; i++) {
        void* obj = edit->get_selected_object(i);
        if (obj) {
            int layer = GetObjectLayerIndex(obj);
            if (layer < minLayer) minLayer = layer;
        }
    }
    
    if (count == 1) {
        void* obj = edit->get_selected_object(0);
        if (obj) {
            LPCSTR raw = edit->get_object_alias(obj);
            if (raw) bodyData = raw;
        }
    } 
    else {
        std::vector<int> relativeGroups(count, 0);
        std::map<int, int> groupIdMap;
        std::map<int, int> groupCounts;
        std::vector<int> originalGroups(count, 0);
        int nextGroupId = 1;

        for (int i = 0; i < count; i++) {
            void* obj = edit->get_selected_object(i);
            if (!obj) continue;
            LPCSTR raw = edit->get_object_alias(obj);

            int originalGroupId = raw ? ExtractHeaderIntValue(raw, "group=") : 0;
            if (originalGroupId <= 0) originalGroupId = GetObjectGroupIndex(obj);
            originalGroups[i] = originalGroupId;
            if (originalGroupId > 0) groupCounts[originalGroupId]++;
        }

        for (int i = 0; i < count; i++) {
            int originalGroupId = originalGroups[i];
            if (originalGroupId <= 0) continue;
            if (groupCounts[originalGroupId] < 2) continue;

            auto it = groupIdMap.find(originalGroupId);
            if (it == groupIdMap.end()) {
                groupIdMap[originalGroupId] = nextGroupId;
                relativeGroups[i] = nextGroupId;
                nextGroupId++;
            } else {
                relativeGroups[i] = it->second;
            }
        }

        for(int i = 0; i < count; i++) { 
            void* obj = edit->get_selected_object(i); 
            if (obj) { 
                LPCSTR raw = edit->get_object_alias(obj); 
                if (raw) { 
                    std::string s = raw; 
                    std::string headerNew = "[" + std::to_string(i) + "]";
                    std::string sectionNew = "[" + std::to_string(i) + ".";
                    int currentLayer = GetObjectLayerIndex(obj);
                    int relativeLayer = (currentLayer - minLayer) + 1;
                    ReplaceStringAll(s, "[Object.", sectionNew);
                    ReplaceStringAll(s, "[Object]", headerNew);
                    RemoveLinesStartingWith(s, "group=");
                    size_t pos = s.find(headerNew);
                    if (pos != std::string::npos) {
                        std::string headerLines = "\r\nlayer=" + std::to_string(relativeLayer);
                        if (relativeGroups[i] > 0) {
                            headerLines += "\r\ngroup=" + std::to_string(relativeGroups[i]);
                        }
                        s.insert(pos + headerNew.length(), headerLines);
                    }
                    bodyData += s + "\r\n"; 
                } 
            } 
        }
    }
    OpenAddDialog(bodyData, false);
}

void OnShowMainWindow(EDIT_SECTION_SAFE* edit) { if (g_hwnd) { ShowWindow(g_hwnd, SW_RESTORE); ShowWindow(g_hwnd, SW_SHOW); SetForegroundWindow(g_hwnd); } }

void CreatePluginWindow() { LoadWindowConfig(); WNDCLASSW wc = {0}; wc.lpfnWndProc = WndProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAssetMgr_Modern"; wc.hCursor = LoadCursor(nullptr, IDC_ARROW); wc.hbrBackground = g_hBrBg; wc.style = CS_HREDRAW | CS_VREDRAW; RegisterClassW(&wc); g_hwnd = CreateWindowExW(WS_EX_TOPMOST, L"MyAssetMgr_Modern", L"My Asset Manager", WS_POPUP | WS_CLIPCHILDREN | WS_THICKFRAME | WS_VSCROLL, g_winX, g_winY, g_winW, g_winH, nullptr, nullptr, g_hInst, nullptr); }

extern "C" __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host) { 
    HMODULE hm; GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)RegisterPlugin, &hm); g_hInst = (HINSTANCE)hm;
    g_host = host; host->set_plugin_information(L"My Asset Manager v1.5"); 
    if (host->create_edit_handle) g_editHandle = (EDIT_HANDLE_SAFE*)host->create_edit_handle();
    if (host->f2) {
        auto registerOutput = (void(*)(OUTPUT_PLUGIN_TABLE*))host->f2;
        registerOutput(&g_outputPluginTable);
    }
    RegisterCustomDialogs(); CreatePluginWindow(); 
    if (g_hwnd) host->register_window_client(L"My Asset Manager", g_hwnd); 
    if (host->register_edit_menu) host->register_edit_menu(L"表示 > My Asset Manager", OnShowMainWindow); 
    if (host->register_layer_menu) { host->register_layer_menu(L"MyAsset: 一覧を開く", OnShowMainWindow); host->register_layer_menu(L"MyAsset: 追加", OnAddAsset); }
    if (host->register_object_menu) { host->register_object_menu(L"MyAsset: 一覧を開く", OnShowMainWindow); host->register_object_menu(L"MyAsset: 追加", OnAddAsset); }
}
