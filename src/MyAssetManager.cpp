// ------------------------------------------------------------
// MyAssetManager.cpp
// v1.8.3
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
#include <cwctype>
#include <climits>
#include <cstdint>
#include <fstream>
#include <functional>
#include <iomanip>
#include <cstdio>
#include <cmath>

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
#define WM_APP_DEBUG_LAYER_COMPARE (WM_APP + 77)
#define WM_APP_BUILD_ADD_ASSET  (WM_APP + 78)
#define ID_TIMER_HOVER          998
#define ID_TIMER_ADD_PREVIEW    997 
#define ID_TIMER_TOOLTIP        996 
#define ID_TIMER_GIF_HINT       995
#define ID_TIMER_GIF_HINT_POLL  994
#define ID_TIMER_HIDE_MAIN_AFTER_EXPORT 993
#define ID_TIMER_LAYER_API_COMPARE 992

#define ID_COMBO_CATEGORY 200
#define ID_BTN_FAV_SORT   201
#define ID_BTN_REFRESH    202
#define ID_EDIT_SEARCH    203
#define ID_BTN_SNIP       204 
#define ID_BTN_GIF_EXPORT 205
#define ID_BTN_PRESET_A   206
#define ID_BTN_PRESET_B   207
#define ID_BTN_PRESET_C   208
#define ID_BTN_CAPTURE_SAVE 209
#define ID_BTN_CAPTURE_REUSE 210

#define IDM_FAVORITE      300
#define IDM_EDIT          301
#define IDM_DELETE        302
#define IDM_SETTINGS      303
#define IDM_TOGGLE_FIXED  304 
#define IDM_OPEN_FOLDER   305
#define IDM_INFO_RELEASES 306
#define IDM_INFO_BUGREPORT 307
#define IDM_SORT_NAME     308
#define IDM_SORT_FAVORITE 309
#define IDM_SORT_CATEGORY 310
#define IDM_SORT_CUSTOM   314
#define IDM_APPLY_TEXT_STYLE 311
#define IDM_TEXT_STYLE_COMMIT 312
#define IDM_TEXT_STYLE_CANCEL 313

#define IDC_EDIT_NAME     101
#define IDC_COMBO_CAT     102
#define ID_BTN_SAVE       103
#define ID_BTN_CANCEL     104
#define IDC_SLIDER_SPEED  105
#define IDC_EDIT_SPEED    106 
#define IDC_CHK_GIF_ORIGINAL 107
#define ID_SETTING_SIDEBAR 108
#define IDC_RAD_PREVIEW_ALL 109
#define IDC_RAD_PREVIEW_HOVER 110
#define IDC_CHK_TEXT_STYLE 111
#define IDC_CHK_ENABLE_TEXT_STYLE 112
#define IDC_ST_INHERIT_TITLE 113
#define IDC_RAD_INSERT_END 114
#define IDC_RAD_INSERT_AFTER_STD 115
#define IDC_CHK_CLEAR_EFFECTS_FROM2 116
#define IDC_CHK_HIDE_TEXTSTYLE_IN_LIST 117
#define IDC_RAD_PREVIEW_PERF_NORMAL 118
#define IDC_RAD_PREVIEW_PERF_LOW 119
#define IDC_SLIDER_PREVIEW_ZOOM 120
#define IDC_EDIT_PREVIEW_ZOOM 121
#define IDC_CHK_PREVIEW_THUMB_ONLY 122
#define IDC_ST_THEME_BASE 1400
#define IDC_EDIT_THEME_BASE 1450
#define IDC_BTN_PICK_THEME_BASE 1500
#define IDC_BTN_THEME_RESET 1550
#define IDC_CHK_INH_GROUP_BASE 1200
#define IDC_BTN_INH_EXPAND_BASE 1220
#define IDC_CHK_INH_ITEM_BASE 1300

#define ID_BTN_MSG_OK     401
#define ID_BTN_MSG_YES    402
#define ID_BTN_MSG_NO     403
#define ID_BTN_GUIDE_OK   404
#define ID_BTN_GUIDE_CANCEL 405
#define IDC_CHK_GUIDE_HIDE 406

// --- UI定数 ---
static constexpr int TITLE_H        = 32;
static constexpr int TAB_H          = 30;
static constexpr int STYLE_ACTION_H = 36;
static constexpr int FOOTER_H       = 40;
static constexpr int ITEM_HEIGHT    = 90; 
static constexpr int THUMB_W        = 120;
static constexpr int THUMB_H        = 68;
static constexpr int SCROLL_SPD     = 30;
static constexpr int MIN_ITEM_WIDTH = 240;
static constexpr int RESIZE_MARGIN  = 6;

static constexpr COLORREF DEF_COL_BG        = RGB(30, 30, 30);
static constexpr COLORREF DEF_COL_TITLE_BG  = RGB(45, 45, 45);
static constexpr COLORREF DEF_COL_BORDER    = RGB(60, 60, 60);
static constexpr COLORREF DEF_COL_ITEM_BG   = RGB(40, 40, 40);
static constexpr COLORREF DEF_COL_ITEM_SEL  = RGB(60, 70, 90);
static constexpr COLORREF DEF_COL_TEXT      = RGB(220, 220, 220);
static constexpr COLORREF DEF_COL_SUBTEXT   = RGB(160, 160, 160);
static constexpr COLORREF DEF_COL_FOOTER    = RGB(40, 40, 40);
static constexpr COLORREF DEF_COL_BTN_BG    = RGB(60, 60, 60);
static constexpr COLORREF DEF_COL_BTN_PUSH  = RGB(100, 100, 100);
static constexpr COLORREF DEF_COL_BTN_ACT   = RGB(100, 120, 200);
static constexpr COLORREF DEF_COL_INPUT_BG  = RGB(20, 20, 20);
static constexpr COLORREF DEF_COL_FIXED_TEXT = RGB(31, 205, 219); 
static constexpr COLORREF DEF_COL_TIP_BG     = RGB(50, 50, 50);
static constexpr COLORREF DEF_COL_TIP_BORDER = RGB(100, 100, 100);
static constexpr COLORREF DEF_COL_TIP_TEXT   = RGB(230, 230, 230);

static COLORREF COL_BG        = DEF_COL_BG;
static COLORREF COL_TITLE_BG  = DEF_COL_TITLE_BG;
static COLORREF COL_BORDER    = DEF_COL_BORDER;
static COLORREF COL_ITEM_BG   = DEF_COL_ITEM_BG;
static COLORREF COL_ITEM_SEL  = DEF_COL_ITEM_SEL;
static COLORREF COL_TEXT      = DEF_COL_TEXT;
static COLORREF COL_SUBTEXT   = DEF_COL_SUBTEXT;
static COLORREF COL_FOOTER    = DEF_COL_FOOTER;
static COLORREF COL_BTN_BG    = DEF_COL_BTN_BG;
static COLORREF COL_BTN_PUSH  = DEF_COL_BTN_PUSH;
static COLORREF COL_BTN_ACT   = DEF_COL_BTN_ACT;
static COLORREF COL_INPUT_BG  = DEF_COL_INPUT_BG;
static COLORREF COL_FIXED_TEXT = DEF_COL_FIXED_TEXT;
static COLORREF COL_TIP_BG     = DEF_COL_TIP_BG;
static COLORREF COL_TIP_BORDER = DEF_COL_TIP_BORDER;
static COLORREF COL_TIP_TEXT   = DEF_COL_TIP_TEXT;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#define MYASSET_VERSION_W L"1.8.3"
static const wchar_t* kGithubReleaseBase = L"https://github.com/Natadecoco2539/AviUtl2-MyAssetManager/releases/tag/v";
static const wchar_t* kGithubBugReportUrl = L"https://github.com/Natadecoco2539/AviUtl2-MyAssetManager/issues/new/choose";

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// --- 構造体 ---
struct Asset {
    std::wstring name; 
    std::wstring path; 
    std::wstring imagePath; 
    std::wstring category;
    bool isFavorite; 
    bool isFixedFrame; 
    bool isTextStyle;
    bool isMulti;      
    Image* pImage; 
    UINT frameCount; 
    UINT* frameDelays; 
    int currentFrame; 
    double previewCarryMs;
    UINT previewCycleMs;
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
static bool g_addDialogRangeValid = false;
static int g_addDialogRangeStart = 0;
static int g_addDialogRangeEnd = 0;
static HINSTANCE g_hInst = nullptr;
static HWND g_hwnd = nullptr;
static HWND g_hDlg = nullptr;
static HWND g_hSettingDlg = nullptr;
static HWND g_hTextStyleQuickDlg = nullptr;
static HWND g_hSnipWnd = nullptr;
static HWND g_hInfoWnd = nullptr;
static HWND g_hCombo = nullptr;
static HWND g_hSearch = nullptr;
static HWND g_hTooltip = nullptr; 
static HWND g_hostMainWnd = nullptr;

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
static std::set<std::wstring> g_textStylePaths;
// デバッグ時trueに変更
static constexpr bool kEnableDebugByCode = false;
static std::map<std::wstring, std::vector<std::wstring>> g_customOrderByCategory;
static int g_sortMode = 0;
static int g_gifSpeedPercent = 100;
static bool g_showGifExportGuide = true;
static bool g_gifExportKeepOriginal = false;
static int g_previewPlaybackMode = 1;
static int g_previewPerfMode = 0;
static int g_previewZoomPercent = 100;
static bool g_previewThumbOnly = false;
static ULONGLONG g_previewAllLastTick = 0;
static ULONGLONG g_previewHoverLastTick = 0;
static bool g_enableTextStyle = true;
static int g_assetViewTab = 0;
static bool g_addAsTextStyle = false;
static bool g_enableGroupDebugDump = false;
static bool g_enableGifRangeDebug = false;
static bool g_mainWasVisibleBeforeAddDialog = false;
static bool g_suppressMainShow = false;
static bool g_enableMainWindowTrace = false;
static int g_hideMainAfterExportTicks = 0;
static int g_settingsCategory = 0;
static bool g_pendingLayerApiCompare = false;
static std::vector<void*> g_pendingLayerApiCompareObjs;
struct PendingAddObjSnapshot {
    void* obj = nullptr;
    std::string alias;
};
static std::vector<PendingAddObjSnapshot> g_pendingAddObjSnapshots;
static bool g_addDialogFromMyAssetAdd = false;

static int g_winX = 100, g_winY = 100, g_winW = 360, g_winH = 550;
static int g_dlgX = -1, g_dlgY = -1;

static int g_scrollY = 0, g_selectedIndex = -1, g_contextTargetIndex = -1, g_hoverIndex = -1;
static std::string g_tempAliasData; 
static std::wstring g_editOrgPath, g_editName, g_editCat, g_tempImgPath, g_msgTitle, g_msgText;
static bool g_isDragCheck = false, g_isImageRemoved = false, g_isSnipping = false;
static bool g_isReorderDrag = false;
static int g_reorderFromIndex = -1;
static int g_reorderInsertPos = -1;
static POINT g_dragStartPt = {0}, g_snipStart = {0}, g_snipEnd = {0};
static int g_msgResult = 0;
static int g_guideResult = IDCANCEL;
static bool g_guideHideNext = false;
static bool g_isMouseTracking = false;
static RECT g_capturePresetRects[3] = {};
static bool g_capturePresetValid[3] = { false, false, false };
static int g_capturePresetSlot = 0;
static RECT g_lastCapturedRect = {};
static bool g_lastCapturedRectValid = false;

static UINT GetPreviewAllTickMs() {
    return (g_previewPerfMode == 1) ? 100u : 40u;
}

static int ScalePreviewValue(int v) {
    int pct = (std::max)(50, (std::min)(200, g_previewZoomPercent));
    return (v * pct + 50) / 100;
}

static int GetListItemHeight() {
    return g_previewThumbOnly ? ScalePreviewValue(THUMB_H + 14) : ScalePreviewValue(ITEM_HEIGHT);
}

static int GetListItemMinWidth() {
    int w = g_previewThumbOnly ? (ScalePreviewValue(THUMB_W) + 14) : ScalePreviewValue(MIN_ITEM_WIDTH);
    return (std::max)(80, w);
}

static int GetThumbWScaled() {
    return (std::max)(16, ScalePreviewValue(THUMB_W));
}

static int GetThumbHScaled() {
    return (std::max)(12, ScalePreviewValue(THUMB_H));
}

static int GetListTopY() { return g_enableTextStyle ? (TITLE_H + TAB_H) : TITLE_H; }
static int GetStyleActionAreaHeight() { return (g_enableTextStyle && g_assetViewTab == 1) ? STYLE_ACTION_H : 0; }
static int GetBottomReservedHeight() { return FOOTER_H + GetStyleActionAreaHeight(); }

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
static bool g_textStylePending = false;
static int g_textStylePendingLayer = -1;
static int g_textStylePendingStart = -1;
static int g_textStylePendingEnd = -1;
static std::string g_textStyleOriginalAlias;
static int g_inheritExpandedGroup = -1;
static int g_quickInheritExpandedGroup = -1;
static int g_textStyleInsertMode = 0;
static bool g_textStyleClearEffectsFrom2 = false;
static bool g_hideTextStyleInMainList = false;
static COLORREF g_colorPickerCustom[16] = {};

struct ThemeColorDef {
    const wchar_t* label;
    const wchar_t* key;
    COLORREF* value;
    COLORREF defValue;
};
static ThemeColorDef kThemeColors[] = {
    { L"背景", L"ThemeBg", &COL_BG, DEF_COL_BG },
    { L"タイトルバー", L"ThemeTitleBg", &COL_TITLE_BG, DEF_COL_TITLE_BG },
    { L"枠線", L"ThemeBorder", &COL_BORDER, DEF_COL_BORDER },
    { L"フッター帯", L"ThemeFooter", &COL_FOOTER, DEF_COL_FOOTER },
    { L"アセット背景", L"ThemeItemBg", &COL_ITEM_BG, DEF_COL_ITEM_BG },
    { L"アセット選択", L"ThemeItemSel", &COL_ITEM_SEL, DEF_COL_ITEM_SEL },
    { L"文字", L"ThemeText", &COL_TEXT, DEF_COL_TEXT },
    { L"サブ文字", L"ThemeSubText", &COL_SUBTEXT, DEF_COL_SUBTEXT },
    { L"入力背景", L"ThemeInputBg", &COL_INPUT_BG, DEF_COL_INPUT_BG },
    { L"ボタン", L"ThemeBtnBg", &COL_BTN_BG, DEF_COL_BTN_BG },
    { L"ボタン(選択)", L"ThemeBtnAct", &COL_BTN_ACT, DEF_COL_BTN_ACT },
};
static constexpr int kThemeColorCount = (int)(sizeof(kThemeColors) / sizeof(kThemeColors[0]));

struct InheritGroupDef {
    const wchar_t* label;
    int start;
    int count;
};
struct InheritItemDef {
    const wchar_t* effect;
    const wchar_t* key;
    const wchar_t* label;
    int group;
};

static constexpr int kInheritItemCount = 29;
static constexpr int kInheritGroupCount = 7;
static bool g_inheritItemEnabled[kInheritItemCount] = {
    true,  // テキスト
    false, false, false, false, false, false, false, false, false,
    false, false,
    false, false, false, false,
    false, false, false, false, false, false,
    false, false, false, false, false,
    false, false
};

static const InheritGroupDef kInheritGroups[kInheritGroupCount] = {
    { L"テキスト内容", 0, 1 },
    { L"文字基本", 1, 9 },
    { L"文字色", 10, 2 },
    { L"テキスト動作", 12, 4 },
    { L"位置・中心", 16, 6 },
    { L"回転・変形", 22, 5 },
    { L"描画合成", 27, 2 },
};

static const InheritItemDef kInheritItems[kInheritItemCount] = {
    { L"テキスト", L"テキスト", L"テキスト", 0 },
    { L"テキスト", L"サイズ", L"サイズ", 1 },
    { L"テキスト", L"字間", L"字間", 1 },
    { L"テキスト", L"行間", L"行間", 1 },
    { L"テキスト", L"表示速度", L"表示速度", 1 },
    { L"テキスト", L"フォント", L"フォント", 1 },
    { L"テキスト", L"文字装飾", L"文字装飾", 1 },
    { L"テキスト", L"文字揃え", L"文字揃え", 1 },
    { L"テキスト", L"B", L"B", 1 },
    { L"テキスト", L"I", L"I", 1 },
    { L"テキスト", L"文字色", L"文字色", 2 },
    { L"テキスト", L"影・縁色", L"影・縁色", 2 },
    { L"テキスト", L"文字毎に個別オブジェクト", L"文字毎に個別オブジェクト", 3 },
    { L"テキスト", L"自動スクロール", L"自動スクロール", 3 },
    { L"テキスト", L"移動座標上に表示", L"移動座標上に表示", 3 },
    { L"テキスト", L"オブジェクトの長さを自動調節", L"オブジェクトの長さを自動調節", 3 },
    { L"標準描画", L"X", L"座標X", 4 },
    { L"標準描画", L"Y", L"座標Y", 4 },
    { L"標準描画", L"Z", L"座標Z", 4 },
    { L"標準描画", L"中心X", L"中心X", 4 },
    { L"標準描画", L"中心Y", L"中心Y", 4 },
    { L"標準描画", L"中心Z", L"中心Z", 4 },
    { L"標準描画", L"X軸回転", L"X軸回転", 5 },
    { L"標準描画", L"Y軸回転", L"Y軸回転", 5 },
    { L"標準描画", L"Z軸回転", L"Z軸回転", 5 },
    { L"標準描画", L"拡大率", L"拡大率", 5 },
    { L"標準描画", L"縦横比", L"縦横比", 5 },
    { L"標準描画", L"透明度", L"透明度", 6 },
    { L"標準描画", L"合成モード", L"合成モード", 6 },
};

// ============================================================
// 前方宣言
// ============================================================
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
static LRESULT CALLBACK AddDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
static LRESULT CALLBACK TextStyleQuickDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
static int GetObjectLayerIndex(void* obj);
static bool SafeReadIntAtOffset(void* obj, int offset, int& outValue);
static bool SafeReadUIntPtrAtOffset(void* obj, int offset, uintptr_t& outValue);
static std::string WideToUtf8(const std::wstring& w);
void OpenAddDialog(const std::string& data, bool isEdit);
int ShowDarkMsg(HWND parent, LPCWSTR text, LPCWSTR title, UINT type);
void RefreshAssets(bool reloadFav);
void UpdateDisplayList();
void CreatePluginWindow();
static void AppendGroupDebugLine(const std::string& line);
static void OpenTextStyleQuickDialog(HWND parent);
static void BroadcastTextStyleConfigChanged();
static std::wstring ColorToHex(COLORREF c);
static bool TryParseHexColor(const std::wstring& text, COLORREF& out);
static void RecreateUiBrushes();
static void ResetThemeColorsToDefault();
static RECT GetAddPreviewRect();
static void UpdateAddCapturePresetUi(HWND hwnd);

// ============================================================
// 1. ユーティリティ & 設定ファイル処理
// ============================================================
static std::wstring ToLower(const std::wstring& s) {
    std::wstring ret = s;
    std::transform(ret.begin(), ret.end(), ret.begin(), ::towlower);
    return ret;
}

static std::wstring ColorToHex(COLORREF c) {
    wchar_t buf[16] = {};
    swprintf_s(buf, L"#%02X%02X%02X", GetRValue(c), GetGValue(c), GetBValue(c));
    return std::wstring(buf);
}

static bool TryParseHexColor(const std::wstring& text, COLORREF& out) {
    std::wstring t = text;
    t.erase(std::remove_if(t.begin(), t.end(), [](wchar_t ch) { return iswspace(ch) != 0; }), t.end());
    if (t.empty()) return false;
    if (t[0] == L'#') t.erase(t.begin());
    if (t.size() != 6) return false;
    for (wchar_t ch : t) {
        if (!iswxdigit(ch)) return false;
    }
    unsigned int v = 0;
    if (swscanf_s(t.c_str(), L"%x", &v) != 1) return false;
    out = RGB((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
    return true;
}

static std::wstring CaptureRectToConfigText(const RECT& rc) {
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    return std::to_wstring(rc.left) + L"," + std::to_wstring(rc.top) + L"," + std::to_wstring(w) + L"," + std::to_wstring(h);
}

static bool ParseCaptureRectText(const std::wstring& text, RECT& outRc) {
    int x = 0, y = 0, w = 0, h = 0;
    if (swscanf_s(text.c_str(), L"%d,%d,%d,%d", &x, &y, &w, &h) != 4) return false;
    if (w <= 0 || h <= 0) return false;
    outRc.left = x;
    outRc.top = y;
    outRc.right = x + w;
    outRc.bottom = y + h;
    return true;
}

static bool ClampRectToVirtualScreen(const RECT& src, RECT& out) {
    int vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
    int vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
    int vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    int vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    RECT vr = { vx, vy, vx + vw, vy + vh };
    RECT rc = src;
    if (rc.left < vr.left) {
        int d = vr.left - rc.left;
        rc.left += d; rc.right += d;
    }
    if (rc.top < vr.top) {
        int d = vr.top - rc.top;
        rc.top += d; rc.bottom += d;
    }
    if (rc.right > vr.right) {
        int d = rc.right - vr.right;
        rc.left -= d; rc.right -= d;
    }
    if (rc.bottom > vr.bottom) {
        int d = rc.bottom - vr.bottom;
        rc.top -= d; rc.bottom -= d;
    }
    if (rc.left < vr.left) rc.left = vr.left;
    if (rc.top < vr.top) rc.top = vr.top;
    if (rc.right > vr.right) rc.right = vr.right;
    if (rc.bottom > vr.bottom) rc.bottom = vr.bottom;
    if (rc.right - rc.left <= 0 || rc.bottom - rc.top <= 0) return false;
    out = rc;
    return true;
}

static void RecreateUiBrushes() {
    if (g_hBrInputBg) { DeleteObject(g_hBrInputBg); g_hBrInputBg = nullptr; }
    if (g_hBrBg) { DeleteObject(g_hBrBg); g_hBrBg = nullptr; }
    g_hBrInputBg = CreateSolidBrush(COL_INPUT_BG);
    g_hBrBg = CreateSolidBrush(COL_BG);
}

static void ResetThemeColorsToDefault() {
    for (int i = 0; i < kThemeColorCount; ++i) {
        *kThemeColors[i].value = kThemeColors[i].defValue;
    }
}

static int GetMaxScrollY(HWND hwnd) {
    RECT rc = {0};
    GetClientRect(hwnd, &rc);
    int clientW = (int)(rc.right - rc.left);
    int clientH = (int)(rc.bottom - rc.top);
    int viewH = clientH - GetListTopY() - GetBottomReservedHeight();
    if (viewH <= 0) return 0;

    int itemH = GetListItemHeight();
    int cols = (std::max)(1, clientW / GetListItemMinWidth());
    int rows = ((int)g_displayAssets.size() + cols - 1) / cols;
    int contentH = rows * itemH;
    if (rows > 0) contentH -= 2;
    return (std::max)(0, contentH - viewH);
}

static void UpdateScrollBar(HWND hwnd) {
    if (!hwnd) return;

    RECT rc = {0};
    GetClientRect(hwnd, &rc);
    int clientH = (int)(rc.bottom - rc.top);
    int viewH = clientH - GetListTopY() - GetBottomReservedHeight();
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

static std::string TrimAscii(const std::string& s) {
    size_t b = 0;
    while (b < s.size() && (s[b] == ' ' || s[b] == '\t')) b++;
    size_t e = s.size();
    while (e > b && (s[e - 1] == ' ' || s[e - 1] == '\t')) e--;
    return s.substr(b, e - b);
}

static bool StartsWithNoCase(const std::string& s, const std::string& prefix) {
    if (s.size() < prefix.size()) return false;
    for (size_t i = 0; i < prefix.size(); ++i) {
        unsigned char a = (unsigned char)s[i];
        unsigned char b = (unsigned char)prefix[i];
        if (std::tolower(a) != std::tolower(b)) return false;
    }
    return true;
}

static int ExtractHeaderIntValue(const std::string& src, const std::string& key) {
    std::stringstream ss(src);
    std::string line;
    bool inTopHeader = false;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        std::string t = TrimAscii(line);
        if (t.empty()) continue;

        if (!inTopHeader) {
            if (t.front() == '[' && t.back() == ']') {
                std::string sec = t.substr(1, t.size() - 2);
                if (sec.find('.') == std::string::npos) {
                    inTopHeader = true;
                }
            }
            continue;
        }

        if (t.front() == '[' && t.back() == ']') break;
        if (StartsWithNoCase(t, key)) {
            std::string v = TrimAscii(t.substr(key.length()));
            try { return std::stoi(v); }
            catch (...) { return 0; }
        }
    }
    return 0;
}

// aliasテキスト内の frame= 行を走査して、全体の最小開始/最大終了を返す。

static bool ExtractHeaderFrameRange(const std::string& src, int& outStart, int& outEnd) {
    auto parseFramePair = [](const std::string& line, int& s, int& e) -> bool {
        std::string t = TrimAscii(line);
        if (t.empty()) return false;
        if (!StartsWithNoCase(t, "frame")) return false;

        size_t i = 5;
        while (i < t.size() && isspace((unsigned char)t[i])) i++;
        if (i >= t.size() || t[i] != '=') return false;
        i++;
        while (i < t.size() && isspace((unsigned char)t[i])) i++;

        // frame= a,b の2値だけでなく、余分な値があっても先頭/末尾を使う。
        std::vector<int> vals;
        while (i < t.size()) {
            while (i < t.size() && (isspace((unsigned char)t[i]) || t[i] == ',')) i++;
            if (i >= t.size()) break;

            size_t b = i;
            if (t[i] == '+' || t[i] == '-') i++;
            size_t d = i;
            while (i < t.size() && isdigit((unsigned char)t[i])) i++;
            if (d == i) break;

            try {
                vals.push_back(std::stoi(t.substr(b, i - b)));
            } catch (...) {
                return false;
            }
        }

        if (vals.size() < 2) return false;
        s = vals.front();
        e = vals.back();
        return true;
    };

    std::stringstream ss(src);
    std::string line;
    int start = INT_MAX;
    int end = INT_MIN;
    bool found = false;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        std::string t = TrimAscii(line);
        if (t.empty()) continue;

        int s = 0, e = 0;
        if (parseFramePair(t, s, e)) {
            start = (std::min)(start, s);
            end = (std::max)(end, e);
            found = true;
        }
    }

    if (found) {
        outStart = start;
        outEnd = end;
        return true;
    }
    return false;
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return L"";
    int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    if (len <= 0) return L"";
    std::wstring w((size_t)len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &w[0], len);
    if (!w.empty() && w.back() == L'\0') w.pop_back();
    return w;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::string s((size_t)len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], len, nullptr, nullptr);
    if (!s.empty() && s.back() == '\0') s.pop_back();
    return s;
}

struct ObjectAliasSection {
    std::string header;
    std::vector<std::string> lines;
};

static bool IsObjectRootHeader(const std::string& header) {
    return _stricmp(header.c_str(), "[Object]") == 0;
}

static bool TryParseObjectSectionIndex(const std::string& header, int& outIndex) {
    outIndex = -1;
    if (header.size() < 10) return false;
    if (_strnicmp(header.c_str(), "[Object.", 8) != 0) return false;
    if (header.back() != ']') return false;
    std::string num = header.substr(8, header.size() - 9);
    if (num.empty()) return false;
    for (char c : num) if (!isdigit((unsigned char)c)) return false;
    try {
        outIndex = std::stoi(num);
        return true;
    } catch (...) {
        return false;
    }
}

static bool ParseObjectAliasSections(const std::string& src, std::vector<ObjectAliasSection>& outSections) {
    outSections.clear();
    std::stringstream ss(src);
    std::string line;
    ObjectAliasSection cur = {};
    bool hasSection = false;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        std::string t = TrimAscii(line);
        bool isHeader = (t.size() >= 2 && t.front() == '[' && t.back() == ']');
        if (isHeader) {
            if (hasSection) outSections.push_back(cur);
            cur = {};
            cur.header = t;
            hasSection = true;
        } else if (hasSection) {
            cur.lines.push_back(line);
        }
    }
    if (hasSection) outSections.push_back(cur);
    return !outSections.empty();
}

static std::string SerializeObjectAliasSections(const std::vector<ObjectAliasSection>& sections) {
    std::string out;
    for (size_t i = 0; i < sections.size(); ++i) {
        if (!out.empty()) out += "\r\n";
        out += sections[i].header;
        out += "\r\n";
        for (const auto& ln : sections[i].lines) {
            out += ln;
            out += "\r\n";
        }
    }
    return out;
}

static bool HasMyAssetMarker(const ObjectAliasSection& sec) {
    for (const auto& ln : sec.lines) {
        if (TrimAscii(ln) == "#MyAsset") return true;
    }
    return false;
}

static void RemoveMyAssetMarkerLines(ObjectAliasSection& sec) {
    std::vector<std::string> filtered;
    filtered.reserve(sec.lines.size());
    for (const auto& ln : sec.lines) {
        if (TrimAscii(ln) == "#MyAsset") continue;
        filtered.push_back(ln);
    }
    sec.lines.swap(filtered);
}

static void RenumberObjectSections(std::vector<ObjectAliasSection>& sections) {
    int idx = 0;
    for (auto& sec : sections) {
        if (IsObjectRootHeader(sec.header)) continue;
        int dummy = -1;
        if (TryParseObjectSectionIndex(sec.header, dummy)) {
            sec.header = "[Object." + std::to_string(idx++) + "]";
        }
    }
}

static bool RemoveMarkedStyleSections(std::string& data) {
    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(data, sections)) return false;
    std::vector<ObjectAliasSection> kept;
    kept.reserve(sections.size());
    bool changed = false;
    for (auto& sec : sections) {
        if (HasMyAssetMarker(sec)) {
            changed = true;
            continue;
        }
        kept.push_back(sec);
    }
    if (!changed) return false;
    RenumberObjectSections(kept);
    data = SerializeObjectAliasSections(kept);
    return true;
}

static bool RemoveOnlyStyleMarkers(std::string& data) {
    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(data, sections)) return false;
    bool changed = false;
    for (auto& sec : sections) {
        size_t before = sec.lines.size();
        RemoveMyAssetMarkerLines(sec);
        if (sec.lines.size() != before) changed = true;
    }
    if (!changed) return false;
    data = SerializeObjectAliasSections(sections);
    return true;
}

static bool IsSingleTextObjectAlias(const std::string& alias) {
    if (alias.empty()) return false;
    if (alias.find("[0]") != std::string::npos) return false;

    std::string lower = alias;
    std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char c) { return (char)std::tolower(c); });
    if (lower.find("text=") != std::string::npos) return true;

    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(alias, sections)) return false;
    bool hasRoot = false;
    bool hasText = false;
    for (const auto& sec : sections) {
        if (IsObjectRootHeader(sec.header)) {
            hasRoot = true;
            for (const auto& ln : sec.lines) {
                if (StartsWithNoCase(TrimAscii(ln), "text=")) {
                    hasText = true;
                    break;
                }
            }
        }
    }
    return hasRoot && hasText;
}

static bool IsLikelySingleObjectAlias(const std::string& alias) {
    if (alias.empty()) return false;
    if (alias.find("[0]") != std::string::npos) return false; // multi-object形式は除外
    if (alias.find("[Object]") != std::string::npos) return true;
    if (alias.find("[Object.") != std::string::npos) return true;
    return false;
}

static std::string GetSectionEffectName(const ObjectAliasSection& sec);
static ObjectAliasSection* FindFirstSectionByEffect(std::vector<ObjectAliasSection>& sections, const std::string& effectName);

static bool ExtractStyleSectionsFromAssetAlias(const std::string& alias, std::vector<ObjectAliasSection>& outStyleSections) {
    outStyleSections.clear();
    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(alias, sections)) return false;
    for (auto sec : sections) {
        int idx = -1;
        if (!TryParseObjectSectionIndex(sec.header, idx)) continue;
        if (idx < 0) continue;
        RemoveMyAssetMarkerLines(sec);
        outStyleSections.push_back(sec);
    }
    return !outStyleSections.empty();
}

static bool InsertNewStyleSections(std::string& data, const std::vector<ObjectAliasSection>& styleSections) {
    if (styleSections.empty()) return false;
    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(data, sections)) return false;

    size_t insertPos = sections.size();
    if (g_textStyleInsertMode == 1) {
        for (size_t i = 0; i < sections.size(); ++i) {
            if (GetSectionEffectName(sections[i]) == WideToUtf8(L"標準描画")) {
                insertPos = i + 1;
                break;
            }
        }
    }
    sections.insert(sections.begin() + (std::min)(insertPos, sections.size()), styleSections.begin(), styleSections.end());
    RenumberObjectSections(sections);
    data = SerializeObjectAliasSections(sections);
    return true;
}

static void RemoveSectionsFromObjectIndex2(std::string& data) {
    std::vector<ObjectAliasSection> sections;
    if (!ParseObjectAliasSections(data, sections)) return;
    std::vector<ObjectAliasSection> kept;
    kept.reserve(sections.size());
    for (const auto& sec : sections) {
        if (IsObjectRootHeader(sec.header)) {
            kept.push_back(sec);
            continue;
        }
        int idx = -1;
        if (TryParseObjectSectionIndex(sec.header, idx) && idx <= 1) {
            kept.push_back(sec);
        }
    }
    RenumberObjectSections(kept);
    data = SerializeObjectAliasSections(kept);
}

static bool BuildAliasWithAppliedTextStyle(const std::string& targetAlias, const std::string& styleAssetAlias, std::string& outAlias) {
    outAlias = targetAlias;
    if (g_textStyleClearEffectsFrom2) {
        RemoveSectionsFromObjectIndex2(outAlias);
    }
    std::vector<ObjectAliasSection> styleSections;
    if (!ExtractStyleSectionsFromAssetAlias(styleAssetAlias, styleSections)) return false;
    return InsertNewStyleSections(outAlias, styleSections);
}

static std::string GetSectionEffectName(const ObjectAliasSection& sec) {
    for (const auto& ln : sec.lines) {
        std::string t = TrimAscii(ln);
        if (!StartsWithNoCase(t, "effect.name=")) continue;
        return t.substr(12);
    }
    return "";
}

static bool TryGetSectionItemValue(const ObjectAliasSection& sec, const std::string& key, std::string& outVal) {
    for (const auto& ln : sec.lines) {
        std::string t = TrimAscii(ln);
        size_t eq = t.find('=');
        if (eq == std::string::npos) continue;
        if (TrimAscii(t.substr(0, eq)) != key) continue;
        outVal = t.substr(eq + 1);
        return true;
    }
    return false;
}

static void SetSectionItemValue(ObjectAliasSection& sec, const std::string& key, const std::string& val) {
    for (auto& ln : sec.lines) {
        std::string t = TrimAscii(ln);
        size_t eq = t.find('=');
        if (eq == std::string::npos) continue;
        if (TrimAscii(t.substr(0, eq)) != key) continue;
        ln = key + "=" + val;
        return;
    }
    sec.lines.push_back(key + "=" + val);
}

static ObjectAliasSection* FindFirstSectionByEffect(std::vector<ObjectAliasSection>& sections, const std::string& effectName) {
    for (auto& sec : sections) {
        if (GetSectionEffectName(sec) == effectName) return &sec;
    }
    return nullptr;
}

static void ApplyTextStyleInheritanceItems(const std::string& originalAlias, std::string& styledAlias) {
    std::vector<ObjectAliasSection> srcSections;
    std::vector<ObjectAliasSection> dstSections;
    if (!ParseObjectAliasSections(originalAlias, srcSections)) return;
    if (!ParseObjectAliasSections(styledAlias, dstSections)) return;

    for (int i = 0; i < kInheritItemCount; ++i) {
        if (!g_inheritItemEnabled[i]) continue;
        std::string eff = WideToUtf8(kInheritItems[i].effect);
        std::string key = WideToUtf8(kInheritItems[i].key);
        if (eff.empty() || key.empty()) continue;

        ObjectAliasSection* srcSec = FindFirstSectionByEffect(srcSections, eff);
        if (!srcSec) continue;

        std::string v;
        if (!TryGetSectionItemValue(*srcSec, key, v)) continue;
        bool foundFirst = false;
        for (auto& dstSec : dstSections) {
            if (GetSectionEffectName(dstSec) != eff) continue;
            if (!foundFirst) { foundFirst = true; continue; }
            SetSectionItemValue(dstSec, key, v);
        }
    }
    styledAlias = SerializeObjectAliasSections(dstSections);
}

struct TextStyleApplyContext {
    std::string styleAlias;
    std::string originalAlias;
    int appliedCount = 0;
    bool success = false;
    bool hasError = false;
    std::wstring errorMessage;
    int pendingLayer = -1;
    int pendingStart = -1;
    int pendingEnd = -1;
};


static int GetVerifiedTimelineGroupId(void* obj) {
    if (!obj) return 0;

    auto isLikelyGroupToken = [](uintptr_t v) -> bool {
        if (v == 0) return false;
        if (v == (uintptr_t)~(uintptr_t)0) return false;
#if INTPTR_MAX == INT64_MAX
        if (v == 0xFFFFFFFFFFFFFFFEULL) return false;
#else
        if (v == 0xFFFFFFFEUL) return false;
#endif
        if (v < 0x10000) return false;
        return true;
    };

    const int kGroupPtrOffsets[] = { 0x30, 0x38 };
    for (int off : kGroupPtrOffsets) {
        uintptr_t token = 0;
        if (!SafeReadUIntPtrAtOffset(obj, off, token)) continue;
        if (!isLikelyGroupToken(token)) continue;
        unsigned long long mixed = (unsigned long long)token;
        mixed ^= (mixed >> 33);
        mixed ^= (mixed >> 17);
        int gid = (int)(mixed & 0x7FFFFFFF);
        if (gid == 0) gid = 1;
        return gid;
    }

    return 0;
}

static void DumpApiItemValuesFromAlias(void* editSafe, void* obj, const std::string& alias, int idx) {
    if (!editSafe || !obj) return;
    if (alias.empty()) return;

    auto* edit = (EDIT_SECTION_SAFE*)editSafe;
    using FN_GET_OBJECT_ITEM_VALUE = LPCSTR (*)(void*, LPCWSTR, LPCWSTR);
    FN_GET_OBJECT_ITEM_VALUE fnGetValue = (FN_GET_OBJECT_ITEM_VALUE)edit->f6;
    if (!fnGetValue) {
        AppendGroupDebugLine("  ---- api item dump skipped (get_object_item_value unavailable) ----");
        return;
    }

    std::vector<std::pair<std::wstring, std::wstring>> effectItems;
    std::set<std::string> seen;
    std::vector<std::wstring> effectNames;

    std::stringstream ss(alias);
    std::string line;
    bool inObjectSection = false;
    std::wstring currentEffect;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        std::string t = TrimAscii(line);
        if (t.empty()) continue;

        if (t.size() >= 2 && t.front() == '[' && t.back() == ']') {
            inObjectSection = (t.rfind("[Object.", 0) == 0);
            currentEffect.clear();
            continue;
        }
        if (!inObjectSection) continue;

        size_t eq = t.find('=');
        if (eq == std::string::npos) continue;
        std::string keyUtf8 = TrimAscii(t.substr(0, eq));
        std::string valUtf8 = TrimAscii(t.substr(eq + 1));
        if (keyUtf8.empty()) continue;

        if (_stricmp(keyUtf8.c_str(), "effect.name") == 0) {
            currentEffect = Utf8ToWide(valUtf8);
            if (!currentEffect.empty()) effectNames.push_back(currentEffect);
            continue;
        }

        if (currentEffect.empty()) continue;
        std::wstring itemW = Utf8ToWide(keyUtf8);
        if (itemW.empty()) continue;

        std::string dedup = WideToUtf8(currentEffect) + "|" + WideToUtf8(itemW);
        if (seen.insert(dedup).second) {
            effectItems.push_back({currentEffect, itemW});
        }
    }

    {
        std::ostringstream os;
        os << "  ---- api item dump begin idx=" << idx << " count=" << effectItems.size() << " ----";
        AppendGroupDebugLine(os.str());
    }

    for (const auto& ei : effectItems) {
        LPCSTR v = fnGetValue(obj, ei.first.c_str(), ei.second.c_str());
        std::ostringstream os;
        os << "    effect=" << WideToUtf8(ei.first)
           << " item=" << WideToUtf8(ei.second)
           << " value=" << (v ? v : "(null)");
        AppendGroupDebugLine(os.str());
    }


    static const wchar_t* kProbeItems[] = {
        L"group", L"Group", L"グループ", L"group_id", L"GroupId", L"グループID", L"対象グループ"
    };
    std::set<std::wstring> uniqueEffects(effectNames.begin(), effectNames.end());
    for (const auto& eff : uniqueEffects) {
        for (const wchar_t* probeItem : kProbeItems) {
            LPCSTR v = fnGetValue(obj, eff.c_str(), probeItem);
            if (!v || !*v) continue;
            std::ostringstream os;
            os << "    [probe-hit] effect=" << WideToUtf8(eff)
               << " item=" << WideToUtf8(probeItem)
               << " value=" << v;
            AppendGroupDebugLine(os.str());
        }
    }
    AppendGroupDebugLine("  ---- api item dump end ----");
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

static int GetObjectLayerIndex(void* obj, int* outOffset) {
    if (outOffset) *outOffset = 0;
    if (!obj) return 0;

    const int kLayerOffsets[] = { 0x60, 0x68, 0x118 };
    for (int off : kLayerOffsets) {
        int v = 0;
        if (!SafeReadIntAtOffset(obj, off, v)) continue;
        if (v > 0 && v < 1000000) {
            if (outOffset) *outOffset = off;
            return v;
        }
    }
    return 0;
}

static int GetObjectLayerIndex(void* obj) {
    return GetObjectLayerIndex(obj, nullptr);
}

static int ResolveObjectLayerForAssetBuild(void* obj, LPCSTR rawAlias, const char** outSource = nullptr) {
    if (outSource) *outSource = "fallback";
    if (rawAlias) {
        int aliasLayer = ExtractHeaderIntValue(rawAlias, "layer=");
        if (aliasLayer > 0) {
            if (outSource) *outSource = "alias";
            return aliasLayer;
        }
    }
    int memOff = 0;
    int memLayer = GetObjectLayerIndex(obj, &memOff);
    if (memLayer > 0 && memLayer < 1000000) {
        if (outSource) {
            if (memOff == 0x60) *outSource = "mem+0x60";
            else if (memOff == 0x68) *outSource = "mem+0x68";
            else if (memOff == 0x118) *outSource = "mem+0x118";
            else *outSource = "mem+offset";
        }
        return memLayer;
    }
    if (outSource) *outSource = "fallback=1";
    return 1;
}

static bool TryGetObjectFrameRangeFromMemory(void* obj, int& outStart, int& outEnd) {
    int s = 0, e = 0;
    if (!SafeReadIntAtOffset(obj, 0x40, s)) return false;
    if (!SafeReadIntAtOffset(obj, 0x44, e)) return false;
    if (e < s) return false;
    if (s < -1000000 || e > 10000000) return false;
    outStart = s;
    outEnd = e;
    return true;
}

static bool SafeReadIntAtOffset(void* obj, int offset, int& outValue) {
    if (!obj || offset < 0) return false;
#if defined(_MSC_VER)
    __try {
        outValue = *(int*)((char*)obj + offset);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
#else
    outValue = *(int*)((char*)obj + offset);
    return true;
#endif
}

static bool IsReadableMemoryRange(const unsigned char* p, size_t size);

static bool SafeReadUIntPtrAtOffset(void* obj, int offset, uintptr_t& outValue) {
    outValue = 0;
    if (!obj || offset < 0) return false;
    const unsigned char* p = (const unsigned char*)obj + offset;
    if (!IsReadableMemoryRange(p, sizeof(uintptr_t))) return false;
    memcpy(&outValue, p, sizeof(uintptr_t));
    return true;
}

static bool IsReadableMemoryRange(const unsigned char* p, size_t size) {
    size_t checked = 0;
    while (checked < size) {
        MEMORY_BASIC_INFORMATION mbi = {};
        if (VirtualQuery(p + checked, &mbi, sizeof(mbi)) == 0) return false;
        if (mbi.State != MEM_COMMIT) return false;
        if (mbi.Protect & (PAGE_NOACCESS | PAGE_GUARD)) return false;
        size_t regionRemain = (size_t)((const unsigned char*)mbi.BaseAddress + mbi.RegionSize - (p + checked));
        if (regionRemain == 0) return false;
        checked += (std::min)(regionRemain, size - checked);
    }
    return true;
}

static void DumpObjectMemoryWide(void* obj, int idx) {
    if (!obj) return;

    constexpr int kDumpStart = 0x000;
    constexpr int kDumpSize  = 0x800; // 2KB
    const unsigned char* base = (const unsigned char*)obj;

    {
        std::ostringstream os;
        os << "  ---- raw memory dump begin idx=" << idx
           << " obj=0x" << std::hex << (uintptr_t)obj << std::dec
           << " range=+0x" << std::hex << kDumpStart << "..+0x" << (kDumpStart + kDumpSize - 1) << std::dec
           << " ----";
        AppendGroupDebugLine(os.str());
    }

    if (!IsReadableMemoryRange(base + kDumpStart, (size_t)kDumpSize)) {
        AppendGroupDebugLine("  raw memory dump skipped: unreadable range");
        AppendGroupDebugLine("  ---- raw memory dump end ----");
        return;
    }

    std::vector<unsigned char> buf((size_t)kDumpSize, 0);
    memcpy(buf.data(), base + kDumpStart, (size_t)kDumpSize);

    for (int off = 0; off < kDumpSize; off += 16) {
        std::ostringstream os;
        os << "    +0x" << std::hex << (kDumpStart + off) << ":";
        for (int i = 0; i < 16; ++i) {
            os << " " << std::setw(2) << std::setfill('0') << std::hex << (int)buf[(size_t)off + (size_t)i];
        }

        os << "  | dword:";
        for (int i = 0; i < 16; i += 4) {
            int v = 0;
            memcpy(&v, &buf[(size_t)off + (size_t)i], sizeof(int));
            os << " " << std::dec << v;
        }
        AppendGroupDebugLine(os.str());
    }

    AppendGroupDebugLine("  ---- raw memory dump end ----");
}

static void DumpGroupMemoryProbe(const std::vector<void*>& selectedObjs) {
    if (selectedObjs.empty()) return;
    AppendGroupDebugLine("  ---- memory probe begin ----");
    AppendGroupDebugLine("  probe range: offset 0x20..0x140 step 4");

    for (int off = 0x20; off <= 0x140; off += 4) {
        std::map<int, int> freq;
        int validCount = 0;
        bool hasBad = false;
        std::ostringstream vals;
        vals << "  off=0x" << std::hex << off << std::dec << " values=";

        for (size_t i = 0; i < selectedObjs.size(); ++i) {
            int v = 0;
            if (!SafeReadIntAtOffset(selectedObjs[i], off, v)) {
                hasBad = true;
                vals << " [i" << i << ":ERR]";
                continue;
            }
            validCount++;
            freq[v]++;
            vals << " [i" << i << ":" << v << "]";
        }

        if (validCount <= 0) continue;

        int uniqueCount = (int)freq.size();
        int maxBucket = 0;
        for (const auto& kv : freq) if (kv.second > maxBucket) maxBucket = kv.second;

        bool looksGrouped =
            (uniqueCount >= 1 && uniqueCount <= (int)selectedObjs.size() / 2 + 1) &&
            (maxBucket >= 2);

        if (looksGrouped || off == 0x44 || off == 0x64) {
            std::ostringstream head;
            head << "  probe off=0x" << std::hex << off << std::dec
                 << " unique=" << uniqueCount
                 << " maxBucket=" << maxBucket
                 << " valid=" << validCount
                 << " readErr=" << (hasBad ? 1 : 0);
            AppendGroupDebugLine(head.str());
            AppendGroupDebugLine(vals.str());
        }
    }
    AppendGroupDebugLine("  ---- memory probe end ----");
}

static void DumpGroupIdCandidateProbe(const std::vector<void*>& selectedObjs) {
    if (selectedObjs.size() < 2) return;
    auto isSentinelInt = [](int v) -> bool {
        return v == 0 || v == -1 || v == -2;
    };
    auto isSentinelPtr = [](uintptr_t v) -> bool {
        if (v == 0) return true;
        if (v == (uintptr_t)~(uintptr_t)0) return true;
#if INTPTR_MAX == INT64_MAX
        if (v == 0xFFFFFFFFFFFFFFFEULL) return true;
#else
        if (v == 0xFFFFFFFEUL) return true;
#endif
        return false;
    };

    struct IntCand {
        int off = 0;
        int uniqueCount = 0;
        int maxBucket = 0;
        int validCount = 0;
        int zeroCount = 0;
        int score = 0;
        std::string vals;
    };
    std::vector<IntCand> intCands;

    for (int off = 0x20; off <= 0x240; off += 4) {
        std::map<int, int> freq;
        std::map<int, int> freqNonSentinel;
        int validCount = 0;
        int zeroCount = 0;
        int nonSentinelCount = 0;
        std::ostringstream vals;
        vals << "    int off=0x" << std::hex << off << std::dec << " values=";

        for (size_t i = 0; i < selectedObjs.size(); ++i) {
            int v = 0;
            if (!SafeReadIntAtOffset(selectedObjs[i], off, v)) {
                vals << " [i" << i << ":ERR]";
                continue;
            }
            validCount++;
            if (v == 0) zeroCount++;
            freq[v]++;
            if (!isSentinelInt(v)) {
                freqNonSentinel[v]++;
                nonSentinelCount++;
            }
            vals << " [i" << i << ":" << v << "]";
        }
        if (validCount < 2) continue;

        int uniqueCount = (int)freq.size();
        int maxBucket = 0;
        for (const auto& kv : freq) if (kv.second > maxBucket) maxBucket = kv.second;
        int uniqueNonSentinel = (int)freqNonSentinel.size();
        int maxBucketNonSentinel = 0;
        int repeatedBuckets = 0;
        for (const auto& kv : freqNonSentinel) if (kv.second > maxBucketNonSentinel) maxBucketNonSentinel = kv.second;
        for (const auto& kv : freqNonSentinel) if (kv.second >= 2) repeatedBuckets++;

        bool candidate =
            (uniqueNonSentinel >= 2) &&
            (uniqueNonSentinel < validCount) &&
            (maxBucketNonSentinel >= 2) &&
            (repeatedBuckets >= 1);
        if (!candidate) continue;

        IntCand c = {};
        c.off = off;
        c.uniqueCount = uniqueCount;
        c.maxBucket = maxBucket;
        c.validCount = validCount;
        c.zeroCount = zeroCount;
        c.score = maxBucketNonSentinel * 140 + repeatedBuckets * 40 + nonSentinelCount * 4 - uniqueNonSentinel * 12 - zeroCount * 3;
        c.vals = vals.str();
        intCands.push_back(c);
    }

    struct PtrCand {
        int off = 0;
        int uniqueCount = 0;
        int maxBucket = 0;
        int validCount = 0;
        int nullCount = 0;
        int score = 0;
        std::string vals;
    };
    std::vector<PtrCand> ptrCands;

    for (int off = 0x20; off <= 0x180; off += (int)sizeof(uintptr_t)) {
        std::map<uintptr_t, int> freq;
        std::map<uintptr_t, int> freqNonSentinel;
        int validCount = 0;
        int nullCount = 0;
        int nonSentinelCount = 0;
        std::ostringstream vals;
        vals << "    ptr off=0x" << std::hex << off << std::dec << " values=";

        for (size_t i = 0; i < selectedObjs.size(); ++i) {
            uintptr_t v = 0;
            if (!SafeReadUIntPtrAtOffset(selectedObjs[i], off, v)) {
                vals << " [i" << i << ":ERR]";
                continue;
            }
            validCount++;
            if (v == 0) nullCount++;
            freq[v]++;
            if (!isSentinelPtr(v)) {
                freqNonSentinel[v]++;
                nonSentinelCount++;
            }
            vals << " [i" << i << ":0x" << std::hex << (unsigned long long)v << std::dec << "]";
        }
        if (validCount < 2) continue;

        int uniqueCount = (int)freq.size();
        int maxBucket = 0;
        for (const auto& kv : freq) if (kv.second > maxBucket) maxBucket = kv.second;
        int uniqueNonSentinel = (int)freqNonSentinel.size();
        int maxBucketNonSentinel = 0;
        int repeatedBuckets = 0;
        for (const auto& kv : freqNonSentinel) if (kv.second > maxBucketNonSentinel) maxBucketNonSentinel = kv.second;
        for (const auto& kv : freqNonSentinel) if (kv.second >= 2) repeatedBuckets++;

        bool candidate =
            (uniqueNonSentinel >= 2) &&
            (uniqueNonSentinel < validCount) &&
            (maxBucketNonSentinel >= 2) &&
            (repeatedBuckets >= 1);
        if (!candidate) continue;

        PtrCand c = {};
        c.off = off;
        c.uniqueCount = uniqueCount;
        c.maxBucket = maxBucket;
        c.validCount = validCount;
        c.nullCount = nullCount;
        c.score = maxBucketNonSentinel * 140 + repeatedBuckets * 40 + nonSentinelCount * 4 - uniqueNonSentinel * 12 - nullCount * 3;
        c.vals = vals.str();
        ptrCands.push_back(c);
    }

    std::sort(intCands.begin(), intCands.end(), [](const IntCand& a, const IntCand& b) {
        if (a.score != b.score) return a.score > b.score;
        return a.off < b.off;
    });
    std::sort(ptrCands.begin(), ptrCands.end(), [](const PtrCand& a, const PtrCand& b) {
        if (a.score != b.score) return a.score > b.score;
        return a.off < b.off;
    });

    AppendGroupDebugLine("  ---- group-id candidate probe begin ----");
    {
        std::ostringstream os;
        os << "  selected=" << selectedObjs.size()
           << " intCandidates=" << intCands.size()
           << " ptrCandidates=" << ptrCands.size();
        AppendGroupDebugLine(os.str());
    }

    int intLimit = (std::min)(20, (int)intCands.size());
    for (int i = 0; i < intLimit; ++i) {
        const auto& c = intCands[(size_t)i];
        std::ostringstream os;
        os << "  [int#" << i
           << "] off=0x" << std::hex << c.off << std::dec
           << " score=" << c.score
           << " unique=" << c.uniqueCount
           << " maxBucket=" << c.maxBucket
           << " valid=" << c.validCount
           << " zero=" << c.zeroCount;
        AppendGroupDebugLine(os.str());
        AppendGroupDebugLine(c.vals);
    }

    int ptrLimit = (std::min)(12, (int)ptrCands.size());
    for (int i = 0; i < ptrLimit; ++i) {
        const auto& c = ptrCands[(size_t)i];
        std::ostringstream os;
        os << "  [ptr#" << i
           << "] off=0x" << std::hex << c.off << std::dec
           << " score=" << c.score
           << " unique=" << c.uniqueCount
           << " maxBucket=" << c.maxBucket
           << " valid=" << c.validCount
           << " null=" << c.nullCount;
        AppendGroupDebugLine(os.str());
        AppendGroupDebugLine(c.vals);
    }
    AppendGroupDebugLine("  ---- group-id candidate probe end ----");
}

static void DumpLayerOffsetProbe(const std::vector<void*>& selectedObjs) {
    if (selectedObjs.empty()) return;
    AppendGroupDebugLine("  ---- layer offset probe begin ----");
    AppendGroupDebugLine("  candidates: 0x60, 0x68, 0x118");
    for (size_t i = 0; i < selectedObjs.size(); ++i) {
        void* obj = selectedObjs[i];
        int v60 = 0, v68 = 0, v118 = 0;
        bool ok60 = SafeReadIntAtOffset(obj, 0x60, v60);
        bool ok68 = SafeReadIntAtOffset(obj, 0x68, v68);
        bool ok118 = SafeReadIntAtOffset(obj, 0x118, v118);
        int chosenOff = 0;
        int chosen = GetObjectLayerIndex(obj, &chosenOff);

        std::ostringstream os;
        os << "  layer-probe idx=" << i
           << " chosen=" << chosen
           << " chosenOff=0x" << std::hex << chosenOff << std::dec
           << " v60=" << (ok60 ? std::to_string(v60) : std::string("ERR"))
           << " v68=" << (ok68 ? std::to_string(v68) : std::string("ERR"))
           << " v118=" << (ok118 ? std::to_string(v118) : std::string("ERR"));
        AppendGroupDebugLine(os.str());
    }
    AppendGroupDebugLine("  ---- layer offset probe end ----");
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
static std::wstring GetTextStyleFilePath() {
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetTextStyle.txt");
    return path;
}
static std::wstring GetMetaFilePath() {
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetMeta.json");
    return path;
}
static std::wstring GetConfigPath() { 
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH); 
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetConfig.ini"); 
    return path; 
}
static std::wstring GetGroupDebugPath() {
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetGroupDebug.txt");
    return path;
}
static std::wstring GetUiDebugPath() {
    wchar_t path[MAX_PATH]; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAssetUiDebug.txt");
    return path;
}
static void AppendUiDebugLine(const std::string& line) {
    if (!g_enableMainWindowTrace) return;
    HANDLE hFile = CreateFileW(GetUiDebugPath().c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    SYSTEMTIME st = {};
    GetLocalTime(&st);
    std::ostringstream os;
    os << "[" << st.wYear << "-" << st.wMonth << "-" << st.wDay
       << " " << st.wHour << ":" << st.wMinute << ":" << st.wSecond << "." << st.wMilliseconds << "] "
       << line << "\r\n";
    std::string out = os.str();
    DWORD written = 0;
    WriteFile(hFile, out.c_str(), (DWORD)out.size(), &written, NULL);
    CloseHandle(hFile);
}

struct UiWinDumpContext {
    DWORD pid;
    std::string tag;
    int count;
};

static BOOL CALLBACK EnumWindowsForUiDumpProc(HWND hwnd, LPARAM lp) {
    auto* ctx = (UiWinDumpContext*)lp;
    if (!ctx) return TRUE;

    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid != ctx->pid) return TRUE;

    wchar_t clsW[256] = {};
    wchar_t titleW[256] = {};
    GetClassNameW(hwnd, clsW, 255);
    GetWindowTextW(hwnd, titleW, 255);

    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    LONG_PTR exStyle = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    HWND owner = GetWindow(hwnd, GW_OWNER);
    HWND rootOwner = GetAncestor(hwnd, GA_ROOTOWNER);

    std::ostringstream os;
    os << "WinDump[" << ctx->tag << "]"
       << " hwnd=" << (void*)hwnd
       << " vis=" << (IsWindowVisible(hwnd) ? 1 : 0)
       << " en=" << (IsWindowEnabled(hwnd) ? 1 : 0)
       << " owner=" << (void*)owner
       << " rootOwner=" << (void*)rootOwner
       << " style=0x" << std::hex << (unsigned long long)style
       << " ex=0x" << (unsigned long long)exStyle << std::dec
       << " class=" << WideToUtf8(clsW)
       << " title=" << WideToUtf8(titleW);
    AppendUiDebugLine(os.str());
    ctx->count++;
    return TRUE;
}

static void DumpProcessWindowsForUiDebug(const char* tag) {
    if (!g_enableMainWindowTrace) return;
    UiWinDumpContext ctx = {};
    ctx.pid = GetCurrentProcessId();
    ctx.tag = tag ? tag : "";
    ctx.count = 0;
    AppendUiDebugLine("WinDump begin: " + ctx.tag);
    EnumWindows(EnumWindowsForUiDumpProc, (LPARAM)&ctx);
    std::ostringstream os;
    os << "WinDump end: " << ctx.tag << " count=" << ctx.count;
    AppendUiDebugLine(os.str());
}

struct HostMainFindContext {
    DWORD pid = 0;
    HWND found = nullptr;
};

static BOOL CALLBACK EnumWindowsFindHostMainProc(HWND hwnd, LPARAM lp) {
    auto* ctx = (HostMainFindContext*)lp;
    if (!ctx) return TRUE;
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid != ctx->pid) return TRUE;
    if (!IsWindowVisible(hwnd)) return TRUE;

    wchar_t clsW[256] = {};
    GetClassNameW(hwnd, clsW, 255);
    if (wcsncmp(clsW, L"MyAsset", 7) == 0) return TRUE;

    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    if (!(style & WS_CAPTION)) return TRUE;

    HWND owner = GetWindow(hwnd, GW_OWNER);
    if (owner && IsWindow(owner)) return TRUE;

    ctx->found = hwnd;
    return FALSE;
}

static HWND FindHostMainWindow() {
    if (g_hostMainWnd && IsWindow(g_hostMainWnd)) return g_hostMainWnd;
    HostMainFindContext ctx = {};
    ctx.pid = GetCurrentProcessId();
    EnumWindows(EnumWindowsFindHostMainProc, (LPARAM)&ctx);
    g_hostMainWnd = ctx.found;
    return g_hostMainWnd;
}

static void AppendGroupDebugText(const std::string& text) {
    HANDLE hFile = CreateFileW(GetGroupDebugPath().c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    DWORD written = 0;
    WriteFile(hFile, text.c_str(), (DWORD)text.size(), &written, NULL);
    CloseHandle(hFile);
}
static void AppendGroupDebugLine(const std::string& line) {
    HANDLE hFile = CreateFileW(GetGroupDebugPath().c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    DWORD written = 0;
    std::string out = line + "\r\n";
    WriteFile(hFile, out.c_str(), (DWORD)out.size(), &written, NULL);
    CloseHandle(hFile);
}
static void InitBaseDir() {
    if (!g_baseDir.empty()) return;
    wchar_t path[MAX_PATH] = {0}; GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path); PathAppendW(path, L"MyAsset");
    g_baseDir = path; CreateDirectoryW(g_baseDir.c_str(), nullptr);
}

struct AssetMetaFlags {
    bool favorite = false;
    bool fixedFrame = false;
    bool textStyle = false;
};

static bool g_assetMetaLoaded = false;
static std::map<std::wstring, AssetMetaFlags> g_assetMetaMap;

static void LoadPathSetFromLegacyFile(const std::wstring& filePath, std::set<std::wstring>& outSet) {
    outSet.clear();
    std::string data = ReadFileContent(filePath);
    std::stringstream ss(data);
    std::string line;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (line.empty()) continue;
        int sz = MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, NULL, 0);
        if (sz <= 0) continue;
        std::vector<wchar_t> buf((size_t)sz);
        MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, buf.data(), sz);
        outSet.insert(std::wstring(buf.data()));
    }
}

static std::string EscapeJsonString(const std::string& s) {
    std::string out;
    out.reserve(s.size() + 16);
    for (unsigned char ch : s) {
        switch (ch) {
        case '\"': out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\b': out += "\\b"; break;
        case '\f': out += "\\f"; break;
        case '\n': out += "\\n"; break;
        case '\r': out += "\\r"; break;
        case '\t': out += "\\t"; break;
        default:
            if (ch < 0x20) {
                char buf[8] = {};
                sprintf_s(buf, "\\u%04X", (unsigned int)ch);
                out += buf;
            } else {
                out.push_back((char)ch);
            }
            break;
        }
    }
    return out;
}

static void SyncSetsFromMetaMap() {
    // JSONメタを既存の高速参照セットに展開する。
    g_favPaths.clear();
    g_fixedPaths.clear();
    g_textStylePaths.clear();
    for (const auto& kv : g_assetMetaMap) {
        if (kv.second.favorite) g_favPaths.insert(kv.first);
        if (kv.second.fixedFrame) g_fixedPaths.insert(kv.first);
        if (kv.second.textStyle) g_textStylePaths.insert(kv.first);
    }
}

static void RebuildMetaMapFromSets() {
    // 既存処理が更新した set 群をJSON保存用mapへ戻す。
    std::map<std::wstring, AssetMetaFlags> nextMap;
    for (const auto& p : g_favPaths) nextMap[p].favorite = true;
    for (const auto& p : g_fixedPaths) nextMap[p].fixedFrame = true;
    for (const auto& p : g_textStylePaths) nextMap[p].textStyle = true;
    g_assetMetaMap.swap(nextMap);
}

static std::string BuildMetaJsonText() {
    RebuildMetaMapFromSets();
    std::ostringstream os;
    os << "{\n";
    os << "  \"version\": 1,\n";
    os << "  \"assets\": {\n";
    bool first = true;
    for (const auto& kv : g_assetMetaMap) {
        if (!first) os << ",\n";
        first = false;
        std::string pathUtf8 = EscapeJsonString(WideToUtf8(kv.first));
        os << "    \"" << pathUtf8 << "\": {";
        os << "\"favorite\": " << (kv.second.favorite ? "true" : "false") << ", ";
        os << "\"fixedFrame\": " << (kv.second.fixedFrame ? "true" : "false") << ", ";
        os << "\"textStyle\": " << (kv.second.textStyle ? "true" : "false") << "}";
    }
    os << "\n";
    os << "  }\n";
    os << "}\n";
    return os.str();
}

static bool JsonSkipWs(const std::string& s, size_t& i) {
    while (i < s.size() && isspace((unsigned char)s[i])) ++i;
    return i < s.size();
}

static bool JsonParseString(const std::string& s, size_t& i, std::string& out) {
    out.clear();
    if (i >= s.size() || s[i] != '\"') return false;
    ++i;
    while (i < s.size()) {
        char c = s[i++];
        if (c == '\"') return true;
        if (c == '\\') {
            if (i >= s.size()) return false;
            char e = s[i++];
            switch (e) {
            case '\"': out.push_back('\"'); break;
            case '\\': out.push_back('\\'); break;
            case '/': out.push_back('/'); break;
            case 'b': out.push_back('\b'); break;
            case 'f': out.push_back('\f'); break;
            case 'n': out.push_back('\n'); break;
            case 'r': out.push_back('\r'); break;
            case 't': out.push_back('\t'); break;
            case 'u':
                if (i + 4 > s.size()) return false;
                out.push_back('?');
                i += 4;
                break;
            default:
                return false;
            }
        } else {
            out.push_back(c);
        }
    }
    return false;
}

static bool JsonParseBool(const std::string& s, size_t& i, bool& out) {
    if (i + 4 <= s.size() && s.compare(i, 4, "true") == 0) { i += 4; out = true; return true; }
    if (i + 5 <= s.size() && s.compare(i, 5, "false") == 0) { i += 5; out = false; return true; }
    return false;
}

static bool JsonSkipValue(const std::string& s, size_t& i);

static bool JsonSkipObject(const std::string& s, size_t& i) {
    if (i >= s.size() || s[i] != '{') return false;
    ++i;
    JsonSkipWs(s, i);
    if (i < s.size() && s[i] == '}') { ++i; return true; }
    while (i < s.size()) {
        std::string key;
        if (!JsonParseString(s, i, key)) return false;
        JsonSkipWs(s, i);
        if (i >= s.size() || s[i] != ':') return false;
        ++i;
        JsonSkipWs(s, i);
        if (!JsonSkipValue(s, i)) return false;
        JsonSkipWs(s, i);
        if (i < s.size() && s[i] == ',') { ++i; JsonSkipWs(s, i); continue; }
        if (i < s.size() && s[i] == '}') { ++i; return true; }
        return false;
    }
    return false;
}

static bool JsonSkipArray(const std::string& s, size_t& i) {
    if (i >= s.size() || s[i] != '[') return false;
    ++i;
    JsonSkipWs(s, i);
    if (i < s.size() && s[i] == ']') { ++i; return true; }
    while (i < s.size()) {
        if (!JsonSkipValue(s, i)) return false;
        JsonSkipWs(s, i);
        if (i < s.size() && s[i] == ',') { ++i; JsonSkipWs(s, i); continue; }
        if (i < s.size() && s[i] == ']') { ++i; return true; }
        return false;
    }
    return false;
}

static bool JsonSkipValue(const std::string& s, size_t& i) {
    JsonSkipWs(s, i);
    if (i >= s.size()) return false;
    if (s[i] == '{') return JsonSkipObject(s, i);
    if (s[i] == '[') return JsonSkipArray(s, i);
    if (s[i] == '\"') { std::string dummy; return JsonParseString(s, i, dummy); }
    bool b = false;
    if (JsonParseBool(s, i, b)) return true;
    if (i + 4 <= s.size() && s.compare(i, 4, "null") == 0) { i += 4; return true; }
    size_t begin = i;
    if (s[i] == '-' || s[i] == '+') ++i;
    while (i < s.size() && (isdigit((unsigned char)s[i]) || s[i] == '.' || s[i] == 'e' || s[i] == 'E' || s[i] == '+' || s[i] == '-')) ++i;
    return i > begin;
}

static bool ParseAssetFlagsObject(const std::string& s, size_t& i, AssetMetaFlags& flags) {
    if (i >= s.size() || s[i] != '{') return false;
    ++i;
    JsonSkipWs(s, i);
    if (i < s.size() && s[i] == '}') { ++i; return true; }
    while (i < s.size()) {
        std::string key;
        if (!JsonParseString(s, i, key)) return false;
        JsonSkipWs(s, i);
        if (i >= s.size() || s[i] != ':') return false;
        ++i;
        JsonSkipWs(s, i);
        if (key == "favorite") {
            if (!JsonParseBool(s, i, flags.favorite)) return false;
        } else if (key == "fixedFrame") {
            if (!JsonParseBool(s, i, flags.fixedFrame)) return false;
        } else if (key == "textStyle") {
            if (!JsonParseBool(s, i, flags.textStyle)) return false;
        } else {
            if (!JsonSkipValue(s, i)) return false;
        }
        JsonSkipWs(s, i);
        if (i < s.size() && s[i] == ',') { ++i; JsonSkipWs(s, i); continue; }
        if (i < s.size() && s[i] == '}') { ++i; return true; }
        return false;
    }
    return false;
}

static bool ParseMetaAssetsObject(const std::string& s, size_t& i, std::map<std::wstring, AssetMetaFlags>& outMap) {
    if (i >= s.size() || s[i] != '{') return false;
    ++i;
    JsonSkipWs(s, i);
    if (i < s.size() && s[i] == '}') { ++i; return true; }
    while (i < s.size()) {
        std::string pathKey;
        if (!JsonParseString(s, i, pathKey)) return false;
        JsonSkipWs(s, i);
        if (i >= s.size() || s[i] != ':') return false;
        ++i;
        JsonSkipWs(s, i);
        AssetMetaFlags flags = {};
        if (!ParseAssetFlagsObject(s, i, flags)) return false;
        std::wstring pathW = Utf8ToWide(pathKey);
        if (!pathW.empty()) outMap[pathW] = flags;
        JsonSkipWs(s, i);
        if (i < s.size() && s[i] == ',') { ++i; JsonSkipWs(s, i); continue; }
        if (i < s.size() && s[i] == '}') { ++i; return true; }
        return false;
    }
    return false;
}

static bool ParseMetaJsonText(const std::string& s, std::map<std::wstring, AssetMetaFlags>& outMap) {
    outMap.clear();
    size_t i = 0;
    JsonSkipWs(s, i);
    if (i >= s.size() || s[i] != '{') return false;
    ++i;
    JsonSkipWs(s, i);
    while (i < s.size() && s[i] != '}') {
        std::string key;
        if (!JsonParseString(s, i, key)) return false;
        JsonSkipWs(s, i);
        if (i >= s.size() || s[i] != ':') return false;
        ++i;
        JsonSkipWs(s, i);
        if (key == "assets") {
            if (!ParseMetaAssetsObject(s, i, outMap)) return false;
        } else {
            if (!JsonSkipValue(s, i)) return false;
        }
        JsonSkipWs(s, i);
        if (i < s.size() && s[i] == ',') { ++i; JsonSkipWs(s, i); continue; }
        break;
    }
    JsonSkipWs(s, i);
    return (i < s.size() && s[i] == '}');
}

static bool SaveAssetMeta() {
    std::string jsonText = BuildMetaJsonText();
    WriteFileContent(GetMetaFilePath(), jsonText);
    return true;
}

static void LoadLegacyMetaAndMigrate() {
    LoadPathSetFromLegacyFile(GetFavFilePath(), g_favPaths);
    LoadPathSetFromLegacyFile(GetFixedFilePath(), g_fixedPaths);
    LoadPathSetFromLegacyFile(GetTextStyleFilePath(), g_textStylePaths);
    if (SaveAssetMeta()) {
        DeleteFileW(GetFavFilePath().c_str());
        DeleteFileW(GetFixedFilePath().c_str());
        DeleteFileW(GetTextStyleFilePath().c_str());
    }
}

static void LoadAssetMetaUnified(bool forceReload) {
    if (g_assetMetaLoaded && !forceReload) return;
    g_assetMetaLoaded = true;

    std::string text = ReadFileContent(GetMetaFilePath());
    std::map<std::wstring, AssetMetaFlags> parsedMap;
    if (!text.empty() && ParseMetaJsonText(text, parsedMap)) {
        g_assetMetaMap.swap(parsedMap);
        SyncSetsFromMetaMap();
        return;
    }

    // JSON がない or 壊れている場合は legacy から移行
    LoadLegacyMetaAndMigrate();
}

void SaveFavorites() {
    SaveAssetMeta();
}
void LoadFavorites() {
    LoadAssetMetaUnified(true);
}
static void ToggleFavorite(const std::wstring& path) { 
    if (g_favPaths.count(path)) g_favPaths.erase(path); else g_favPaths.insert(path); 
    SaveFavorites(); 
}

void SaveFixedFrames() { 
    SaveAssetMeta();
}
void LoadFixedFrames() { 
    LoadAssetMetaUnified(false);
} 
static void ToggleFixedFrame(const std::wstring& path) { 
    if (g_fixedPaths.count(path)) g_fixedPaths.erase(path); else g_fixedPaths.insert(path); 
    SaveFixedFrames(); 
}

void SaveTextStyleFlags() {
    SaveAssetMeta();
}
void LoadTextStyleFlags() {
    LoadAssetMetaUnified(false);
}

static void SaveConfig() { 
    WritePrivateProfileStringW(L"Settings", L"GifSpeed", std::to_wstring(g_gifSpeedPercent).c_str(), GetConfigPath().c_str()); 
    WritePrivateProfileStringW(L"Settings", L"ShowGifExportGuide", g_showGifExportGuide ? L"1" : L"0", GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"GifExportKeepOriginal", g_gifExportKeepOriginal ? L"1" : L"0", GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"SortMode", std::to_wstring(g_sortMode).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"PreviewPlaybackMode", std::to_wstring(g_previewPlaybackMode).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"PreviewPerfMode", std::to_wstring(g_previewPerfMode).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"PreviewZoom", std::to_wstring(g_previewZoomPercent).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"PreviewThumbOnly", g_previewThumbOnly ? L"1" : L"0", GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"CapturePresetSlot", std::to_wstring(g_capturePresetSlot).c_str(), GetConfigPath().c_str());
    for (int i = 0; i < 3; ++i) {
        std::wstring keyRect = L"CapturePreset" + std::wstring(1, (wchar_t)(L'A' + i));
        std::wstring keyValid = keyRect + L"Valid";
        if (g_capturePresetValid[i]) WritePrivateProfileStringW(L"Settings", keyRect.c_str(), CaptureRectToConfigText(g_capturePresetRects[i]).c_str(), GetConfigPath().c_str());
        else WritePrivateProfileStringW(L"Settings", keyRect.c_str(), L"", GetConfigPath().c_str());
        WritePrivateProfileStringW(L"Settings", keyValid.c_str(), g_capturePresetValid[i] ? L"1" : L"0", GetConfigPath().c_str());
    }
    WritePrivateProfileStringW(L"Settings", L"EnableTextStyle", g_enableTextStyle ? L"1" : L"0", GetConfigPath().c_str());
    unsigned long long inheritMask = 0ULL;
    for (int i = 0; i < kInheritItemCount; ++i) if (g_inheritItemEnabled[i]) inheritMask |= (1ULL << i);
    WritePrivateProfileStringW(L"Settings", L"TextStyleInheritMask", std::to_wstring(inheritMask).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"TextStyleInsertMode", std::to_wstring(g_textStyleInsertMode).c_str(), GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"TextStyleClearEffectsFrom2", g_textStyleClearEffectsFrom2 ? L"1" : L"0", GetConfigPath().c_str());
    WritePrivateProfileStringW(L"Settings", L"TextStyleHideInMainList", g_hideTextStyleInMainList ? L"1" : L"0", GetConfigPath().c_str());
    for (int i = 0; i < kThemeColorCount; ++i) {
        WritePrivateProfileStringW(L"Settings", kThemeColors[i].key, ColorToHex(*kThemeColors[i].value).c_str(), GetConfigPath().c_str());
    }
}
static void LoadConfig() { 
    g_gifSpeedPercent = GetPrivateProfileIntW(L"Settings", L"GifSpeed", 100, GetConfigPath().c_str()); 
    if (g_gifSpeedPercent < 10) g_gifSpeedPercent = 100;
    g_showGifExportGuide = (GetPrivateProfileIntW(L"Settings", L"ShowGifExportGuide", 1, GetConfigPath().c_str()) != 0);
    g_gifExportKeepOriginal = (GetPrivateProfileIntW(L"Settings", L"GifExportKeepOriginal", 0, GetConfigPath().c_str()) != 0);
    g_sortMode = GetPrivateProfileIntW(L"Settings", L"SortMode", 0, GetConfigPath().c_str());
    if (g_sortMode < 0 || g_sortMode > 3) g_sortMode = 0;
    g_previewPlaybackMode = GetPrivateProfileIntW(L"Settings", L"PreviewPlaybackMode", 1, GetConfigPath().c_str());
    if (g_previewPlaybackMode < 0 || g_previewPlaybackMode > 1) g_previewPlaybackMode = 1;
    g_previewPerfMode = GetPrivateProfileIntW(L"Settings", L"PreviewPerfMode", 0, GetConfigPath().c_str());
    if (g_previewPerfMode < 0 || g_previewPerfMode > 1) g_previewPerfMode = 0;
    g_previewZoomPercent = GetPrivateProfileIntW(L"Settings", L"PreviewZoom", 100, GetConfigPath().c_str());
    if (g_previewZoomPercent < 50) g_previewZoomPercent = 50;
    if (g_previewZoomPercent > 200) g_previewZoomPercent = 200;
    g_previewThumbOnly = (GetPrivateProfileIntW(L"Settings", L"PreviewThumbOnly", 0, GetConfigPath().c_str()) != 0);
    g_enableGroupDebugDump = kEnableDebugByCode;
    g_capturePresetSlot = GetPrivateProfileIntW(L"Settings", L"CapturePresetSlot", 0, GetConfigPath().c_str());
    if (g_capturePresetSlot < 0 || g_capturePresetSlot > 2) g_capturePresetSlot = 0;
    for (int i = 0; i < 3; ++i) {
        std::wstring keyRect = L"CapturePreset" + std::wstring(1, (wchar_t)(L'A' + i));
        std::wstring keyValid = keyRect + L"Valid";
        g_capturePresetValid[i] = (GetPrivateProfileIntW(L"Settings", keyValid.c_str(), 0, GetConfigPath().c_str()) != 0);
        wchar_t buf[128] = {};
        GetPrivateProfileStringW(L"Settings", keyRect.c_str(), L"", buf, 128, GetConfigPath().c_str());
        RECT rc = {};
        if (!ParseCaptureRectText(buf, rc)) g_capturePresetValid[i] = false;
        else g_capturePresetRects[i] = rc;
    }
    g_enableTextStyle = (GetPrivateProfileIntW(L"Settings", L"EnableTextStyle", 1, GetConfigPath().c_str()) != 0);
    if (!g_enableTextStyle) g_assetViewTab = 0;
    wchar_t maskBuf[64] = {};
    GetPrivateProfileStringW(L"Settings", L"TextStyleInheritMask", L"1", maskBuf, 64, GetConfigPath().c_str());
    unsigned long long mask = _wcstoui64(maskBuf, nullptr, 10);
    for (int i = 0; i < kInheritItemCount; ++i) g_inheritItemEnabled[i] = ((mask >> i) & 1ULL) != 0;
    g_textStyleInsertMode = GetPrivateProfileIntW(L"Settings", L"TextStyleInsertMode", 0, GetConfigPath().c_str());
    if (g_textStyleInsertMode < 0 || g_textStyleInsertMode > 1) g_textStyleInsertMode = 0;
    g_textStyleClearEffectsFrom2 = (GetPrivateProfileIntW(L"Settings", L"TextStyleClearEffectsFrom2", 0, GetConfigPath().c_str()) != 0);
    g_hideTextStyleInMainList = (GetPrivateProfileIntW(L"Settings", L"TextStyleHideInMainList", 0, GetConfigPath().c_str()) != 0);
    for (int i = 0; i < kThemeColorCount; ++i) {
        wchar_t buf[32] = {};
        GetPrivateProfileStringW(L"Settings", kThemeColors[i].key, ColorToHex(kThemeColors[i].defValue).c_str(), buf, 32, GetConfigPath().c_str());
        COLORREF c = kThemeColors[i].defValue;
        if (TryParseHexColor(buf, c)) *kThemeColors[i].value = c;
        else *kThemeColors[i].value = kThemeColors[i].defValue;
    }
    RecreateUiBrushes();
}

static void BroadcastTextStyleConfigChanged() {
    if (g_hwnd && IsWindow(g_hwnd)) InvalidateRect(g_hwnd, NULL, FALSE);
    if (g_hDlg && IsWindow(g_hDlg)) InvalidateRect(g_hDlg, NULL, FALSE);
    if (g_hInfoWnd && IsWindow(g_hInfoWnd)) InvalidateRect(g_hInfoWnd, NULL, FALSE);
    if (g_hSettingDlg && IsWindow(g_hSettingDlg)) PostMessageW(g_hSettingDlg, WM_APP + 21, 0, 0);
    if (g_hTextStyleQuickDlg && IsWindow(g_hTextStyleQuickDlg)) PostMessageW(g_hTextStyleQuickDlg, WM_APP + 21, 0, 0);
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

static std::wstring GetCategoryDirPath(const std::wstring& category) {
    InitBaseDir();
    if (_wcsicmp(category.c_str(), L"ALL") == 0) return g_baseDir;
    std::wstring dir = g_baseDir + L"\\" + category;
    if (PathFileExistsW(dir.c_str())) return dir;
    if (_wcsicmp(category.c_str(), L"Main") == 0) return g_baseDir;
    return dir;
}

static std::wstring GetCategoryOrderFilePath(const std::wstring& category) {
    if (_wcsicmp(category.c_str(), L"ALL") == 0) return g_baseDir + L"\\MyAssetOrderAll.json";
    return GetCategoryDirPath(category) + L"\\MyAssetOrder.json";
}

static std::wstring GetAssetOrderKey(const std::wstring& category, const std::wstring& assetPath) {
    if (_wcsicmp(category.c_str(), L"ALL") == 0) {
        std::wstring prefix = g_baseDir + L"\\";
        if (assetPath.size() > prefix.size()) {
            std::wstring head = assetPath.substr(0, prefix.size());
            if (_wcsicmp(head.c_str(), prefix.c_str()) == 0) {
                return assetPath.substr(prefix.size());
            }
        }
    }
    std::wstring base = GetCategoryDirPath(category);
    std::wstring prefix = base + L"\\";
    if (assetPath.size() > prefix.size()) {
        std::wstring head = assetPath.substr(0, prefix.size());
        if (_wcsicmp(head.c_str(), prefix.c_str()) == 0) {
            return assetPath.substr(prefix.size());
        }
    }
    size_t p = assetPath.find_last_of(L"\\/");
    return (p == std::wstring::npos) ? assetPath : assetPath.substr(p + 1);
}

static bool ParseOrderJson(const std::string& text, std::vector<std::wstring>& outOrder) {
    outOrder.clear();
    size_t keyPos = text.find("\"order\"");
    if (keyPos == std::string::npos) return false;
    size_t i = text.find('[', keyPos);
    if (i == std::string::npos) return false;
    ++i;

    auto skipWs = [&](size_t& p) {
        while (p < text.size() && isspace((unsigned char)text[p])) ++p;
    };

    skipWs(i);
    if (i < text.size() && text[i] == ']') return true;

    while (i < text.size()) {
        skipWs(i);
        std::string s;
        if (!JsonParseString(text, i, s)) return false;
        std::wstring w = Utf8ToWide(s);
        if (!w.empty()) outOrder.push_back(w);
        skipWs(i);
        if (i < text.size() && text[i] == ',') { ++i; continue; }
        if (i < text.size() && text[i] == ']') return true;
        return false;
    }
    return false;
}

static std::string BuildOrderJson(const std::vector<std::wstring>& order) {
    std::ostringstream os;
    os << "{\n";
    os << "  \"version\": 1,\n";
    os << "  \"order\": [\n";
    for (size_t i = 0; i < order.size(); ++i) {
        if (i > 0) os << ",\n";
        os << "    \"" << EscapeJsonString(WideToUtf8(order[i])) << "\"";
    }
    os << "\n";
    os << "  ]\n";
    os << "}\n";
    return os.str();
}

static bool LoadCustomOrderForCategory(const std::wstring& category, std::vector<std::wstring>& outOrder) {
    outOrder.clear();
    std::string text = ReadFileContent(GetCategoryOrderFilePath(category));
    if (text.empty()) return false;
    return ParseOrderJson(text, outOrder);
}

static bool SaveCustomOrderForCategory(const std::wstring& category, const std::vector<std::wstring>& order) {
    CreateDirectoryW(GetCategoryDirPath(category).c_str(), nullptr);
    std::wstring finalPath = GetCategoryOrderFilePath(category);
    std::wstring tempPath = finalPath + L".tmp";
    std::string json = BuildOrderJson(order);
    WriteFileContent(tempPath, json);
    if (!MoveFileExW(tempPath.c_str(), finalPath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        DeleteFileW(tempPath.c_str());
        WriteFileContent(finalPath, json);
    }
    return true;
}

static void ReconcileCustomOrderForCategory(const std::wstring& category) {
    if (category.empty()) return;

    std::vector<std::wstring> order;
    auto itCache = g_customOrderByCategory.find(category);
    if (itCache != g_customOrderByCategory.end()) {
        order = itCache->second;
    } else {
        LoadCustomOrderForCategory(category, order);
    }

    std::vector<std::wstring> existing;
    existing.reserve(g_assets.size());
    std::set<std::wstring> existingSet;
    for (const auto& a : g_assets) {
        if (category != L"ALL" && a.category != category) continue;
        std::wstring key = GetAssetOrderKey(category, a.path);
        existing.push_back(key);
        existingSet.insert(key);
    }

    std::vector<std::wstring> nextOrder;
    nextOrder.reserve(existing.size());
    std::set<std::wstring> used;
    for (const auto& k : order) {
        if (existingSet.count(k) == 0) continue;
        if (used.insert(k).second) nextOrder.push_back(k);
    }
    for (const auto& k : existing) {
        if (used.insert(k).second) nextOrder.push_back(k);
    }

    bool changed = (nextOrder.size() != order.size());
    if (!changed) {
        for (size_t i = 0; i < nextOrder.size(); ++i) {
            if (nextOrder[i] != order[i]) { changed = true; break; }
        }
    }

    g_customOrderByCategory[category] = nextOrder;
    if (changed) SaveCustomOrderForCategory(category, nextOrder);
}

static void EnsureCustomOrderReadyForCurrentCategory() {
    if (g_currentCategory.empty()) return;
    ReconcileCustomOrderForCategory(g_currentCategory);
}

static bool CanStartCustomReorder() {
    if (g_sortMode != 3) return false;
    if (g_currentCategory.empty()) return false;
    if (!g_searchQuery.empty()) return false;
    if (g_enableTextStyle && g_assetViewTab != 0) return false;
    if (g_hideTextStyleInMainList) return false;
    return true;
}

static int GetInsertPosFromPoint(int x, int y, const RECT& rcClient) {
    int count = (int)g_displayAssets.size();
    if (count <= 0) return 0;
    int listTop = GetListTopY();
    int listBottom = rcClient.bottom - GetBottomReservedHeight();
    int itemH = GetListItemHeight();
    int cols = (std::max)(1, (int)rcClient.right / GetListItemMinWidth());
    int cw = (int)rcClient.right / cols;

    if (y <= listTop) return 0;
    if (y >= listBottom) return count;

    int row = (y - listTop + g_scrollY) / itemH;
    if (row < 0) row = 0;
    int col = x / cw;
    if (col < 0) col = 0;
    if (col >= cols) col = cols - 1;

    int idx = row * cols + col;
    if (idx < 0) idx = 0;
    if (idx >= count) return count;

    int localX = x - col * cw;
    bool after = localX >= (cw / 2);
    int insertPos = idx + (after ? 1 : 0);
    if (insertPos < 0) insertPos = 0;
    if (insertPos > count) insertPos = count;
    return insertPos;
}

static bool ApplyCustomReorderFromInsertPos() {
    int count = (int)g_displayAssets.size();
    if (count <= 1) return false;
    if (g_reorderFromIndex < 0 || g_reorderFromIndex >= count) return false;
    if (g_reorderInsertPos < 0 || g_reorderInsertPos > count) return false;

    int dst = g_reorderInsertPos;
    if (dst > g_reorderFromIndex) dst--;
    if (dst < 0) dst = 0;
    if (dst >= count) dst = count - 1;
    if (dst == g_reorderFromIndex) return false;

    Asset* moving = g_displayAssets[g_reorderFromIndex];
    g_displayAssets.erase(g_displayAssets.begin() + g_reorderFromIndex);
    g_displayAssets.insert(g_displayAssets.begin() + dst, moving);
    g_selectedIndex = dst;
    return true;
}

static bool IsGifFile(const std::wstring& path) {
    std::wstring ext = path.substr(path.find_last_of(L"."));
    return (ToLower(ext) == L".gif");
}

static HWND ResolveHostWindow(HWND requester) {
    auto normalizeToTopOwner = [](HWND h) -> HWND {
        HWND cur = h;
        for (int i = 0; i < 8 && cur && IsWindow(cur); ++i) {
            HWND owner = GetWindow(cur, GW_OWNER);
            if (!owner || !IsWindow(owner)) break;
            cur = owner;
        }
        return cur;
    };

    auto pickHost = [](HWND base) -> HWND {
        if (!base || !IsWindow(base)) return nullptr;
        HWND owner = GetWindow(base, GW_OWNER);
        if (owner && IsWindow(owner) && owner != g_hwnd) return owner;
        HWND rootOwner = GetAncestor(base, GA_ROOTOWNER);
        if (rootOwner && IsWindow(rootOwner) && rootOwner != g_hwnd) return rootOwner;
        return nullptr;
    };

    HWND host = pickHost(g_hwnd);
    if (host) return normalizeToTopOwner(host);
    host = pickHost(requester);
    if (host) return normalizeToTopOwner(host);

    HWND fg = GetForegroundWindow();
    if (fg && IsWindow(fg) && fg != g_hwnd) {
        HWND fgRootOwner = GetAncestor(fg, GA_ROOTOWNER);
        if (fgRootOwner != g_hwnd) return normalizeToTopOwner(fg);
    }
    return nullptr;
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
    return ExtractHeaderFrameRange(alias, outStart, outEnd);
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

static bool CacheRangeFromTimelineSelection() {
    g_cachedOutputRangeValid = false;
    if (!g_editHandle || !g_editHandle->call_edit_section_param) return false;

    OutputRangeContext ctx = {0, 0, false};
    if (!g_editHandle->call_edit_section_param(&ctx, CollectSelectedRange)) return false;
    if (!ctx.valid) return false;

    g_cachedOutputRangeStart = ctx.frameStart;
    g_cachedOutputRangeEnd = ctx.frameEnd;
    g_cachedOutputRangeValid = true;
    return true;
}

static bool ApplyCachedRangeToTimeline() {
    if (!g_cachedOutputRangeValid) return false;
    if (!g_editHandle || !g_editHandle->call_edit_section_param) return false;
    OutputRangeContext ctx = { g_cachedOutputRangeStart, g_cachedOutputRangeEnd, true };
    return g_editHandle->call_edit_section_param(&ctx, ApplyRangeFromContextProc) ? true : false;
}

struct LayerApiCompareContext {
    const std::vector<void*>* objs = nullptr;
    int total = 0;
    int apiValid = 0;
    int match = 0;
    int mismatch = 0;
};

static void LayerApiCompareProc(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (LayerApiCompareContext*)param;
    if (!ctx || !ctx->objs || !edit || !edit->get_object_layer_frame) return;
    if (ctx->objs->empty()) return;

    AppendGroupDebugLine("  ---- layer api compare begin ----");
    for (size_t i = 0; i < ctx->objs->size(); ++i) {
        void* obj = (*ctx->objs)[i];
        if (!obj) continue;
        ctx->total++;

        OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(obj);
        int apiLayer = lf.layer;
        int memOff = 0;
        int memLayer = GetObjectLayerIndex(obj, &memOff);

        bool ok = (apiLayer > 0);
        if (ok) ctx->apiValid++;
        bool matched = (ok && memLayer > 0 && apiLayer == memLayer);
        if (matched) ctx->match++;
        else if (ok) ctx->mismatch++;

        std::ostringstream os;
        os << "  layer-api idx=" << i
           << " api=" << apiLayer
           << " mem=" << memLayer
           << " memOff=0x" << std::hex << memOff << std::dec
           << " match=" << (matched ? 1 : 0);
        AppendGroupDebugLine(os.str());
    }
    std::ostringstream sum;
    sum << "  layer-api summary total=" << ctx->total
        << " apiValid=" << ctx->apiValid
        << " match=" << ctx->match
        << " mismatch=" << ctx->mismatch;
    AppendGroupDebugLine(sum.str());
    AppendGroupDebugLine("  ---- layer api compare end ----");
}

static void RunLayerApiCompareDebug() {
    // 追加直後は選択状態が変わりやすいため、事前保持した配列で比較する。
    if (!g_enableGroupDebugDump) return;
    if (g_pendingLayerApiCompareObjs.empty()) {
        AppendGroupDebugLine("  layer-api compare skipped: pending object list empty");
        return;
    }
    if (!g_editHandle || !g_editHandle->call_edit_section_param) {
        AppendGroupDebugLine("  layer-api compare skipped: edit handle unavailable");
        return;
    }
    LayerApiCompareContext ctx = {};
    ctx.objs = &g_pendingLayerApiCompareObjs;
    if (!g_editHandle->call_edit_section_param(&ctx, LayerApiCompareProc)) {
        AppendGroupDebugLine("  layer-api compare failed: call_edit_section_param=false");
    }
    g_pendingLayerApiCompareObjs.clear();
}

struct AddAssetApiLayerContext {
    const std::vector<PendingAddObjSnapshot>* snapshots = nullptr;
    std::vector<int>* layers = nullptr;
    std::vector<int>* starts = nullptr;
    std::vector<int>* ends = nullptr;
};

static void CollectApiLayersForPendingAddProc(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (AddAssetApiLayerContext*)param;
    if (!ctx || !ctx->snapshots || !ctx->layers || !ctx->starts || !ctx->ends) return;
    if (!edit || !edit->get_object_layer_frame) return;

    size_t n = ctx->snapshots->size();
    ctx->layers->assign(n, 0);
    ctx->starts->assign(n, 0);
    ctx->ends->assign(n, 0);
    for (size_t i = 0; i < n; ++i) {
        void* obj = (*ctx->snapshots)[i].obj;
        if (!obj) continue;
        OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(obj);
        (*ctx->layers)[i] = lf.layer;
        (*ctx->starts)[i] = lf.start;
        (*ctx->ends)[i] = lf.end;
    }
}

static void ProcessPendingAddAssetRequestByApi() {
    // call_edit_section内で重い処理を続けると固まりやすいため、
    // OnAddAssetでスナップショット化→ここで非同期に本処理を行う。
    if (g_pendingAddObjSnapshots.empty()) return;

    std::vector<PendingAddObjSnapshot> snaps;
    snaps.swap(g_pendingAddObjSnapshots);

    std::vector<int> apiLayers;
    std::vector<int> apiStarts;
    std::vector<int> apiEnds;
    if (g_editHandle && g_editHandle->call_edit_section_param) {
        AddAssetApiLayerContext ctx = {};
        ctx.snapshots = &snaps;
        ctx.layers = &apiLayers;
        ctx.starts = &apiStarts;
        ctx.ends = &apiEnds;
        g_editHandle->call_edit_section_param(&ctx, CollectApiLayersForPendingAddProc);
    }
    if (apiLayers.size() != snaps.size()) apiLayers.assign(snaps.size(), 0);
    if (apiStarts.size() != snaps.size()) apiStarts.assign(snaps.size(), 0);
    if (apiEnds.size() != snaps.size()) apiEnds.assign(snaps.size(), 0);

    // SDK差分対策:
    // get_object_layer_frame().layer が 0-based で返る環境では最上段が 0 になる。
    // 0 が含まれる場合のみ +1 して内部表現を 1-based に揃える。
    bool apiLayerZeroBased = false;
    for (int v : apiLayers) {
        if (v == 0) { apiLayerZeroBased = true; break; }
    }
    auto normalizeApiLayer = [&](int v) -> int {
        if (v < 0) return 0;
        return apiLayerZeroBased ? (v + 1) : v;
    };

    g_addDialogRangeValid = false;
    g_addDialogRangeStart = 0;
    g_addDialogRangeEnd = 0;
    // GIF書き出し用の選択範囲もここで再構築しておく。
    int rStart = INT_MAX, rEnd = INT_MIN;
    bool hasRange = false;
    for (size_t i = 0; i < snaps.size(); ++i) {
        // レイヤー値は範囲計算に不要。frameのみ有効なら採用する。
        if (apiEnds[i] < apiStarts[i]) continue;
        rStart = (std::min)(rStart, apiStarts[i]);
        rEnd = (std::max)(rEnd, apiEnds[i]);
        hasRange = true;
    }
    if (hasRange) {
        g_addDialogRangeValid = true;
        g_addDialogRangeStart = rStart;
        g_addDialogRangeEnd = rEnd;
    }

    std::string bodyData;
    int count = (int)snaps.size();
    if (count == 1) {
        bodyData = snaps[0].alias;
    } else {
        // 相対layer化の基準。最小レイヤーを1として再マップする。
        int minLayer = INT_MAX;
        for (int i = 0; i < count; ++i) {
            int layer = normalizeApiLayer(apiLayers[(size_t)i]);
            if (layer <= 0) {
                layer = ResolveObjectLayerForAssetBuild(snaps[(size_t)i].obj, snaps[(size_t)i].alias.empty() ? nullptr : snaps[(size_t)i].alias.c_str());
            }
            if (layer > 0 && layer < minLayer) minLayer = layer;
        }
        if (minLayer == INT_MAX) minLayer = 1;

        std::vector<int> relativeGroups((size_t)count, 0);
        std::map<long long, int> groupIdMap;
        std::map<long long, int> groupCounts;
        std::vector<long long> originalGroups((size_t)count, 0);
        int nextGroupId = 1;
        // group再採番のために、同一グループの対応表を作る。
        bool hasGroupToken = false;

        for (int i = 0; i < count; ++i) {
            const std::string& rawAlias = snaps[(size_t)i].alias;
            int aliasGroupId = rawAlias.empty() ? 0 : ExtractHeaderIntValue(rawAlias, "group=");
            long long originalGroupId = aliasGroupId;
            if (originalGroupId <= 0) {
                int verifiedGroup = GetVerifiedTimelineGroupId(snaps[(size_t)i].obj);
                if (verifiedGroup > 0) originalGroupId = verifiedGroup;
            }
            uintptr_t ptr28 = 0;
            if (originalGroupId <= 0 && SafeReadUIntPtrAtOffset(snaps[(size_t)i].obj, 0x28, ptr28) && ptr28 != 0) {
                originalGroupId = (long long)ptr28;
            }
            originalGroups[(size_t)i] = originalGroupId;
            if (originalGroupId > 0) {
                hasGroupToken = true;
                groupCounts[originalGroupId]++;
            }
        }

        if (hasGroupToken) {
            for (int i = 0; i < count; ++i) {
                long long gid = originalGroups[(size_t)i];
                if (gid <= 0) continue;
                if (groupCounts[gid] < 2) continue;
                auto it = groupIdMap.find(gid);
                if (it == groupIdMap.end()) {
                    groupIdMap[gid] = nextGroupId;
                    relativeGroups[(size_t)i] = nextGroupId;
                    nextGroupId++;
                } else {
                    relativeGroups[(size_t)i] = it->second;
                }
            }
        }

        bool shouldEmitGroupField = false;
        for (int g : relativeGroups) if (g > 0) { shouldEmitGroupField = true; break; }

        for (int i = 0; i < count; ++i) {
            const std::string& rawAlias = snaps[(size_t)i].alias;
            if (rawAlias.empty()) continue;
            std::string s = rawAlias;
            std::string headerNew = "[" + std::to_string(i) + "]";
            std::string sectionNew = "[" + std::to_string(i) + ".";

            int currentLayer = normalizeApiLayer(apiLayers[(size_t)i]);
            if (currentLayer <= 0) {
                currentLayer = ResolveObjectLayerForAssetBuild(snaps[(size_t)i].obj, rawAlias.c_str());
            }
            int relativeLayer = (currentLayer > 0) ? ((currentLayer - minLayer) + 1) : 1;

            ReplaceStringAll(s, "[Object.", sectionNew);
            ReplaceStringAll(s, "[Object]", headerNew);
            RemoveLinesStartingWith(s, "group=");
            size_t pos = s.find(headerNew);
            if (pos != std::string::npos) {
                std::string headerLines = "\r\nlayer=" + std::to_string(relativeLayer);
                if (relativeGroups[(size_t)i] > 0) {
                    headerLines += "\r\ngroup=" + std::to_string(relativeGroups[(size_t)i]);
                } else if (shouldEmitGroupField) {
                    headerLines += "\r\ngroup=0";
                }
                s.insert(pos + headerNew.length(), headerLines);
            }
            bodyData += s + "\r\n";
        }
    }

    int parsedStart = 0, parsedEnd = 0;
    if (!g_addDialogRangeValid && TryParseFrameRangeFromAliasData(bodyData, parsedStart, parsedEnd)) {
        g_addDialogRangeValid = true;
        g_addDialogRangeStart = parsedStart;
        g_addDialogRangeEnd = parsedEnd;
    }
    g_addDialogFromMyAssetAdd = true;
    OpenAddDialog(bodyData, false);
}

static bool ReplaceObjectAliasByRecreate(EDIT_SECTION_OUT_SAFE* edit, void* targetObj, const std::string& newAlias, void** outNewObj) {
    if (outNewObj) *outNewObj = nullptr;
    if (!edit || !targetObj || !edit->get_object_layer_frame || !edit->delete_object) return false;
    if (!edit->create_object_from_alias || !edit->get_object_alias) return false;

    using FN_CREATE_OBJECT_FROM_ALIAS = void* (*)(LPCSTR alias, int layer, int frame, int length);
    FN_CREATE_OBJECT_FROM_ALIAS fnCreate = (FN_CREATE_OBJECT_FROM_ALIAS)edit->create_object_from_alias;
    if (!fnCreate) return false;

    OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(targetObj);
    int len = lf.end - lf.start + 1;
    if (len < 1) len = 1;

    // 失敗時復旧用に旧aliasを保持
    std::string oldAlias;
    if (LPCSTR oldRaw = edit->get_object_alias(targetObj)) oldAlias = oldRaw;

    struct CreateCandidate { int layer; int frame; };
    std::vector<CreateCandidate> cands;
    cands.push_back({ lf.layer, lf.end + (std::max)(2, len + 2) });
    cands.push_back({ lf.layer + 32, lf.start });
    cands.push_back({ (std::max)(1, lf.layer - 32), lf.start });
    cands.push_back({ lf.layer + 64, lf.end + (std::max)(2, len + 8) });

    void* newObj = nullptr;
    for (const auto& c : cands) {
        newObj = fnCreate(newAlias.c_str(), c.layer, c.frame, len);
        if (newObj) break;
    }
    if (!newObj) return false;

    // 新規生成成功後に旧オブジェクトを削除して置換
    edit->delete_object(targetObj);

    if (!edit->move_object) {
        edit->delete_object(newObj);
        if (!oldAlias.empty()) {
            void* rollbackObj = fnCreate(oldAlias.c_str(), lf.layer, lf.start, len);
            if (rollbackObj && edit->set_focus_object) edit->set_focus_object(rollbackObj);
        }
        return false;
    }

    if (!edit->move_object(newObj, lf.layer, lf.start)) {
        edit->delete_object(newObj);
        if (!oldAlias.empty()) {
            void* rollbackObj = fnCreate(oldAlias.c_str(), lf.layer, lf.start, len);
            if (rollbackObj && edit->set_focus_object) edit->set_focus_object(rollbackObj);
        }
        return false;
    }

    if (edit->set_focus_object) edit->set_focus_object(newObj);
    if (outNewObj) *outNewObj = newObj;
    return true;
}

static void ApplyTextStyleProc(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (TextStyleApplyContext*)param;
    if (!ctx || !edit) return;
    if (!edit->get_object_alias) {
        ctx->hasError = true;
        ctx->errorMessage = L"オブジェクト取得APIが利用できません。";
        return;
    }

    std::vector<void*> targets;
    if (edit->get_selected_object_num && edit->get_selected_object) {
        int n = edit->get_selected_object_num();
        for (int i = 0; i < n; ++i) {
            void* obj = edit->get_selected_object(i);
            if (obj) targets.push_back(obj);
        }
    }
    if (targets.empty() && edit->get_focus_object) {
        void* obj = edit->get_focus_object();
        if (obj) targets.push_back(obj);
    }
    if (targets.empty()) {
        ctx->hasError = true;
        ctx->errorMessage = L"タイムラインでテキストオブジェクトを選択してください。";
        return;
    }

    bool singleTarget = (targets.size() == 1);
    for (size_t i = 0; i < targets.size(); ++i) {
        void* obj = targets[i];
        LPCSTR raw = edit->get_object_alias(obj);
        if (!raw) {
            ctx->hasError = true;
            ctx->errorMessage = L"対象オブジェクトのデータ取得に失敗しました。";
            return;
        }

        std::string targetAlias = raw;
        if (!IsLikelySingleObjectAlias(targetAlias)) {
            ctx->hasError = true;
            ctx->errorMessage = L"対象は単体オブジェクトに限定されています。";
            return;
        }

        std::string newAlias;
        if (!BuildAliasWithAppliedTextStyle(targetAlias, ctx->styleAlias, newAlias)) {
            ctx->hasError = true;
            ctx->errorMessage = L"スタイルデータの解析に失敗しました。";
            return;
        }
        ApplyTextStyleInheritanceItems(targetAlias, newAlias);

        void* newObj = nullptr;
        if (!ReplaceObjectAliasByRecreate(edit, obj, newAlias, &newObj) || !newObj) {
            ctx->hasError = true;
            ctx->errorMessage = L"スタイル適用に失敗しました。";
            return;
        }

        ctx->appliedCount++;
        if (singleTarget && i == 0) {
            ctx->originalAlias = targetAlias;
            if (edit->get_object_layer_frame) {
                OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(newObj);
                ctx->pendingLayer = lf.layer;
                ctx->pendingStart = lf.start;
                ctx->pendingEnd = lf.end;
            }
        }
    }
    ctx->success = (ctx->appliedCount > 0);
}

static void RestorePendingTextStyleProc(void* param, EDIT_SECTION_OUT_SAFE* edit) {
    auto* ctx = (TextStyleApplyContext*)param;
    if (!ctx || !edit) return;
    if (ctx->originalAlias.empty()) return;
    if (g_textStylePendingLayer < 0 || g_textStylePendingStart < 0) return;

    void* targetObj = nullptr;
    auto matchByLocator = [&](void* obj) -> bool {
        if (!obj || !edit->get_object_layer_frame) return false;
        OBJECT_LAYER_FRAME_SAFE lf = edit->get_object_layer_frame(obj);
        if (lf.layer != g_textStylePendingLayer) return false;
        if (lf.start != g_textStylePendingStart) return false;
        return true;
    };

    if (edit->get_focus_object) {
        void* focus = edit->get_focus_object();
        if (matchByLocator(focus)) targetObj = focus;
    }
    if (!targetObj && edit->get_selected_object_num && edit->get_selected_object) {
        int n = edit->get_selected_object_num();
        for (int i = 0; i < n; ++i) {
            void* obj = edit->get_selected_object(i);
            if (matchByLocator(obj)) {
                targetObj = obj;
                break;
            }
        }
    }
    if (!targetObj) return;

    void* replacedObj = nullptr;
    if (!ReplaceObjectAliasByRecreate(edit, targetObj, ctx->originalAlias, &replacedObj) || !replacedObj) return;
    ctx->success = true;
}

static bool ApplyTextStyleFromAssetPath(HWND owner, const std::wstring& assetPath) {
    if (!g_enableTextStyle) return false;
    std::string styleAlias = ReadFileContent(assetPath);
    if (styleAlias.empty()) {
        ShowDarkMsg(owner, L"スタイルアセットの読み込みに失敗しました。", L"Error", MB_OK);
        return false;
    }

    TextStyleApplyContext ctx = {};
    ctx.styleAlias = styleAlias;
    if (!g_editHandle || !g_editHandle->call_edit_section_param || !g_editHandle->call_edit_section_param(&ctx, ApplyTextStyleProc)) {
        ShowDarkMsg(owner, L"編集セクション呼び出しに失敗しました。", L"Error", MB_OK);
        return false;
    }
    if (!ctx.success) {
        ShowDarkMsg(owner, ctx.hasError ? ctx.errorMessage.c_str() : L"スタイル適用に失敗しました。", L"Info", MB_OK);
        return false;
    }

    if (ctx.appliedCount == 1 && !ctx.originalAlias.empty() && ctx.pendingLayer >= 0 && ctx.pendingStart >= 0) {
        g_textStylePending = true;
        g_textStylePendingLayer = ctx.pendingLayer;
        g_textStylePendingStart = ctx.pendingStart;
        g_textStylePendingEnd = ctx.pendingEnd;
        g_textStyleOriginalAlias = ctx.originalAlias;
    } else {
        g_textStylePending = false;
        g_textStylePendingLayer = -1;
        g_textStylePendingStart = -1;
        g_textStylePendingEnd = -1;
        g_textStyleOriginalAlias.clear();
    }
    return true;
}

static bool CommitPendingTextStyle(HWND owner) {
    if (!g_textStylePending || g_textStylePendingLayer < 0 || g_textStylePendingStart < 0) return false;
    g_textStylePendingLayer = -1;
    g_textStylePendingStart = -1;
    g_textStylePendingEnd = -1;
    g_textStylePending = false;
    g_textStyleOriginalAlias.clear();
    (void)owner;
    return true;
}

static bool CancelPendingTextStyle(HWND owner) {
    if (!g_textStylePending || g_textStylePendingLayer < 0 || g_textStylePendingStart < 0) return false;
    TextStyleApplyContext ctx = {};
    ctx.originalAlias = g_textStyleOriginalAlias;
    if (!g_editHandle || !g_editHandle->call_edit_section_param || !g_editHandle->call_edit_section_param(&ctx, RestorePendingTextStyleProc) || !ctx.success) {
        if (owner) ShowDarkMsg(owner, L"スタイル取消に失敗しました。", L"Error", MB_OK);
        return false;
    }
    g_textStylePendingLayer = -1;
    g_textStylePendingStart = -1;
    g_textStylePendingEnd = -1;
    g_textStylePending = false;
    g_textStyleOriginalAlias.clear();
    return true;
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
        int delay = 4;
        ok = SaveAnimatedGifToPath(GetForcedPreviewOutputPath(), frames, delay);
    }

    for (auto* bmp : frames) delete bmp;
    return ok;
}

static OUTPUT_PLUGIN_TABLE g_outputPluginTable = {
    OUTPUT_PLUGIN_TABLE::FLAG_VIDEO,
    L"MyAsset Preview GIF",
    L"GIF File (*.gif)\0*.gif\0\0",
    L"My Asset Manager アセットプレビュー用GIF出力プラグイン",
    OutputSelectedRangeGif,
    nullptr,
    nullptr
};

static void CaptureRect(const RECT& rc, const std::wstring& savePath) {
    int w = rc.right - rc.left, h = rc.bottom - rc.top;
    if (w <= 0 || h <= 0) return;
    HDC hdcS = GetDC(NULL), hdcM = CreateCompatibleDC(hdcS);
    HBITMAP hB = CreateCompatibleBitmap(hdcS, w, h);
    HBITMAP hOld = (HBITMAP)SelectObject(hdcM, hB);
    BitBlt(hdcM, 0, 0, w, h, hdcS, rc.left, rc.top, SRCCOPY);
    Bitmap* bmp = new Bitmap(hB, NULL); CLSID clsid;
    GetEncoderClsid(L"image/png", &clsid); bmp->Save(savePath.c_str(), &clsid, NULL);
    delete bmp;
    SelectObject(hdcM, hOld);
    DeleteObject(hB);
    DeleteDC(hdcM);
    ReleaseDC(NULL, hdcS);
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
    if (needRepair > 0) msg += L"・複数オブジェクト(旧仕様)の修復：" + std::to_wstring(needRepair) + L" 個\n";
    if (needOptimize > 0) msg += L"・単体オブジェクトの最適化：" + std::to_wstring(needOptimize) + L" 個\n";
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
    RECT rc = dis->rcItem;
    int id = GetDlgCtrlID(dis->hwndItem);
    bool isPresetBtn = (id == ID_BTN_PRESET_A || id == ID_BTN_PRESET_B || id == ID_BTN_PRESET_C);
    bool isPresetActive = false;
    if (isPresetBtn) {
        int slot = (id == ID_BTN_PRESET_A) ? 0 : (id == ID_BTN_PRESET_B ? 1 : 2);
        isPresetActive = (slot == g_capturePresetSlot);
    }
    COLORREF bg = (dis->itemState & ODS_SELECTED) ? COL_BTN_PUSH : (isPresetActive ? COL_BTN_ACT : COL_BTN_BG);
    HBRUSH br = CreateSolidBrush(bg); 
    FillRect(dis->hDC, &rc, br); DeleteObject(br); FrameRect(dis->hDC, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH)); 
    wchar_t buf[256]; GetWindowTextW(dis->hwndItem, buf, 256); SetBkMode(dis->hDC, TRANSPARENT); SetTextColor(dis->hDC, COL_TEXT); SelectObject(dis->hDC, g_hFontUI); 
    DrawTextW(dis->hDC, buf, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE); 
}

static void DrawTitleBar(HDC hdc, int w, LPCWSTR title, bool showClose = true) { 
    RECT rt = {0, 0, w, TITLE_H}; HBRUSH br = CreateSolidBrush(COL_TITLE_BG); FillRect(hdc, &rt, br); DeleteObject(br); 
    SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontUI); 
    RECT rtx = {10, 0, w - 40, TITLE_H}; DrawTextW(hdc, title, -1, &rtx, DT_LEFT | DT_VCENTER | DT_SINGLELINE); 
    if (showClose) {
        RECT rc = {w - 40, 0, w, TITLE_H}; DrawTextW(hdc, L"✕", -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
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

static void OpenUrlInBrowser(HWND hwnd, const std::wstring& url) {
    HINSTANCE h = ShellExecuteW(hwnd, L"open", url.c_str(), NULL, NULL, SW_SHOWNORMAL);
    if ((INT_PTR)h <= 32) {
        ShowDarkMsg(hwnd, L"ブラウザでURLを開けませんでした。", L"Error", MB_OK);
    }
}

static void OpenReleaseNotesInBrowser(HWND hwnd) {
    std::wstring url = std::wstring(kGithubReleaseBase) + MYASSET_VERSION_W;
    OpenUrlInBrowser(hwnd, url);
}

static void OpenBugReportInBrowser(HWND hwnd) {
    OpenUrlInBrowser(hwnd, kGithubBugReportUrl);
}

static void AppendInfoSubMenu(HMENU parent) {
    HMENU infoSub = CreatePopupMenu();
    AppendMenuW(infoSub, MF_STRING, IDM_INFO_RELEASES, L"更新履歴");
    AppendMenuW(infoSub, MF_STRING, IDM_INFO_BUGREPORT, L"不具合報告");
    AppendMenuW(parent, MF_POPUP, (UINT_PTR)infoSub, L"情報");
}

static void ShowSettingsRefreshMenu(HWND hwnd, int xScreen, int yScreen) {
    HMENU hm = CreatePopupMenu();
    AppendMenuW(hm, MF_STRING, IDM_SETTINGS, L"設定");
    AppendInfoSubMenu(hm);
    AppendMenuW(hm, MF_STRING, IDM_OPEN_FOLDER, L"MyAssetフォルダを開く");
    TrackPopupMenu(hm, TPM_RIGHTBUTTON, xScreen, yScreen, 0, hwnd, NULL);
    DestroyMenu(hm);
}

static LRESULT DrawSpeedSliderCustom(NMCUSTOMDRAW* cd) {
    if (!cd || !cd->hdr.hwndFrom) return CDRF_DODEFAULT;

    HWND hSlider = cd->hdr.hwndFrom;
    HDC hdc = cd->hdc;
    if (cd->dwDrawStage != CDDS_PREPAINT) {
        return CDRF_DODEFAULT;
    }

    RECT rc = {};
    GetClientRect(hSlider, &rc);
    HBRUSH brBg = CreateSolidBrush(COL_BG);
    FillRect(hdc, &rc, brBg);
    DeleteObject(brBg);

    RECT ch = {};
    RECT th = {};
    SendMessageW(hSlider, TBM_GETCHANNELRECT, 0, (LPARAM)&ch);
    SendMessageW(hSlider, TBM_GETTHUMBRECT, 0, (LPARAM)&th);

    int thumbX = (th.left + th.right) / 2;
    if (thumbX < ch.left) thumbX = ch.left;
    if (thumbX > ch.right) thumbX = ch.right;

    HBRUSH brAll = CreateSolidBrush(RGB(220, 220, 220));
    FillRect(hdc, &ch, brAll);
    DeleteObject(brAll);

    RECT chCur = ch;
    chCur.right = thumbX;
    HBRUSH brCur = CreateSolidBrush(COL_BTN_ACT);
    FillRect(hdc, &chCur, brCur);
    DeleteObject(brCur);

    HBRUSH brThumb = CreateSolidBrush(RGB(30, 170, 240));
    FillRect(hdc, &th, brThumb);
    DeleteObject(brThumb);
    FrameRect(hdc, &th, (HBRUSH)GetStockObject(BLACK_BRUSH));

    return CDRF_SKIPDEFAULT;
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
    auto updateInheritGroupChecks = [&]() {
        for (int g = 0; g < kInheritGroupCount; ++g) {
            bool allOn = true;
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                if (!g_inheritItemEnabled[idx]) { allOn = false; break; }
            }
            SendMessageW(GetDlgItem(hwnd, IDC_CHK_INH_GROUP_BASE + g), BM_SETCHECK, allOn ? BST_CHECKED : BST_UNCHECKED, 0);
            SetWindowTextW(GetDlgItem(hwnd, IDC_BTN_INH_EXPAND_BASE + g), (g_inheritExpandedGroup == g) ? L"▲" : L"▼");
        }
    };

    auto layoutInheritControls = [&]() {
        int y = TITLE_H + 222;
        for (int g = 0; g < kInheritGroupCount; ++g) {
            HWND hGrp = GetDlgItem(hwnd, IDC_CHK_INH_GROUP_BASE + g);
            HWND hBtn = GetDlgItem(hwnd, IDC_BTN_INH_EXPAND_BASE + g);
            MoveWindow(hGrp, 155, y, 250, 22, TRUE);
            MoveWindow(hBtn, 410, y, 28, 22, TRUE);
            y += 24;

            bool expanded = (g_inheritExpandedGroup == g);
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                HWND hItem = GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + idx);
                if (expanded) {
                    MoveWindow(hItem, 175, y, 290, 20, TRUE);
                    y += 20;
                }
            }
        }
    };

    auto applyCategoryVisibility = [&](int cat) {
        g_settingsCategory = cat;
        ShowWindow(GetDlgItem(hwnd, IDC_CHK_ENABLE_TEXT_STYLE), (cat == 0) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10004), (cat == 0) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10005), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_INSERT_END), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_INSERT_AFTER_STD), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_CHK_CLEAR_EFFECTS_FROM2), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_CHK_HIDE_TEXTSTYLE_IN_LIST), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_ST_INHERIT_TITLE), (cat == 0 && g_enableTextStyle) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_SLIDER_SPEED), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_EDIT_SPEED), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10001), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10003), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_PREVIEW_ALL), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_PREVIEW_HOVER), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_PREVIEW_PERF_NORMAL), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_RAD_PREVIEW_PERF_LOW), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10006), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10007), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_SLIDER_PREVIEW_ZOOM), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_EDIT_PREVIEW_ZOOM), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_CHK_PREVIEW_THUMB_ONLY), (cat == 1) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, IDC_CHK_GIF_ORIGINAL), (cat == 2) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(hwnd, 10002), (cat == 2) ? SW_SHOW : SW_HIDE);
        for (int i = 0; i < kThemeColorCount; ++i) {
            ShowWindow(GetDlgItem(hwnd, IDC_ST_THEME_BASE + i), (cat == 3) ? SW_SHOW : SW_HIDE);
            ShowWindow(GetDlgItem(hwnd, IDC_EDIT_THEME_BASE + i), (cat == 3) ? SW_SHOW : SW_HIDE);
            ShowWindow(GetDlgItem(hwnd, IDC_BTN_PICK_THEME_BASE + i), (cat == 3) ? SW_SHOW : SW_HIDE);
        }
        ShowWindow(GetDlgItem(hwnd, IDC_BTN_THEME_RESET), (cat == 3) ? SW_SHOW : SW_HIDE);

        for (int g = 0; g < kInheritGroupCount; ++g) {
            bool showGrp = (cat == 0 && g_enableTextStyle);
            ShowWindow(GetDlgItem(hwnd, IDC_CHK_INH_GROUP_BASE + g), showGrp ? SW_SHOW : SW_HIDE);
            ShowWindow(GetDlgItem(hwnd, IDC_BTN_INH_EXPAND_BASE + g), showGrp ? SW_SHOW : SW_HIDE);
            bool expanded = (g_inheritExpandedGroup == g);
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                bool showItem = showGrp && expanded;
                ShowWindow(GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + idx), showItem ? SW_SHOW : SW_HIDE);
            }
        }
        updateInheritGroupChecks();
        layoutInheritControls();
        RedrawWindow(hwnd, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    };

    switch (msg) {
    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp); 
    case WM_NCACTIVATE: return TRUE;
    case WM_CREATE: {
        HWND hSide = CreateWindowExW(0, L"LISTBOX", L"", WS_VISIBLE|WS_CHILD|LBS_NOTIFY|WS_VSCROLL, 10, TITLE_H+12, 120, 470, hwnd, (HMENU)ID_SETTING_SIDEBAR, g_hInst, NULL);
        SendMessageW(hSide, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hSide, L"", L"");
        SendMessageW(hSide, LB_ADDSTRING, 0, (LPARAM)L"テキストスタイル");
        SendMessageW(hSide, LB_ADDSTRING, 0, (LPARAM)L"プレビュー設定");
        SendMessageW(hSide, LB_ADDSTRING, 0, (LPARAM)L"出力設定");
        SendMessageW(hSide, LB_ADDSTRING, 0, (LPARAM)L"外観");
        if (g_settingsCategory < 0 || g_settingsCategory > 3) g_settingsCategory = 0;
        SendMessageW(hSide, LB_SETCURSEL, g_settingsCategory, 0);

        HWND h4 = CreateWindowW(L"STATIC", L"表示機能", WS_VISIBLE|WS_CHILD, 155, TITLE_H+26, 320, 20, hwnd, (HMENU)10004, g_hInst, NULL); SendMessageW(h4, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hChkTs = CreateWindowW(L"BUTTON", L"テキストスタイル機能を使用する", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+56, 320, 24, hwnd, (HMENU)IDC_CHK_ENABLE_TEXT_STYLE, g_hInst, NULL);
        SendMessageW(hChkTs, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChkTs, L"", L"");
        SendMessageW(hChkTs, BM_SETCHECK, g_enableTextStyle ? BST_CHECKED : BST_UNCHECKED, 0);
        HWND hInsTitle = CreateWindowW(L"STATIC", L"挿入位置", WS_VISIBLE|WS_CHILD, 155, TITLE_H+88, 320, 20, hwnd, (HMENU)10005, g_hInst, NULL);
        SendMessageW(hInsTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hRadInsEnd = CreateWindowW(L"BUTTON", L"末尾に追加", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON|WS_GROUP, 155, TITLE_H+110, 180, 22, hwnd, (HMENU)IDC_RAD_INSERT_END, g_hInst, NULL);
        HWND hRadInsAfter = CreateWindowW(L"BUTTON", L"標準描画の直後に追加", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON, 155, TITLE_H+132, 220, 22, hwnd, (HMENU)IDC_RAD_INSERT_AFTER_STD, g_hInst, NULL);
        SendMessageW(hRadInsEnd, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SendMessageW(hRadInsAfter, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hRadInsEnd, L"", L"");
        SetWindowTheme(hRadInsAfter, L"", L"");
        SendMessageW(hRadInsEnd, BM_SETCHECK, g_textStyleInsertMode == 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(hRadInsAfter, BM_SETCHECK, g_textStyleInsertMode == 1 ? BST_CHECKED : BST_UNCHECKED, 0);

        HWND hChkClear = CreateWindowW(L"BUTTON", L"追加効果を削除してから適用", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+154, 280, 22, hwnd, (HMENU)IDC_CHK_CLEAR_EFFECTS_FROM2, g_hInst, NULL);
        SendMessageW(hChkClear, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChkClear, L"", L"");
        SendMessageW(hChkClear, BM_SETCHECK, g_textStyleClearEffectsFrom2 ? BST_CHECKED : BST_UNCHECKED, 0);

        HWND hChkHideInList = CreateWindowW(L"BUTTON", L"テキストスタイルを一覧から除外", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+178, 280, 22, hwnd, (HMENU)IDC_CHK_HIDE_TEXTSTYLE_IN_LIST, g_hInst, NULL);
        SendMessageW(hChkHideInList, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChkHideInList, L"", L"");
        SendMessageW(hChkHideInList, BM_SETCHECK, g_hideTextStyleInMainList ? BST_CHECKED : BST_UNCHECKED, 0);

        HWND hInheritTitle = CreateWindowW(L"STATIC", L"スタイル適用時に引き継ぐ項目", WS_VISIBLE|WS_CHILD, 155, TITLE_H+202, 320, 20, hwnd, (HMENU)IDC_ST_INHERIT_TITLE, g_hInst, NULL);
        SendMessageW(hInheritTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        for (int g = 0; g < kInheritGroupCount; ++g) {
            HWND hGrp = CreateWindowW(L"BUTTON", kInheritGroups[g].label, WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+116, 250, 22, hwnd, (HMENU)(INT_PTR)(IDC_CHK_INH_GROUP_BASE + g), g_hInst, NULL);
            HWND hBtn = CreateWindowW(L"BUTTON", L"▼", WS_VISIBLE|WS_CHILD, 410, TITLE_H+116, 28, 22, hwnd, (HMENU)(INT_PTR)(IDC_BTN_INH_EXPAND_BASE + g), g_hInst, NULL);
            SendMessageW(hGrp, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SendMessageW(hBtn, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SetWindowTheme(hGrp, L"", L"");
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                HWND hItem = CreateWindowW(L"BUTTON", kInheritItems[idx].label, WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 175, TITLE_H+116, 290, 20, hwnd, (HMENU)(INT_PTR)(IDC_CHK_INH_ITEM_BASE + idx), g_hInst, NULL);
                SendMessageW(hItem, WM_SETFONT, (WPARAM)g_hFontUI, 0);
                SetWindowTheme(hItem, L"", L"");
                SendMessageW(hItem, BM_SETCHECK, g_inheritItemEnabled[idx] ? BST_CHECKED : BST_UNCHECKED, 0);
            }
        }

        HWND h = CreateWindowW(L"STATIC", L"GIF再生速度倍率 (50% - 1000%)", WS_VISIBLE|WS_CHILD, 155, TITLE_H+26, 320, 20, hwnd, (HMENU)10001, g_hInst, NULL); SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hs = CreateWindowExW(0, TRACKBAR_CLASSW, L"", WS_VISIBLE|WS_CHILD|TBS_HORZ|TBS_AUTOTICKS, 155, TITLE_H+56, 250, 30, hwnd, (HMENU)IDC_SLIDER_SPEED, g_hInst, NULL);
        SendMessage(hs, TBM_SETRANGE, TRUE, MAKELONG(50, 1000)); SendMessage(hs, TBM_SETPOS, TRUE, g_gifSpeedPercent);
        SetWindowTheme(hs, L"", L"");
        HWND he = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", (std::to_wstring(g_gifSpeedPercent)).c_str(), WS_VISIBLE|WS_CHILD|ES_NUMBER|ES_CENTER|ES_AUTOHSCROLL, 415, TITLE_H+60, 60, 24, hwnd, (HMENU)IDC_EDIT_SPEED, g_hInst, NULL); SendMessageW(he, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(he, L"", L"");
        HWND h3 = CreateWindowW(L"STATIC", L"プレビュー再生", WS_VISIBLE|WS_CHILD, 155, TITLE_H+102, 320, 20, hwnd, (HMENU)10003, g_hInst, NULL); SendMessageW(h3, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hRadAll = CreateWindowW(L"BUTTON", L"一覧にカーソルがある間は一斉再生", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON|WS_GROUP, 155, TITLE_H+128, 320, 24, hwnd, (HMENU)IDC_RAD_PREVIEW_ALL, g_hInst, NULL);
        HWND hRadHover = CreateWindowW(L"BUTTON", L"ホバー中のアセットのみ再生", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON, 155, TITLE_H+154, 320, 24, hwnd, (HMENU)IDC_RAD_PREVIEW_HOVER, g_hInst, NULL);
        HWND hRadPerfNormal = CreateWindowW(L"BUTTON", L"通常プレビュー (25fps)", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON|WS_GROUP, 155, TITLE_H+222, 220, 24, hwnd, (HMENU)IDC_RAD_PREVIEW_PERF_NORMAL, g_hInst, NULL);
        HWND hRadPerfLow = CreateWindowW(L"BUTTON", L"低負荷プレビュー (10fps)", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON, 155, TITLE_H+246, 220, 24, hwnd, (HMENU)IDC_RAD_PREVIEW_PERF_LOW, g_hInst, NULL);
        SendMessageW(hRadAll, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SendMessageW(hRadHover, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SendMessageW(hRadPerfNormal, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SendMessageW(hRadPerfLow, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hRadAll, L"", L"");
        SetWindowTheme(hRadHover, L"", L"");
        SetWindowTheme(hRadPerfNormal, L"", L"");
        SetWindowTheme(hRadPerfLow, L"", L"");
        SendMessageW(hRadAll, BM_SETCHECK, (g_previewPlaybackMode == 1) ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(hRadHover, BM_SETCHECK, (g_previewPlaybackMode == 0) ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(hRadPerfNormal, BM_SETCHECK, (g_previewPerfMode == 0) ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(hRadPerfLow, BM_SETCHECK, (g_previewPerfMode == 1) ? BST_CHECKED : BST_UNCHECKED, 0);

        HWND hPerfTitle = CreateWindowW(L"STATIC", L"プレビュー負荷", WS_VISIBLE|WS_CHILD, 155, TITLE_H+200, 320, 20, hwnd, (HMENU)10006, g_hInst, NULL);
        SendMessageW(hPerfTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hZoomTitle = CreateWindowW(L"STATIC", L"表示倍率 (50% - 200%)", WS_VISIBLE|WS_CHILD, 155, TITLE_H+286, 320, 20, hwnd, (HMENU)10007, g_hInst, NULL);
        SendMessageW(hZoomTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hZoomSlider = CreateWindowExW(0, TRACKBAR_CLASSW, L"", WS_VISIBLE|WS_CHILD|TBS_HORZ|TBS_AUTOTICKS, 155, TITLE_H+314, 250, 30, hwnd, (HMENU)IDC_SLIDER_PREVIEW_ZOOM, g_hInst, NULL);
        SendMessage(hZoomSlider, TBM_SETRANGE, TRUE, MAKELONG(50, 200));
        SendMessage(hZoomSlider, TBM_SETPOS, TRUE, g_previewZoomPercent);
        SetWindowTheme(hZoomSlider, L"", L"");
        HWND hZoomEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", (std::to_wstring(g_previewZoomPercent)).c_str(), WS_VISIBLE|WS_CHILD|ES_NUMBER|ES_CENTER|ES_AUTOHSCROLL, 415, TITLE_H+318, 60, 24, hwnd, (HMENU)IDC_EDIT_PREVIEW_ZOOM, g_hInst, NULL);
        SendMessageW(hZoomEdit, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hZoomEdit, L"", L"");
        HWND hThumbOnly = CreateWindowW(L"BUTTON", L"サムネイルのみ表示", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+348, 290, 24, hwnd, (HMENU)IDC_CHK_PREVIEW_THUMB_ONLY, g_hInst, NULL);
        SendMessageW(hThumbOnly, WM_SETFONT, (WPARAM)g_hFontUI, TRUE);
        SetWindowTheme(hThumbOnly, L"", L"");
        SendMessageW(hThumbOnly, BM_SETCHECK, g_previewThumbOnly ? BST_CHECKED : BST_UNCHECKED, 0);

        HWND h2 = CreateWindowW(L"STATIC", L"GIF出力解像度", WS_VISIBLE|WS_CHILD, 155, TITLE_H+26, 320, 20, hwnd, (HMENU)10002, g_hInst, NULL); SendMessageW(h2, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hChk = CreateWindowW(L"BUTTON", L"元解像度で出力（OFF：最大幅480px）", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 155, TITLE_H+56, 320, 24, hwnd, (HMENU)IDC_CHK_GIF_ORIGINAL, g_hInst, NULL);
        SendMessageW(hChk, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChk, L"", L"");
        SendMessageW(hChk, BM_SETCHECK, g_gifExportKeepOriginal ? BST_CHECKED : BST_UNCHECKED, 0);

        for (int i = 0; i < kThemeColorCount; ++i) {
            int y = TITLE_H + 26 + i * 34;
            HWND hLbl = CreateWindowW(L"STATIC", kThemeColors[i].label, WS_VISIBLE|WS_CHILD, 155, y, 120, 22, hwnd, (HMENU)(INT_PTR)(IDC_ST_THEME_BASE + i), g_hInst, NULL);
            SendMessageW(hLbl, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", ColorToHex(*kThemeColors[i].value).c_str(), WS_VISIBLE|WS_CHILD|ES_AUTOHSCROLL, 280, y - 2, 120, 24, hwnd, (HMENU)(INT_PTR)(IDC_EDIT_THEME_BASE + i), g_hInst, NULL);
            SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SetWindowTheme(hEdit, L"", L"");
            HWND hPick = CreateWindowW(L"BUTTON", L"選択...", WS_VISIBLE|WS_CHILD, 408, y - 2, 70, 24, hwnd, (HMENU)(INT_PTR)(IDC_BTN_PICK_THEME_BASE + i), g_hInst, NULL);
            SendMessageW(hPick, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SetWindowTheme(hPick, L"", L"");
        }
        HWND hReset = CreateWindowW(L"BUTTON", L"デフォルトに戻す", WS_VISIBLE|WS_CHILD, 280, TITLE_H + 26 + kThemeColorCount * 34 + 8, 198, 28, hwnd, (HMENU)(INT_PTR)IDC_BTN_THEME_RESET, g_hInst, NULL);
        SendMessageW(hReset, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hReset, L"", L"");
        CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 390, 510, 90, 30, hwnd, (HMENU)ID_BTN_MSG_OK, g_hInst, NULL);
        applyCategoryVisibility(g_settingsCategory);
        return 0;
    }
    
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, L"設定");
        RECT split = {140, TITLE_H + 10, 141, rc.bottom - 10};
        HBRUSH br = CreateSolidBrush(COL_BORDER);
        FillRect(hdc, &split, br);
        DeleteObject(br);
        DrawWindowBorder(hwnd); EndPaint(hwnd, &ps); return 0;
    }
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wp;
        RECT rc = {};
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrBg);
        return 1;
    }
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
        HWND hSpeedSlider = GetDlgItem(hwnd, IDC_SLIDER_SPEED);
        HWND hZoomSlider = GetDlgItem(hwnd, IDC_SLIDER_PREVIEW_ZOOM);
        if ((HWND)lp == hSpeedSlider) {
            g_gifSpeedPercent = (int)SendMessage(hSpeedSlider, TBM_GETPOS, 0, 0);
            SetDlgItemInt(hwnd, IDC_EDIT_SPEED, g_gifSpeedPercent, FALSE);
            SaveConfig();
            InvalidateRect(hSpeedSlider, NULL, FALSE);
        } else if ((HWND)lp == hZoomSlider) {
            g_previewZoomPercent = (int)SendMessage(hZoomSlider, TBM_GETPOS, 0, 0);
            if (g_previewZoomPercent < 50) g_previewZoomPercent = 50;
            if (g_previewZoomPercent > 200) g_previewZoomPercent = 200;
            SetDlgItemInt(hwnd, IDC_EDIT_PREVIEW_ZOOM, g_previewZoomPercent, FALSE);
            SaveConfig();
            UpdateDisplayList();
            InvalidateRect(hZoomSlider, NULL, FALSE);
        } else {
            return 0;
        }
        return 0; 
    }
    case WM_NOTIFY: {
        NMHDR* hdr = (NMHDR*)lp;
        if (hdr && (hdr->idFrom == IDC_SLIDER_SPEED || hdr->idFrom == IDC_SLIDER_PREVIEW_ZOOM) && hdr->code == NM_CUSTOMDRAW) {
            return DrawSpeedSliderCustom((NMCUSTOMDRAW*)lp);
        }
        return 0;
    }
    case WM_COMMAND: {
        if (LOWORD(wp) == IDC_EDIT_SPEED && HIWORD(wp) == EN_CHANGE) {
            if (GetFocus() == GetDlgItem(hwnd, IDC_EDIT_SPEED)) {
                int val = GetDlgItemInt(hwnd, IDC_EDIT_SPEED, NULL, FALSE);
                if (val < 50) val = 50; if (val > 1000) val = 1000;
                g_gifSpeedPercent = val;
                HWND hSlider = GetDlgItem(hwnd, IDC_SLIDER_SPEED);
                SendMessage(hSlider, TBM_SETPOS, TRUE, val);
                SaveConfig();
                InvalidateRect(hSlider, NULL, FALSE);
            }
        }
        else if (LOWORD(wp) == IDC_EDIT_PREVIEW_ZOOM && HIWORD(wp) == EN_CHANGE) {
            if (GetFocus() == GetDlgItem(hwnd, IDC_EDIT_PREVIEW_ZOOM)) {
                int val = GetDlgItemInt(hwnd, IDC_EDIT_PREVIEW_ZOOM, NULL, FALSE);
                if (val < 50) val = 50; if (val > 200) val = 200;
                g_previewZoomPercent = val;
                HWND hSlider = GetDlgItem(hwnd, IDC_SLIDER_PREVIEW_ZOOM);
                SendMessage(hSlider, TBM_SETPOS, TRUE, val);
                SaveConfig();
                UpdateDisplayList();
                InvalidateRect(hSlider, NULL, FALSE);
            }
        }
        else if (LOWORD(wp) == IDC_CHK_GIF_ORIGINAL && HIWORD(wp) == BN_CLICKED) {
            g_gifExportKeepOriginal = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_GIF_ORIGINAL), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
        }
        else if (LOWORD(wp) == IDC_CHK_PREVIEW_THUMB_ONLY && HIWORD(wp) == BN_CLICKED) {
            g_previewThumbOnly = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_PREVIEW_THUMB_ONLY), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
            UpdateDisplayList();
        }
        else if (LOWORD(wp) == IDC_CHK_ENABLE_TEXT_STYLE && HIWORD(wp) == BN_CLICKED) {
            g_enableTextStyle = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_ENABLE_TEXT_STYLE), BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (!g_enableTextStyle) g_assetViewTab = 0;
            SaveConfig();
            BroadcastTextStyleConfigChanged();
            if (g_hDlg && IsWindow(g_hDlg)) {
                HWND hTs = GetDlgItem(g_hDlg, IDC_CHK_TEXT_STYLE);
                if (hTs && IsWindow(hTs)) ShowWindow(hTs, g_enableTextStyle ? SW_SHOW : SW_HIDE);
                if (!g_enableTextStyle) g_addAsTextStyle = false;
            }
            UpdateDisplayList();
            applyCategoryVisibility(g_settingsCategory);
        }
        else if (LOWORD(wp) == IDC_RAD_INSERT_END && HIWORD(wp) == BN_CLICKED) {
            g_textStyleInsertMode = 0;
            SaveConfig();
            BroadcastTextStyleConfigChanged();
        }
        else if (LOWORD(wp) == IDC_RAD_INSERT_AFTER_STD && HIWORD(wp) == BN_CLICKED) {
            g_textStyleInsertMode = 1;
            SaveConfig();
            BroadcastTextStyleConfigChanged();
        }
        else if (LOWORD(wp) == IDC_CHK_CLEAR_EFFECTS_FROM2 && HIWORD(wp) == BN_CLICKED) {
            g_textStyleClearEffectsFrom2 = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_CLEAR_EFFECTS_FROM2), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
            BroadcastTextStyleConfigChanged();
        }
        else if (LOWORD(wp) == IDC_CHK_HIDE_TEXTSTYLE_IN_LIST && HIWORD(wp) == BN_CLICKED) {
            g_hideTextStyleInMainList = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_HIDE_TEXTSTYLE_IN_LIST), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
            UpdateDisplayList();
            BroadcastTextStyleConfigChanged();
        }
        else if (LOWORD(wp) == IDC_RAD_PREVIEW_ALL && HIWORD(wp) == BN_CLICKED) {
            g_previewPlaybackMode = 1;
            SaveConfig();
        }
        else if (LOWORD(wp) == IDC_RAD_PREVIEW_HOVER && HIWORD(wp) == BN_CLICKED) {
            g_previewPlaybackMode = 0;
            SaveConfig();
            if (g_hwnd && g_hoverIndex == -1) KillTimer(g_hwnd, ID_TIMER_HOVER);
            g_previewAllLastTick = 0;
            g_previewHoverLastTick = 0;
        }
        else if (LOWORD(wp) == IDC_RAD_PREVIEW_PERF_NORMAL && HIWORD(wp) == BN_CLICKED) {
            g_previewPerfMode = 0;
            SaveConfig();
            g_previewAllLastTick = 0;
            g_previewHoverLastTick = 0;
            if (g_hwnd) {
                if (g_previewPlaybackMode == 1 && g_hoverIndex != -1) {
                    SetTimer(g_hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
                } else if (g_previewPlaybackMode == 0 && g_hoverIndex != -1) {
                    SetTimer(g_hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
                }
            }
        }
        else if (LOWORD(wp) == IDC_RAD_PREVIEW_PERF_LOW && HIWORD(wp) == BN_CLICKED) {
            g_previewPerfMode = 1;
            SaveConfig();
            g_previewAllLastTick = 0;
            g_previewHoverLastTick = 0;
            if (g_hwnd) {
                if (g_previewPlaybackMode == 1 && g_hoverIndex != -1) {
                    SetTimer(g_hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
                } else if (g_previewPlaybackMode == 0 && g_hoverIndex != -1) {
                    SetTimer(g_hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
                }
            }
        }
        else if (LOWORD(wp) == ID_SETTING_SIDEBAR && HIWORD(wp) == LBN_SELCHANGE) {
            int sel = (int)SendMessageW(GetDlgItem(hwnd, ID_SETTING_SIDEBAR), LB_GETCURSEL, 0, 0);
            if (sel < 0) sel = 0; if (sel > 3) sel = 3;
            applyCategoryVisibility(sel);
        }
        else if (LOWORD(wp) >= IDC_BTN_PICK_THEME_BASE && LOWORD(wp) < IDC_BTN_PICK_THEME_BASE + kThemeColorCount && HIWORD(wp) == BN_CLICKED) {
            int idx = LOWORD(wp) - IDC_BTN_PICK_THEME_BASE;
            CHOOSECOLORW cc = {};
            cc.lStructSize = sizeof(cc);
            cc.hwndOwner = hwnd;
            cc.rgbResult = *kThemeColors[idx].value;
            cc.lpCustColors = g_colorPickerCustom;
            cc.Flags = CC_FULLOPEN | CC_RGBINIT;
            if (ChooseColorW(&cc)) {
                *kThemeColors[idx].value = cc.rgbResult;
                RecreateUiBrushes();
                SetWindowTextW(GetDlgItem(hwnd, IDC_EDIT_THEME_BASE + idx), ColorToHex(*kThemeColors[idx].value).c_str());
                SaveConfig();
                BroadcastTextStyleConfigChanged();
                InvalidateRect(hwnd, NULL, FALSE);
            }
        }
        else if (LOWORD(wp) >= IDC_EDIT_THEME_BASE && LOWORD(wp) < IDC_EDIT_THEME_BASE + kThemeColorCount && HIWORD(wp) == EN_KILLFOCUS) {
            int idx = LOWORD(wp) - IDC_EDIT_THEME_BASE;
            wchar_t buf[64] = {};
            GetWindowTextW(GetDlgItem(hwnd, LOWORD(wp)), buf, 64);
            COLORREF parsed = 0;
            if (TryParseHexColor(buf, parsed)) {
                *kThemeColors[idx].value = parsed;
                RecreateUiBrushes();
                SetWindowTextW(GetDlgItem(hwnd, LOWORD(wp)), ColorToHex(*kThemeColors[idx].value).c_str());
                SaveConfig();
                BroadcastTextStyleConfigChanged();
                InvalidateRect(hwnd, NULL, FALSE);
            } else {
                SetWindowTextW(GetDlgItem(hwnd, LOWORD(wp)), ColorToHex(*kThemeColors[idx].value).c_str());
            }
        }
        else if (LOWORD(wp) == IDC_BTN_THEME_RESET && HIWORD(wp) == BN_CLICKED) {
            ResetThemeColorsToDefault();
            RecreateUiBrushes();
            for (int i = 0; i < kThemeColorCount; ++i) {
                SetWindowTextW(GetDlgItem(hwnd, IDC_EDIT_THEME_BASE + i), ColorToHex(*kThemeColors[i].value).c_str());
            }
            SaveConfig();
            BroadcastTextStyleConfigChanged();
            InvalidateRect(hwnd, NULL, FALSE);
        }
        else if (LOWORD(wp) >= IDC_CHK_INH_GROUP_BASE && LOWORD(wp) < IDC_CHK_INH_GROUP_BASE + kInheritGroupCount && HIWORD(wp) == BN_CLICKED) {
            int g = LOWORD(wp) - IDC_CHK_INH_GROUP_BASE;
            bool on = (SendMessageW(GetDlgItem(hwnd, LOWORD(wp)), BM_GETCHECK, 0, 0) == BST_CHECKED);
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                g_inheritItemEnabled[idx] = on;
                SendMessageW(GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + idx), BM_SETCHECK, on ? BST_CHECKED : BST_UNCHECKED, 0);
            }
            SaveConfig();
            BroadcastTextStyleConfigChanged();
            applyCategoryVisibility(g_settingsCategory);
        }
        else if (LOWORD(wp) >= IDC_BTN_INH_EXPAND_BASE && LOWORD(wp) < IDC_BTN_INH_EXPAND_BASE + kInheritGroupCount && HIWORD(wp) == BN_CLICKED) {
            int g = LOWORD(wp) - IDC_BTN_INH_EXPAND_BASE;
            g_inheritExpandedGroup = (g_inheritExpandedGroup == g) ? -1 : g;
            applyCategoryVisibility(g_settingsCategory);
        }
        else if (LOWORD(wp) >= IDC_CHK_INH_ITEM_BASE && LOWORD(wp) < IDC_CHK_INH_ITEM_BASE + kInheritItemCount && HIWORD(wp) == BN_CLICKED) {
            int idx = LOWORD(wp) - IDC_CHK_INH_ITEM_BASE;
            g_inheritItemEnabled[idx] = (SendMessageW(GetDlgItem(hwnd, LOWORD(wp)), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig();
            BroadcastTextStyleConfigChanged();
            applyCategoryVisibility(g_settingsCategory);
        }
        else if (LOWORD(wp) == ID_BTN_MSG_OK) DestroyWindow(hwnd); 
        return 0;
    }
    case WM_APP + 21: {
        SendMessageW(GetDlgItem(hwnd, IDC_CHK_ENABLE_TEXT_STYLE), BM_SETCHECK, g_enableTextStyle ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_INSERT_END), BM_SETCHECK, g_textStyleInsertMode == 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_INSERT_AFTER_STD), BM_SETCHECK, g_textStyleInsertMode == 1 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_PREVIEW_ALL), BM_SETCHECK, g_previewPlaybackMode == 1 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_PREVIEW_HOVER), BM_SETCHECK, g_previewPlaybackMode == 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_PREVIEW_PERF_NORMAL), BM_SETCHECK, g_previewPerfMode == 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_PREVIEW_PERF_LOW), BM_SETCHECK, g_previewPerfMode == 1 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_SLIDER_PREVIEW_ZOOM), TBM_SETPOS, TRUE, g_previewZoomPercent);
        SetDlgItemInt(hwnd, IDC_EDIT_PREVIEW_ZOOM, g_previewZoomPercent, FALSE);
        SendMessageW(GetDlgItem(hwnd, IDC_CHK_PREVIEW_THUMB_ONLY), BM_SETCHECK, g_previewThumbOnly ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_CHK_CLEAR_EFFECTS_FROM2), BM_SETCHECK, g_textStyleClearEffectsFrom2 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_CHK_HIDE_TEXTSTYLE_IN_LIST), BM_SETCHECK, g_hideTextStyleInMainList ? BST_CHECKED : BST_UNCHECKED, 0);
        for (int i = 0; i < kThemeColorCount; ++i) {
            SetWindowTextW(GetDlgItem(hwnd, IDC_EDIT_THEME_BASE + i), ColorToHex(*kThemeColors[i].value).c_str());
        }
        for (int i = 0; i < kInheritItemCount; ++i) {
            SendMessageW(GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + i), BM_SETCHECK, g_inheritItemEnabled[i] ? BST_CHECKED : BST_UNCHECKED, 0);
        }
        applyCategoryVisibility(g_settingsCategory);
        return 0;
    }
    case WM_LBUTTONUP: { POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; RECT rc; GetClientRect(hwnd, &rc); if (pt.y < TITLE_H && pt.x > rc.right - 40) { DestroyWindow(hwnd); } return 0; }
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_INPUT_BG);
        return (LRESULT)g_hBrInputBg;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, COL_TEXT);
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)GetStockObject(NULL_BRUSH);
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_BG);
        return (LRESULT)g_hBrBg;
    }
    case WM_DESTROY: g_hSettingDlg = nullptr; return 0;
    } return DefWindowProc(hwnd, msg, wp, lp);
}
void OpenSettings() {
    if (g_hSettingDlg) { SetForegroundWindow(g_hSettingDlg); return; }
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = SettingsDlgProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"MyAsset_Settings";
    wc.hbrBackground = g_hBrBg;
    RegisterClassW(&wc);
    int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    g_hSettingDlg = CreateWindowExW(
        WS_EX_TOPMOST|WS_EX_TOOLWINDOW,
        L"MyAsset_Settings",
        L"設定",
        WS_POPUP|WS_THICKFRAME,
        (sw-500)/2, (sh-560)/2, 500, 560,
        g_hwnd, nullptr, g_hInst, nullptr);
    if (g_hSettingDlg && IsWindow(g_hSettingDlg)) {
        RedrawWindow(g_hSettingDlg, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        ShowWindow(g_hSettingDlg, SW_SHOW);
    }
}

static LRESULT CALLBACK TextStyleQuickDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    auto updateInheritGroupChecks = [&]() {
        for (int g = 0; g < kInheritGroupCount; ++g) {
            bool allOn = true;
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                if (!g_inheritItemEnabled[idx]) { allOn = false; break; }
            }
            SendMessageW(GetDlgItem(hwnd, IDC_CHK_INH_GROUP_BASE + g), BM_SETCHECK, allOn ? BST_CHECKED : BST_UNCHECKED, 0);
            SetWindowTextW(GetDlgItem(hwnd, IDC_BTN_INH_EXPAND_BASE + g), (g_quickInheritExpandedGroup == g) ? L"▲" : L"▼");
        }
    };
    auto layoutInheritControls = [&]() {
        int y = TITLE_H + 136;
        for (int g = 0; g < kInheritGroupCount; ++g) {
            MoveWindow(GetDlgItem(hwnd, IDC_CHK_INH_GROUP_BASE + g), 20, y, 250, 22, TRUE);
            MoveWindow(GetDlgItem(hwnd, IDC_BTN_INH_EXPAND_BASE + g), 275, y, 28, 22, TRUE);
            y += 24;
            bool expanded = (g_quickInheritExpandedGroup == g);
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                HWND hItem = GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + idx);
                if (expanded) {
                    MoveWindow(hItem, 40, y, 320, 20, TRUE);
                    y += 20;
                }
                ShowWindow(hItem, expanded ? SW_SHOW : SW_HIDE);
            }
        }
    };
    auto refreshControls = [&]() {
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_INSERT_END), BM_SETCHECK, g_textStyleInsertMode == 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_RAD_INSERT_AFTER_STD), BM_SETCHECK, g_textStyleInsertMode == 1 ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(GetDlgItem(hwnd, IDC_CHK_CLEAR_EFFECTS_FROM2), BM_SETCHECK, g_textStyleClearEffectsFrom2 ? BST_CHECKED : BST_UNCHECKED, 0);
        for (int i = 0; i < kInheritItemCount; ++i) {
            SendMessageW(GetDlgItem(hwnd, IDC_CHK_INH_ITEM_BASE + i), BM_SETCHECK, g_inheritItemEnabled[i] ? BST_CHECKED : BST_UNCHECKED, 0);
        }
        updateInheritGroupChecks();
        layoutInheritControls();
        RedrawWindow(hwnd, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    };

    switch (msg) {
    case WM_NCCALCSIZE: if (wp) return 0; return DefWindowProc(hwnd, msg, wp, lp);
    case WM_NCACTIVATE: return TRUE;
    case WM_CREATE: {
        HWND h = CreateWindowW(L"STATIC", L"挿入位置", WS_VISIBLE|WS_CHILD, 20, TITLE_H + 18, 320, 20, hwnd, (HMENU)10005, g_hInst, NULL);
        SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        HWND hRadInsEnd = CreateWindowW(L"BUTTON", L"末尾に追加", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON|WS_GROUP, 20, TITLE_H+40, 170, 22, hwnd, (HMENU)IDC_RAD_INSERT_END, g_hInst, NULL);
        HWND hRadInsAfter = CreateWindowW(L"BUTTON", L"標準描画の直後に追加", WS_VISIBLE|WS_CHILD|BS_AUTORADIOBUTTON, 20, TITLE_H+62, 220, 22, hwnd, (HMENU)IDC_RAD_INSERT_AFTER_STD, g_hInst, NULL);
        SendMessageW(hRadInsEnd, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SendMessageW(hRadInsAfter, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hRadInsEnd, L"", L"");
        SetWindowTheme(hRadInsAfter, L"", L"");

        HWND hChkClear = CreateWindowW(L"BUTTON", L"追加効果を削除してから適用", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 20, TITLE_H+86, 280, 22, hwnd, (HMENU)IDC_CHK_CLEAR_EFFECTS_FROM2, g_hInst, NULL);
        SendMessageW(hChkClear, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        SetWindowTheme(hChkClear, L"", L"");

        h = CreateWindowW(L"STATIC", L"スタイル適用時に引き継ぐ項目", WS_VISIBLE|WS_CHILD, 20, TITLE_H + 114, 320, 20, hwnd, (HMENU)IDC_ST_INHERIT_TITLE, g_hInst, NULL);
        SendMessageW(h, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        for (int g = 0; g < kInheritGroupCount; ++g) {
            HWND hGrp = CreateWindowW(L"BUTTON", kInheritGroups[g].label, WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 20, TITLE_H + 112, 250, 22, hwnd, (HMENU)(INT_PTR)(IDC_CHK_INH_GROUP_BASE + g), g_hInst, NULL);
            HWND hBtn = CreateWindowW(L"BUTTON", L"▼", WS_VISIBLE|WS_CHILD, 275, TITLE_H + 112, 28, 22, hwnd, (HMENU)(INT_PTR)(IDC_BTN_INH_EXPAND_BASE + g), g_hInst, NULL);
            SendMessageW(hGrp, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SendMessageW(hBtn, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SetWindowTheme(hGrp, L"", L"");
            for (int i = 0; i < kInheritGroups[g].count; ++i) {
                int idx = kInheritGroups[g].start + i;
                HWND hItem = CreateWindowW(L"BUTTON", kInheritItems[idx].label, WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 40, TITLE_H + 112, 320, 20, hwnd, (HMENU)(INT_PTR)(IDC_CHK_INH_ITEM_BASE + idx), g_hInst, NULL);
                SendMessageW(hItem, WM_SETFONT, (WPARAM)g_hFontUI, 0);
                SetWindowTheme(hItem, L"", L"");
            }
        }
        CreateWindowW(L"BUTTON", L"閉じる", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 280, 510, 90, 30, hwnd, (HMENU)ID_BTN_MSG_OK, g_hInst, NULL);
        refreshControls();
        return 0;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, L"テキストスタイル詳細設定");
        DrawWindowBorder(hwnd); EndPaint(hwnd, &ps); return 0;
    }
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wp;
        RECT rc = {};
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrBg);
        return 1;
    }
    case WM_DRAWITEM: DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_NCHITTEST: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; ScreenToClient(hwnd, &pt); RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H) { if (pt.x > rc.right - 40) return HTCLIENT; return HTCAPTION; }
        return HTCLIENT;
    }
    case WM_COMMAND: {
        int id = LOWORD(wp), code = HIWORD(wp);
        if (id == ID_BTN_MSG_OK) { DestroyWindow(hwnd); return 0; }
        if (id == IDC_RAD_INSERT_END && code == BN_CLICKED) { g_textStyleInsertMode = 0; SaveConfig(); BroadcastTextStyleConfigChanged(); return 0; }
        if (id == IDC_RAD_INSERT_AFTER_STD && code == BN_CLICKED) { g_textStyleInsertMode = 1; SaveConfig(); BroadcastTextStyleConfigChanged(); return 0; }
        if (id == IDC_CHK_CLEAR_EFFECTS_FROM2 && code == BN_CLICKED) { g_textStyleClearEffectsFrom2 = (SendMessageW(GetDlgItem(hwnd, id), BM_GETCHECK, 0, 0) == BST_CHECKED); SaveConfig(); BroadcastTextStyleConfigChanged(); return 0; }
        if (id >= IDC_CHK_INH_GROUP_BASE && id < IDC_CHK_INH_GROUP_BASE + kInheritGroupCount && code == BN_CLICKED) {
            int g = id - IDC_CHK_INH_GROUP_BASE;
            bool on = (SendMessageW(GetDlgItem(hwnd, id), BM_GETCHECK, 0, 0) == BST_CHECKED);
            for (int i = 0; i < kInheritGroups[g].count; ++i) g_inheritItemEnabled[kInheritGroups[g].start + i] = on;
            SaveConfig(); BroadcastTextStyleConfigChanged(); refreshControls(); return 0;
        }
        if (id >= IDC_BTN_INH_EXPAND_BASE && id < IDC_BTN_INH_EXPAND_BASE + kInheritGroupCount && code == BN_CLICKED) {
            int g = id - IDC_BTN_INH_EXPAND_BASE;
            g_quickInheritExpandedGroup = (g_quickInheritExpandedGroup == g) ? -1 : g;
            refreshControls(); return 0;
        }
        if (id >= IDC_CHK_INH_ITEM_BASE && id < IDC_CHK_INH_ITEM_BASE + kInheritItemCount && code == BN_CLICKED) {
            int idx = id - IDC_CHK_INH_ITEM_BASE;
            g_inheritItemEnabled[idx] = (SendMessageW(GetDlgItem(hwnd, id), BM_GETCHECK, 0, 0) == BST_CHECKED);
            SaveConfig(); BroadcastTextStyleConfigChanged(); refreshControls(); return 0;
        }
        return 0;
    }
    case WM_APP + 21: refreshControls(); return 0;
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_INPUT_BG);
        return (LRESULT)g_hBrInputBg;
    }
    case WM_CTLCOLORSTATIC: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkMode(hdc, TRANSPARENT); SetBkColor(hdc, COL_BG); return (LRESULT)g_hBrBg; }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wp;
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_BG);
        return (LRESULT)g_hBrBg;
    }
    case WM_LBUTTONUP: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) }; RECT rc; GetClientRect(hwnd, &rc);
        if (pt.y < TITLE_H && pt.x > rc.right - 40) DestroyWindow(hwnd);
        return 0;
    }
    case WM_DESTROY: g_hTextStyleQuickDlg = nullptr; return 0;
    }
    return DefWindowProc(hwnd, msg, wp, lp);
}

static void OpenTextStyleQuickDialog(HWND parent) {
    if (g_hTextStyleQuickDlg && IsWindow(g_hTextStyleQuickDlg)) {
        SetForegroundWindow(g_hTextStyleQuickDlg);
        return;
    }
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = TextStyleQuickDlgProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"MyAsset_TextStyleQuick";
    wc.hbrBackground = g_hBrBg;
    RegisterClassW(&wc);
    int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    g_hTextStyleQuickDlg = CreateWindowExW(
        WS_EX_TOPMOST|WS_EX_TOOLWINDOW,
        L"MyAsset_TextStyleQuick",
        L"テキストスタイル詳細設定",
        WS_POPUP|WS_THICKFRAME,
        (sw-400)/2, (sh-560)/2, 400, 560,
        parent ? parent : g_hwnd, nullptr, g_hInst, nullptr);
    if (g_hTextStyleQuickDlg && IsWindow(g_hTextStyleQuickDlg)) {
        RedrawWindow(g_hTextStyleQuickDlg, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        ShowWindow(g_hTextStyleQuickDlg, SW_SHOW);
    }
}

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
                    a.isTextStyle = (g_textStylePaths.count(a.path) > 0);
                    
                    std::string head = ReadFileHead(a.path, 100);
                    a.isMulti = (head.find("[0]") != std::string::npos);

                    a.pImage = nullptr; a.frameCount = 0; a.frameDelays = nullptr; a.currentFrame = 0; a.previewCarryMs = 0.0; a.previewCycleMs = 0; a.pPropertyItem = nullptr;
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
                                        a.frameDelays = new UINT[a.frameCount];
                                        unsigned long long cycleSum = 0;
                                        for(UINT i=0; i<a.frameCount; i++) {
                                            long d = ((long*)a.pPropertyItem->value)[i];
                                            UINT delayMs = (d <= 0) ? 40u : (UINT)d * 10u;
                                            a.frameDelays[i] = delayMs;
                                            cycleSum += (unsigned long long)delayMs;
                                        }
                                        if (cycleSum == 0) cycleSum = (unsigned long long)a.frameCount * 100ULL;
                                        a.previewCycleMs = (cycleSum > 0xFFFFFFFFULL) ? 0xFFFFFFFFu : (UINT)cycleSum;
                                    }
                                }
                            }
                        }
                    }
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
        if (g_enableTextStyle && g_assetViewTab == 1 && !a.isTextStyle) continue;
        if (g_enableTextStyle && g_assetViewTab == 0 && g_hideTextStyleInMainList && a.isTextStyle) continue;
        if ((g_currentCategory == L"ALL" || a.category == g_currentCategory) && (sl.empty() || ToLower(a.name).find(sl) != std::wstring::npos)) 
            g_displayAssets.push_back(&a); 
    }

    std::map<std::wstring, int> customRank;
    if (g_sortMode == 3) {
        EnsureCustomOrderReadyForCurrentCategory();
        auto it = g_customOrderByCategory.find(g_currentCategory);
        if (it != g_customOrderByCategory.end()) {
            for (int i = 0; i < (int)it->second.size(); ++i) customRank[it->second[(size_t)i]] = i;
        }
    }

    std::sort(g_displayAssets.begin(), g_displayAssets.end(), [&customRank](Asset* a, Asset* b) {
        if (g_sortMode == 1) {
            if (a->isFavorite != b->isFavorite) return a->isFavorite > b->isFavorite;
            return a->name < b->name;
        }
        if (g_sortMode == 2) {
            if (a->category != b->category) return a->category < b->category;
            return a->name < b->name;
        }
        if (g_sortMode == 3) {
            std::wstring ka = GetAssetOrderKey(g_currentCategory, a->path);
            std::wstring kb = GetAssetOrderKey(g_currentCategory, b->path);
            int ra = INT_MAX, rb = INT_MAX;
            auto ia = customRank.find(ka);
            if (ia != customRank.end()) ra = ia->second;
            auto ib = customRank.find(kb);
            if (ib != customRank.end()) rb = ib->second;
            if (ra != rb) return ra < rb;
            return a->name < b->name;
        }
        return a->name < b->name;
    });
    if (g_hwnd) {
        UpdateScrollBar(g_hwnd);
        InvalidateRect(g_hwnd, nullptr, FALSE);
    }
}

static void SaveCurrentDisplayOrderForCategory() {
    if (g_currentCategory.empty()) return;
    std::vector<std::wstring> order;
    order.reserve(g_displayAssets.size());
    for (auto* a : g_displayAssets) {
        if (!a) continue;
        if (g_currentCategory != L"ALL" && a->category != g_currentCategory) continue;
        order.push_back(GetAssetOrderKey(g_currentCategory, a->path));
    }
    g_customOrderByCategory[g_currentCategory] = order;
    SaveCustomOrderForCategory(g_currentCategory, order);
}

void RefreshAssets(bool reloadFav) { 
    std::wstring prevCategory = g_currentCategory;
    InitBaseDir(); if (reloadFav) LoadFavorites(); 
    LoadFixedFrames(); 
    LoadTextStyleFlags();
    LoadConfig(); 
    if (!g_lastTempPath.empty()) { DeleteFileW(g_lastTempPath.c_str()); g_lastTempPath = L""; }
    g_customOrderByCategory.clear();
    ClearAssets(); ScanDirectory(g_baseDir, L"Main"); 
    g_categories.clear(); std::set<std::wstring> cs; for(auto& a : g_assets) cs.insert(a.category);
    ReconcileCustomOrderForCategory(L"ALL");
    for (const auto& c : cs) ReconcileCustomOrderForCategory(c);
    g_categories.push_back(L"ALL");
    for(auto& c : cs) g_categories.push_back(c);

    if (std::find(g_categories.begin(), g_categories.end(), prevCategory) != g_categories.end()) g_currentCategory = prevCategory;
    else g_currentCategory = L"ALL";

    if(g_hCombo) {
        SendMessage(g_hCombo, CB_RESETCONTENT, 0, 0);
        SendMessage(g_hCombo, CB_ADDSTRING, 0, (LPARAM)L"すべて (ALL)");
        for(auto& c : cs) SendMessage(g_hCombo, CB_ADDSTRING, 0, (LPARAM)c.c_str());
        int selIndex = 0;
        for (int i = 0; i < (int)g_categories.size(); ++i) {
            if (g_categories[i] == g_currentCategory) { selIndex = i; break; }
        }
        SendMessage(g_hCombo, CB_SETCURSEL, selIndex, 0);
    }
    UpdateDisplayList(); 
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
            g_lastCapturedRect = rc;
            g_lastCapturedRectValid = true;
            g_isImageRemoved = false; UpdateAddPreviewImage(); 
            if (g_hwnd) ShowWindow(g_hwnd, SW_SHOW);
            if (g_hDlg) { ShowWindow(g_hDlg, SW_SHOW); SetForegroundWindow(g_hDlg); UpdateAddCapturePresetUi(g_hDlg); }
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

static RECT GetAddPreviewRect() {
    RECT rp = { 75, TITLE_H + 204, 335, TITLE_H + 350 };
    return rp;
}

static void UpdateAddCapturePresetUi(HWND hwnd) {
    if (!hwnd) return;
    InvalidateRect(GetDlgItem(hwnd, ID_BTN_PRESET_A), NULL, FALSE);
    InvalidateRect(GetDlgItem(hwnd, ID_BTN_PRESET_B), NULL, FALSE);
    InvalidateRect(GetDlgItem(hwnd, ID_BTN_PRESET_C), NULL, FALSE);
    HWND hSave = GetDlgItem(hwnd, ID_BTN_CAPTURE_SAVE);
    if (hSave) EnableWindow(hSave, g_lastCapturedRectValid ? TRUE : FALSE);
    HWND hReuse = GetDlgItem(hwnd, ID_BTN_CAPTURE_REUSE);
    if (hReuse) EnableWindow(hReuse, g_capturePresetValid[g_capturePresetSlot] ? TRUE : FALSE);
}

static bool CaptureUsingPresetSlot(HWND owner) {
    if (g_capturePresetSlot < 0 || g_capturePresetSlot > 2) return false;
    if (!g_capturePresetValid[g_capturePresetSlot]) {
        if (owner) ShowDarkMsg(owner, L"選択中スロットに保存済み範囲がありません。", L"Info", MB_OK);
        return false;
    }
    RECT rc = {};
    if (!ClampRectToVirtualScreen(g_capturePresetRects[g_capturePresetSlot], rc)) {
        if (owner) ShowDarkMsg(owner, L"保存範囲が画面外のため再利用できません。", L"Info", MB_OK);
        return false;
    }
    g_tempImgPath = GetTempPreviewPath(L".png");
    CaptureRect(rc, g_tempImgPath);
    g_lastCapturedRect = rc;
    g_lastCapturedRectValid = true;
    g_isImageRemoved = false;
    UpdateAddPreviewImage();
    return true;
}

static bool TriggerGifExportShortcut() {
    HWND hostWnd = ResolveHostWindow(g_hDlg ? g_hDlg : g_hwnd);
    {
        std::ostringstream os;
        os << "TriggerGifExportShortcut: g_hwnd=" << (void*)g_hwnd
           << " visible=" << (g_hwnd && IsWindowVisible(g_hwnd) ? 1 : 0)
           << " hostWnd=" << (void*)hostWnd;
        AppendUiDebugLine(os.str());
    }
    if (hostWnd && IsWindow(hostWnd)) {
        if (IsIconic(hostWnd)) {
            ShowWindow(hostWnd, SW_RESTORE);
        } else {
            ShowWindow(hostWnd, SW_SHOW);
        }
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
    HWND hostWnd = ResolveHostWindow(requester);
    {
        std::ostringstream os;
        os << "ActivateHostForSelectionRead: requester=" << (void*)requester
           << " g_hwnd=" << (void*)g_hwnd
           << " visible=" << (g_hwnd && IsWindowVisible(g_hwnd) ? 1 : 0)
           << " hostWnd=" << (void*)hostWnd;
        AppendUiDebugLine(os.str());
    }
    if (hostWnd && IsWindow(hostWnd)) {
        if (IsIconic(hostWnd)) {
            ShowWindow(hostWnd, SW_RESTORE);
        } else {
            ShowWindow(hostWnd, SW_SHOW);
        }
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
            L"実際のGIFはプラグイン側でリネーム後 MyAsset フォルダへ自動保存されます。\n"
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
        HWND hCapTitle = CreateWindowW(L"STATIC", L"キャプチャ", WS_VISIBLE|WS_CHILD, 75, TITLE_H+78, 120, 18, hwnd, NULL, g_hInst, NULL);
        SendMessageW(hCapTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        CreateWindowW(L"BUTTON", L"キャプチャ", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, TITLE_H+96, 100, 24, hwnd, (HMENU)ID_BTN_SNIP, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"A", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 190, TITLE_H+96, 30, 24, hwnd, (HMENU)ID_BTN_PRESET_A, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"B", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 225, TITLE_H+96, 30, 24, hwnd, (HMENU)ID_BTN_PRESET_B, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"C", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 260, TITLE_H+96, 30, 24, hwnd, (HMENU)ID_BTN_PRESET_C, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"範囲を保存", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, TITLE_H+124, 100, 24, hwnd, (HMENU)ID_BTN_CAPTURE_SAVE, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"範囲を再利用", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 190, TITLE_H+124, 100, 24, hwnd, (HMENU)ID_BTN_CAPTURE_REUSE, g_hInst, NULL);
        HWND hGifTitle = CreateWindowW(L"STATIC", L"GIF", WS_VISIBLE|WS_CHILD, 75, TITLE_H+152, 120, 18, hwnd, NULL, g_hInst, NULL);
        SendMessageW(hGifTitle, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        CreateWindowW(L"BUTTON", L"GIF生成", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, TITLE_H+170, 100, 24, hwnd, (HMENU)ID_BTN_GIF_EXPORT, g_hInst, NULL);
        if (g_enableTextStyle) {
            HWND hTs = CreateWindowW(L"BUTTON", L"テキストスタイルとして扱う", WS_VISIBLE|WS_CHILD|BS_AUTOCHECKBOX, 75, TITLE_H+358, 220, 24, hwnd, (HMENU)IDC_CHK_TEXT_STYLE, g_hInst, NULL);
            SendMessageW(hTs, WM_SETFONT, (WPARAM)g_hFontUI, 0);
            SetWindowTheme(hTs, L"", L"");
            SendMessageW(hTs, BM_SETCHECK, g_addAsTextStyle ? BST_CHECKED : BST_UNCHECKED, 0);
        }
        CreateWindowW(L"BUTTON", L"保存", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 75, 428, 100, 30, hwnd, (HMENU)ID_BTN_SAVE, g_hInst, NULL);
        CreateWindowW(L"BUTTON", L"キャンセル", WS_VISIBLE|WS_CHILD|BS_OWNERDRAW, 190, 428, 100, 30, hwnd, (HMENU)ID_BTN_CANCEL, g_hInst, NULL);
        if (!g_editOrgPath.empty()) {
            SetDlgItemTextW(hwnd, IDC_EDIT_NAME, g_editName.c_str());
            SetDlgItemTextW(hwnd, IDC_COMBO_CAT, g_editCat.c_str());
        } else if (g_addDialogFromMyAssetAdd) {
            if (!g_currentCategory.empty() && g_currentCategory != L"ALL") {
                LRESULT catIdx = SendMessageW(hc, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)g_currentCategory.c_str());
                if (catIdx != CB_ERR) {
                    SendMessageW(hc, CB_SETCURSEL, (WPARAM)catIdx, 0);
                }
            }
            SetFocus(GetDlgItem(hwnd, IDC_EDIT_NAME));
        }
        UpdateAddCapturePresetUi(hwnd);
        UpdateAddPreviewImage(); return 0;
    }
    case WM_TIMER:
        if (wp == ID_TIMER_ADD_PREVIEW && g_pAddPreviewImage && g_addFrameCount > 1) {
            g_addCurrentFrame = (g_addCurrentFrame + 1) % g_addFrameCount;
            RECT rp = GetAddPreviewRect();
            InvalidateRect(hwnd, &rp, FALSE);
        } else if (wp == ID_TIMER_HIDE_MAIN_AFTER_EXPORT) {
            if (g_hideMainAfterExportTicks > 0) {
                g_hideMainAfterExportTicks--;
                if (!g_mainWasVisibleBeforeAddDialog && g_hwnd && IsWindow(g_hwnd) && IsWindowVisible(g_hwnd)) {
                    AppendUiDebugLine("HideMainTimer: force hide g_hwnd");
                    ShowWindow(g_hwnd, SW_HIDE);
                }
            }
            if (g_hideMainAfterExportTicks <= 0) {
                g_suppressMainShow = false;
                AppendUiDebugLine("HideMainTimer: stop watchdog and release suppress");
                KillTimer(hwnd, ID_TIMER_HIDE_MAIN_AFTER_EXPORT);
            }
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
                            L"ファイル：MyAsset Preview GIF を Ctrl+Shift+M に設定してください");
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
                    L"ファイル：MyAsset Preview GIF を Ctrl+Shift+M に設定してください");
            }
        }
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); FillRect(hdc, &rc, g_hBrBg); DrawTitleBar(hdc, rc.right, L"アセット編集"); DrawWindowBorder(hwnd);
        Graphics graphics(hdc); RECT rp = GetAddPreviewRect(); HBRUSH brP = CreateSolidBrush(RGB(20, 20, 20)); FillRect(hdc, &rp, brP); DeleteObject(brP);
        
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
        RECT rp = GetAddPreviewRect();
        if (g_pAddPreviewImage) { RECT rx = { rp.right - 24, rp.top, rp.right, rp.top + 24 }; if (PtInRect(&rx, pt)) { g_isImageRemoved = true; g_tempImgPath = L""; UpdateAddPreviewImage(); return 0; } }
        if (PtInRect(&rp, pt)) { wchar_t sz[MAX_PATH] = {0}; OPENFILENAMEW of = {0}; of.lStructSize = sizeof(of); of.hwndOwner = hwnd; of.lpstrFile = sz; of.nMaxFile = sizeof(sz); of.lpstrFilter = L"Image Files\0*.png;*.gif;*.jpg;*.jpeg;*.bmp\0All Files\0*.*\0"; of.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST; if (GetOpenFileNameW(&of)) { g_tempImgPath = sz; g_isImageRemoved = false; UpdateAddPreviewImage(); } }
        return 0;
    }
    case WM_DRAWITEM: DrawDarkButton((LPDRAWITEMSTRUCT)lp); return TRUE;
    case WM_CTLCOLOREDIT: case WM_CTLCOLORLISTBOX: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkColor(hdc, COL_INPUT_BG); return (LRESULT)g_hBrInputBg; }
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORSTATIC: { HDC hdc = (HDC)wp; SetTextColor(hdc, COL_TEXT); SetBkMode(hdc, TRANSPARENT); return (LRESULT)g_hBrBg; }
    case WM_COMMAND: {
        if (HIWORD(wp) == BN_CLICKED) {
            int id = LOWORD(wp);
            if (id == ID_BTN_PRESET_A || id == ID_BTN_PRESET_B || id == ID_BTN_PRESET_C) {
                g_capturePresetSlot = (id == ID_BTN_PRESET_A) ? 0 : (id == ID_BTN_PRESET_B ? 1 : 2);
                SaveConfig();
                UpdateAddCapturePresetUi(hwnd);
                return 0;
            }
            if (id == ID_BTN_CAPTURE_SAVE) {
                if (!g_lastCapturedRectValid) return 0;
                g_capturePresetRects[g_capturePresetSlot] = g_lastCapturedRect;
                g_capturePresetValid[g_capturePresetSlot] = true;
                SaveConfig();
                UpdateAddCapturePresetUi(hwnd);
                return 0;
            }
            if (id == ID_BTN_CAPTURE_REUSE) {
                if (CaptureUsingPresetSlot(hwnd)) {
                    UpdateAddCapturePresetUi(hwnd);
                }
                return 0;
            }
        }
        if (LOWORD(wp) == IDC_CHK_TEXT_STYLE && HIWORD(wp) == BN_CLICKED) {
            if (!g_enableTextStyle) { g_addAsTextStyle = false; return 0; }
            g_addAsTextStyle = (SendMessageW(GetDlgItem(hwnd, IDC_CHK_TEXT_STYLE), BM_GETCHECK, 0, 0) == BST_CHECKED);
            return 0;
        }
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
                std::wstring oldPathLower = ToLower(g_editOrgPath);
                std::wstring newPathLower = ToLower(sp);
                if (oldPathLower != newPathLower) {
                    MoveFileW(g_editOrgPath.c_str(), sp.c_str());
                }
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

            if (!g_editOrgPath.empty()) {
                g_textStylePaths.erase(g_editOrgPath);
            }
            g_textStylePaths.erase(sp);
            if (g_enableTextStyle && g_addAsTextStyle) {
                g_textStylePaths.insert(sp);
            }
            SaveTextStyleFlags();

            RefreshAssets(false); 
            if (g_mainWasVisibleBeforeAddDialog && g_hwnd && IsWindow(g_hwnd)) {
                ShowWindow(g_hwnd, SW_SHOW);
                SetForegroundWindow(g_hwnd);
                InvalidateRect(g_hwnd, NULL, FALSE);
            }
            DestroyWindow(hwnd);
        } else if (LOWORD(wp) == ID_BTN_CANCEL) DestroyWindow(hwnd);
        else if (LOWORD(wp) == ID_BTN_SNIP) StartSnipping();
        else if (LOWORD(wp) == ID_BTN_GIF_EXPORT) {
            if (!ShowGifExportGuideIfNeeded(hwnd)) return 0;
            bool mainWasVisible = (g_hwnd && IsWindow(g_hwnd) && IsWindowVisible(g_hwnd));
            auto restoreMainHidden = [&]() {
                if (!mainWasVisible && g_hwnd && IsWindow(g_hwnd) && IsWindowVisible(g_hwnd)) {
                    ShowWindow(g_hwnd, SW_HIDE);
                }
            };
            bool fromCapturedSelection = false;
            bool fromTimeline = false;
            bool fromAlias = false;

            if (g_addDialogRangeValid) {
                g_cachedOutputRangeStart = g_addDialogRangeStart;
                g_cachedOutputRangeEnd = g_addDialogRangeEnd;
                g_cachedOutputRangeValid = true;
                fromCapturedSelection = true;
            } else {
                ActivateHostForSelectionRead(hwnd);
                fromTimeline = CacheRangeFromTimelineSelection();
            }

            if (!fromCapturedSelection && !fromTimeline) {
                fromAlias = CacheRangeFromCurrentAliasData();
            }
            if (!fromCapturedSelection && !fromTimeline && !fromAlias) {
                restoreMainHidden();
                ShowDarkMsg(hwnd, L"書き出し範囲を取得できませんでした。\n先に「MyAsset：追加」で対象オブジェクトを選択してください。", L"Info", MB_OK);
                return 0;
            }
            if (g_enableGifRangeDebug) {
                std::wstring source = L"alias-fallback";
                if (fromCapturedSelection) source = L"captured-selection";
                else if (fromTimeline) source = L"timeline-selection";
                std::wstring msg = L"range-source: " + source +
                    L"\nrange: " + std::to_wstring(g_cachedOutputRangeStart) + L" ～ " + std::to_wstring(g_cachedOutputRangeEnd) +
                    L"\nalias-bytes: " + std::to_wstring(g_tempAliasData.size());
                MessageBoxW(hwnd, msg.c_str(), L"MyAsset GIF Range Debug", MB_OK);
            }
            ActivateHostForSelectionRead(hwnd);
            if (!ApplyCachedRangeToTimeline()) {
                restoreMainHidden();
                ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
                ShowDarkMsg(hwnd, L"タイムラインの範囲指定に失敗しました。", L"Error", MB_OK);
                return 0;
            }
            g_suppressMainShow = true;
            DumpProcessWindowsForUiDebug("before-trigger");
            if (!TriggerGifExportShortcut()) {
                g_suppressMainShow = false;
                restoreMainHidden();
                ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
                ShowDarkMsg(hwnd, L"ショートカット送信に失敗しました。", L"Error", MB_OK);
                return 0;
            }
            DumpProcessWindowsForUiDebug("after-trigger");
            if (!mainWasVisible) {
                g_hideMainAfterExportTicks = 30;
                SetTimer(hwnd, ID_TIMER_HIDE_MAIN_AFTER_EXPORT, 100, NULL);
                AppendUiDebugLine("GIF export: start hide-main watchdog timer");
            } else {
                g_suppressMainShow = false;
            }
            restoreMainHidden();
        }
        return 0;
    }

    case WM_DESTROY: 
        SaveWindowPos(hwnd, true); 
        DeleteFileW(GetForcedPreviewOutputPath().c_str());
        KillTimer(hwnd, ID_TIMER_GIF_HINT_POLL);
        KillTimer(hwnd, ID_TIMER_HIDE_MAIN_AFTER_EXPORT);
        g_hideMainAfterExportTicks = 0;
        if (g_addGifHintTimerStarted) {
            KillTimer(hwnd, ID_TIMER_GIF_HINT);
            g_addGifHintTimerStarted = false;
        }
        g_addDlgMouseTracking = false;
        g_addGifHoverStartTick = 0;
        HideMyTooltip();
        if (g_pAddPreviewImage) { delete g_pAddPreviewImage; g_pAddPreviewImage = nullptr; }
        if (g_pAddPreviewStream) { g_pAddPreviewStream->Release(); g_pAddPreviewStream = nullptr; }
        g_addDialogFromMyAssetAdd = false;
        g_hDlg = nullptr; return 0;
    } return DefWindowProc(hwnd, msg, wp, lp);
}

void OpenAddDialog(const std::string& data, bool isEdit) { 
    if (g_hDlg) { ShowWindow(g_hDlg, SW_SHOW); SetForegroundWindow(g_hDlg); return; } 
    g_mainWasVisibleBeforeAddDialog = (g_hwnd && IsWindow(g_hwnd) && IsWindowVisible(g_hwnd));
    g_lastCapturedRectValid = false;
    g_tempAliasData = data; if (!isEdit) g_editOrgPath = L""; g_isImageRemoved = false; g_tempImgPath = L""; 
    g_addAsTextStyle = false;
    if (isEdit) {
        g_addAsTextStyle = g_enableTextStyle && (g_textStylePaths.count(g_editOrgPath) > 0);
        g_addDialogRangeValid = false;
        g_addDialogRangeStart = 0;
        g_addDialogRangeEnd = 0;
    }
    LoadWindowConfig(); int sw = GetSystemMetrics(SM_CXSCREEN), sh = GetSystemMetrics(SM_CYSCREEN);
    int x = (g_dlgX == -1) ? (sw - 360) / 2 : g_dlgX; int y = (g_dlgY == -1) ? (sh - 490) / 2 : g_dlgY;
    WNDCLASSW wc = {0}; wc.lpfnWndProc = AddDlgProc; wc.hInstance = g_hInst; wc.lpszClassName = L"MyAsset_Add"; wc.hbrBackground = g_hBrBg; RegisterClassW(&wc); 
    g_hDlg = CreateWindowExW(WS_EX_TOPMOST|WS_EX_TOOLWINDOW, L"MyAsset_Add", L"アセット", WS_POPUP|WS_VISIBLE|WS_THICKFRAME, x, y, 360, 490, g_hwnd, nullptr, g_hInst, nullptr); 
}

// ============================================================
// 8. メイン描画 & プロシージャ (WndProc)
// ============================================================

static void DrawContent(HDC hdc, int w, int h) {
    RECT rcA = {0,0,w,h}; FillRect(hdc, &rcA, g_hBrBg); 
    int listTop = GetListTopY();
    auto drawTopChrome = [&]() {
        DrawTitleBar(hdc, w, L"My Asset Manager", false);
        if (g_enableTextStyle) {
            RECT rcTabBg = {0, TITLE_H, w, listTop};
            HBRUSH brTabBg = CreateSolidBrush(COL_BG);
            FillRect(hdc, &rcTabBg, brTabBg);
            DeleteObject(brTabBg);
            RECT rcTabBorder = {0, listTop - 1, w, listTop};
            HBRUSH brTabBorder = CreateSolidBrush(COL_BORDER);
            FillRect(hdc, &rcTabBorder, brTabBorder);
            DeleteObject(brTabBorder);

            RECT rcTabAll = {10, TITLE_H + 4, 130, listTop - 4};
            RECT rcTabStyle = {136, TITLE_H + 4, 276, listTop - 4};
            HBRUSH brTabAll = CreateSolidBrush(g_assetViewTab == 0 ? COL_BTN_ACT : COL_BTN_BG);
            HBRUSH brTabStyle = CreateSolidBrush(g_assetViewTab == 1 ? COL_BTN_ACT : COL_BTN_BG);
            FillRect(hdc, &rcTabAll, brTabAll);
            FillRect(hdc, &rcTabStyle, brTabStyle);
            DeleteObject(brTabAll);
            DeleteObject(brTabStyle);
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, COL_TEXT);
            SelectObject(hdc, g_hFontUI);
            DrawTextW(hdc, L"一覧", -1, &rcTabAll, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
            DrawTextW(hdc, L"テキストスタイル", -1, &rcTabStyle, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
        }
    };
    drawTopChrome();

    int listBottom = h - GetBottomReservedHeight();
    // 一覧領域だけをクリップして、ヘッダ/フッタへのはみ出し描画を防ぐ。
    int yO = listTop - g_scrollY; HRGN hr = CreateRectRgn(0, listTop, w, listBottom); SelectClipRgn(hdc, hr);
    int itemH = GetListItemHeight();
    int thumbW = GetThumbWScaled();
    int thumbH = GetThumbHScaled();
    Graphics g(hdc); int cols = (std::max)(1, w / GetListItemMinWidth()); int cw = w / cols;
    
    SetBkMode(hdc, TRANSPARENT);

    for (int i = 0; i < (int)g_displayAssets.size(); i++) {
        int x = (i % cols) * cw, y = (i / cols) * itemH + yO;
        if (y > listBottom || y + itemH < listTop) continue;
        RECT ri = {x + 2, y + 2, x + cw - 2, y + itemH - 2}; if (!RectVisible(hdc, &ri)) continue;
        
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
            float thumbRatio = (float)thumbW / (float)thumbH;
            float imgRatio = (float)iw / (float)ih;

            float srcX = 0, srcY = 0, srcW = (float)iw, srcH = (float)ih;

            if (imgRatio > thumbRatio) {
                srcW = ih * thumbRatio;
                srcX = (iw - srcW) / 2.0f;
            } else {
                srcH = iw / thumbRatio;
                srcY = (ih - srcH) / 2.0f;
            }

            int imgX = g_previewThumbOnly ? (x + (cw - thumbW) / 2) : (x + 7);
            int imgY = y + (itemH - thumbH) / 2;
            g.DrawImage(p->pImage, 
                Rect(imgX, imgY, thumbW, thumbH),
                (INT)srcX, (INT)srcY, (INT)srcW, (INT)srcH,
                UnitPixel);
        }

        if (!g_previewThumbOnly && p->isFavorite) { SetTextColor(hdc, RGB(255, 215, 0)); SelectObject(hdc, g_hFontUI); RECT rs = {ri.right - 25, ri.top + 5, ri.right - 5, ri.top + 25}; DrawTextW(hdc, L"★", -1, &rs, DT_RIGHT); }
        
        if (!g_previewThumbOnly) {
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

            SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontList); RECT rn = {x + thumbW + 15, y + 15, ri.right - 30, y + 45}; DrawTextW(hdc, p->name.c_str(), -1, &rn, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            SetTextColor(hdc, COL_SUBTEXT); SelectObject(hdc, g_hFontListSub); RECT rc = {x + thumbW + 15, y + 50, ri.right - 10, ri.bottom - 5}; DrawTextW(hdc, p->category.c_str(), -1, &rc, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);
        }

        
    }
    // カスタム並び替え中は、挿入位置を青バーで可視化する。
    if (g_isReorderDrag && g_reorderInsertPos >= 0 && g_reorderInsertPos <= (int)g_displayAssets.size()) {
        int ins = g_reorderInsertPos;
        int barX = 2;
        int barY = listTop + 4 - g_scrollY;
        int barH = itemH - 8;
        if (ins < (int)g_displayAssets.size()) {
            int row = ins / cols;
            int col = ins % cols;
            barX = col * cw + 2;
            barY = row * itemH + listTop - g_scrollY + 4;
        } else if (!g_displayAssets.empty()) {
            int last = (int)g_displayAssets.size() - 1;
            int row = last / cols;
            int col = last % cols;
            if (col == cols - 1) {
                barX = 2;
                barY = (row + 1) * itemH + listTop - g_scrollY + 4;
            } else {
                barX = (col + 1) * cw + 2;
                barY = row * itemH + listTop - g_scrollY + 4;
            }
        }
        RECT rb = { barX - 2, barY, barX + 2, barY + barH };
        HBRUSH brIns = CreateSolidBrush(RGB(80, 170, 255));
        FillRect(hdc, &rb, brIns);
        DeleteObject(brIns);
    }

    SelectClipRgn(hdc, NULL); DeleteObject(hr);
    drawTopChrome();
    if (CanStartCustomReorder() && (GetKeyState(VK_CONTROL) & 0x8000)) {
        RECT rcMode = { w - 220, 4, w - 8, TITLE_H - 4 };
        HBRUSH brMode = CreateSolidBrush(RGB(50, 90, 180));
        FillRect(hdc, &rcMode, brMode);
        DeleteObject(brMode);
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, RGB(235, 245, 255));
        SelectObject(hdc, g_hFontUI);
        DrawTextW(hdc, L"並び替えモード", -1, &rcMode, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
    if (g_enableTextStyle && g_assetViewTab == 1) {
        RECT rcStyle = {0, h - FOOTER_H - STYLE_ACTION_H, w, h - FOOTER_H};
        HBRUSH brStyle = CreateSolidBrush(COL_BG);
        FillRect(hdc, &rcStyle, brStyle);
        DeleteObject(brStyle);
        RECT rcStyleBorder = {0, h - FOOTER_H - STYLE_ACTION_H, w, h - FOOTER_H - STYLE_ACTION_H + 1};
        HBRUSH brStyleBorder = CreateSolidBrush(COL_BORDER);
        FillRect(hdc, &rcStyleBorder, brStyleBorder);
        DeleteObject(brStyleBorder);

        RECT rcCommit = {10, h - FOOTER_H - STYLE_ACTION_H + 6, 100, h - FOOTER_H - 6};
        RECT rcCancel = {106, h - FOOTER_H - STYLE_ACTION_H + 6, 196, h - FOOTER_H - 6};
        RECT rcDetail = {202, h - FOOTER_H - STYLE_ACTION_H + 6, 292, h - FOOTER_H - 6};
        COLORREF commitCol = g_textStylePending ? COL_BTN_ACT : COL_BTN_BG;
        COLORREF cancelCol = g_textStylePending ? RGB(120, 70, 70) : COL_BTN_BG;
        COLORREF detailCol = COL_BTN_BG;
        HBRUSH brCommit = CreateSolidBrush(commitCol);
        HBRUSH brCancel = CreateSolidBrush(cancelCol);
        HBRUSH brDetail = CreateSolidBrush(detailCol);
        FillRect(hdc, &rcCommit, brCommit);
        FillRect(hdc, &rcCancel, brCancel);
        FillRect(hdc, &rcDetail, brDetail);
        DeleteObject(brCommit);
        DeleteObject(brCancel);
        DeleteObject(brDetail);
        SetTextColor(hdc, g_textStylePending ? COL_TEXT : COL_SUBTEXT);
        DrawTextW(hdc, L"確定", -1, &rcCommit, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
        DrawTextW(hdc, L"取消", -1, &rcCancel, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
        SetTextColor(hdc, COL_TEXT);
        DrawTextW(hdc, L"詳細設定", -1, &rcDetail, DT_CENTER|DT_VCENTER|DT_SINGLELINE);

        RECT rcHint = {298, h - FOOTER_H - STYLE_ACTION_H + 8, w - 10, h - FOOTER_H - 6};
        SetTextColor(hdc, COL_SUBTEXT);
        DrawTextW(hdc, L"クリックで適用", -1, &rcHint, DT_LEFT|DT_VCENTER|DT_SINGLELINE|DT_END_ELLIPSIS);
    }

    RECT rcF = {0, h - FOOTER_H, w, h}; HBRUSH brF = CreateSolidBrush(COL_FOOTER); FillRect(hdc, &rcF, brF); DeleteObject(brF);
    RECT rcRef = {w - 70, h - FOOTER_H + 8, w - 10, h - FOOTER_H + 32}; HBRUSH brRef = CreateSolidBrush(COL_BTN_BG); FillRect(hdc, &rcRef, brRef); DeleteObject(brRef);
    SetBkMode(hdc, TRANSPARENT); SetTextColor(hdc, COL_TEXT); SelectObject(hdc, g_hFontUI); DrawTextW(hdc, L"更新", -1, &rcRef, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
    RECT rcFav = {w - 140, h - FOOTER_H + 8, w - 80, h - FOOTER_H + 32}; HBRUSH brFav = CreateSolidBrush(COL_BTN_BG); FillRect(hdc, &rcFav, brFav); DeleteObject(brFav);
    const wchar_t* sortLabel = L"名前順";
    if (g_sortMode == 1) sortLabel = L"★優先";
    else if (g_sortMode == 2) sortLabel = L"カテゴリ";
    else if (g_sortMode == 3) sortLabel = L"カスタム";
    DrawTextW(hdc, sortLabel, -1, &rcFav, DT_CENTER|DT_VCENTER|DT_SINGLELINE);

    DrawWindowBorder(g_hwnd);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_SHOWWINDOW: {
        std::ostringstream os;
        os << "WM_SHOWWINDOW: hwnd=" << (void*)hwnd
           << " show=" << (wp ? 1 : 0)
           << " status=" << (long long)lp
           << " visibleNow=" << (IsWindowVisible(hwnd) ? 1 : 0)
           << " suppress=" << (g_suppressMainShow ? 1 : 0);
        AppendUiDebugLine(os.str());
        return DefWindowProc(hwnd, msg, wp, lp);
    }
    case WM_APP_DEBUG_LAYER_COMPARE:
        RunLayerApiCompareDebug();
        return 0;
    case WM_APP_BUILD_ADD_ASSET:
        ProcessPendingAddAssetRequestByApi();
        return 0;
    case WM_CREATE: {
        OleInitialize(NULL); GdiplusStartupInput si; GdiplusStartup(&g_gdiplusToken, &si, NULL);
        g_hFontUI = CreateFontW(15,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontList = CreateFontW(17,0,0,0,FW_BOLD,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontListSub = CreateFontW(13,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Yu Gothic UI");
        g_hFontType = CreateFontW(16,0,0,0,FW_BOLD,0,0,0,DEFAULT_CHARSET,0,0,0,0,L"Arial");

        RecreateUiBrushes();
        RegisterCustomDialogs();
        g_hCombo = CreateWindowW(L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST, 10, 0, 140, 300, hwnd, (HMENU)ID_COMBO_CATEGORY, g_hInst, NULL); SendMessageW(g_hCombo, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        g_hSearch = CreateWindowW(L"EDIT", L"", WS_CHILD|WS_VISIBLE|WS_BORDER|ES_AUTOHSCROLL, 160, 0, 100, 24, hwnd, (HMENU)ID_EDIT_SEARCH, g_hInst, NULL); SendMessageW(g_hSearch, WM_SETFONT, (WPARAM)g_hFontUI, 0);
        
        CheckAndOptimizeLegacyAssets(hwnd);
        
        RefreshAssets(true); return 0;
    }

    case WM_NCHITTEST: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        ScreenToClient(hwnd, &pt);
        if (pt.y < TITLE_H) return HTCLIENT;
        return HTCLIENT;
    }

    case WM_NCCALCSIZE: if(wp) return 0; return DefWindowProc(hwnd, msg, wp, lp);
    case WM_NCACTIVATE: return DefWindowProc(hwnd, msg, wp, lp);
    case WM_MOUSEACTIVATE: return MA_ACTIVATE;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;

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
        if (wp == ID_TIMER_LAYER_API_COMPARE) {
            KillTimer(hwnd, ID_TIMER_LAYER_API_COMPARE);
            if (g_pendingLayerApiCompare) {
                g_pendingLayerApiCompare = false;
                RunLayerApiCompareDebug();
            }
            return 0;
        }
        
        if (wp == ID_TIMER_HOVER) {
            bool pausePreviewForReorder = CanStartCustomReorder() && (((GetKeyState(VK_CONTROL) & 0x8000) != 0) || g_isReorderDrag);
            if (pausePreviewForReorder) {
                KillTimer(hwnd, ID_TIMER_HOVER);
                g_previewAllLastTick = 0;
                g_previewHoverLastTick = 0;
                return 0;
            }
            RECT rc = {};
            GetClientRect(hwnd, &rc);
            int listTop = GetListTopY();
            POINT pt = {};
            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);
            int listBottom = rc.bottom - GetBottomReservedHeight();
            bool mouseInList = (pt.x >= 0 && pt.x < rc.right && pt.y > listTop && pt.y < listBottom);

            if (g_previewPlaybackMode == 1) {
                if (mouseInList) {
                    ULONGLONG now = GetTickCount64();
                    UINT tickMs = GetPreviewAllTickMs();
                    double elapsedMs = (g_previewAllLastTick == 0) ? (double)tickMs : (double)(now - g_previewAllLastTick);
                    if (elapsedMs < 0.0) elapsedMs = (double)tickMs;
                    g_previewAllLastTick = now;

                    double speedScale = (double)(std::max)(1, g_gifSpeedPercent) / 100.0;
                    double virtualElapsedMs = elapsedMs * speedScale;

                    int viewH = listBottom - listTop;
                    int itemH = GetListItemHeight();
                    int cols = (std::max)(1, (int)rc.right / GetListItemMinWidth());
                    int firstRow = (viewH > 0) ? (g_scrollY / itemH) : 0;
                    int lastRow = (viewH > 0) ? ((g_scrollY + viewH - 1) / itemH) : -1;
                    int firstIdx = firstRow * cols;
                    int lastIdx = ((lastRow + 1) * cols) - 1;
                    if (firstIdx < 0) firstIdx = 0;
                    if (lastIdx >= (int)g_displayAssets.size()) lastIdx = (int)g_displayAssets.size() - 1;

                    bool anyAnimated = false;
                    bool anyFrameChanged = false;
                    for (int i = firstIdx; i <= lastIdx; ++i) {
                        Asset* p = g_displayAssets[i];
                        if (!p || p->frameCount <= 1) continue;
                        anyAnimated = true;
                        p->previewCarryMs += virtualElapsedMs;

                        UINT cycleMs = p->previewCycleMs;
                        if (cycleMs == 0) {
                            unsigned long long cycleSum = 0;
                            if (p->frameDelays) {
                                for (UINT fi = 0; fi < p->frameCount; ++fi) cycleSum += (unsigned long long)(std::max)(1u, p->frameDelays[fi]);
                            } else {
                                cycleSum = (unsigned long long)p->frameCount * 100ULL;
                            }
                            if (cycleSum == 0) cycleSum = (unsigned long long)p->frameCount * 100ULL;
                            p->previewCycleMs = (cycleSum > 0xFFFFFFFFULL) ? 0xFFFFFFFFu : (UINT)cycleSum;
                            cycleMs = p->previewCycleMs;
                        }

                        double cycle = (double)(std::max)(1u, cycleMs);
                        if (p->previewCarryMs >= cycle * 8.0) {
                            p->previewCarryMs = std::fmod(p->previewCarryMs, cycle);
                        }
                        double t = std::fmod(p->previewCarryMs, cycle);
                        if (t < 0.0) t += cycle;

                        UINT acc = 0;
                        int targetFrame = p->currentFrame;
                        for (UINT fi = 0; fi < p->frameCount; ++fi) {
                            UINT delayMs = p->frameDelays ? (std::max)(1u, p->frameDelays[fi]) : 100u;
                            UINT nextAcc = acc + delayMs;
                            if (t < (double)nextAcc) {
                                targetFrame = (int)fi;
                                break;
                            }
                            acc = nextAcc;
                        }
                        if (targetFrame != p->currentFrame) {
                            p->currentFrame = targetFrame;
                            anyFrameChanged = true;
                        }
                    }
                    if (anyAnimated) {
                        if (anyFrameChanged) {
                            RECT rcList = { 0, listTop, rc.right, listBottom };
                            InvalidateRect(hwnd, &rcList, FALSE);
                        }
                        SetTimer(hwnd, ID_TIMER_HOVER, tickMs, NULL);
                    } else {
                        KillTimer(hwnd, ID_TIMER_HOVER);
                        g_previewAllLastTick = 0;
                    }
                } else {
                    KillTimer(hwnd, ID_TIMER_HOVER);
                    g_previewAllLastTick = 0;
                }
            } else {
                if (mouseInList && g_hoverIndex != -1 && g_hoverIndex < (int)g_displayAssets.size()) {
                    Asset* p = g_displayAssets[g_hoverIndex];
                    if (p->frameCount > 1) {
                        ULONGLONG now = GetTickCount64();
                        UINT tickMs = GetPreviewAllTickMs();
                        double elapsedMs = (g_previewHoverLastTick == 0) ? (double)tickMs : (double)(now - g_previewHoverLastTick);
                        if (elapsedMs < 0.0) elapsedMs = (double)tickMs;
                        g_previewHoverLastTick = now;

                        double speedScale = (double)(std::max)(1, g_gifSpeedPercent) / 100.0;
                        p->previewCarryMs += elapsedMs * speedScale;

                        UINT cycleMs = p->previewCycleMs;
                        if (cycleMs == 0) {
                            unsigned long long cycleSum = 0;
                            if (p->frameDelays) {
                                for (UINT fi = 0; fi < p->frameCount; ++fi) cycleSum += (unsigned long long)(std::max)(1u, p->frameDelays[fi]);
                            } else {
                                cycleSum = (unsigned long long)p->frameCount * 100ULL;
                            }
                            if (cycleSum == 0) cycleSum = (unsigned long long)p->frameCount * 100ULL;
                            p->previewCycleMs = (cycleSum > 0xFFFFFFFFULL) ? 0xFFFFFFFFu : (UINT)cycleSum;
                            cycleMs = p->previewCycleMs;
                        }

                        double cycle = (double)(std::max)(1u, cycleMs);
                        if (p->previewCarryMs >= cycle * 8.0) {
                            p->previewCarryMs = std::fmod(p->previewCarryMs, cycle);
                        }
                        double t = std::fmod(p->previewCarryMs, cycle);
                        if (t < 0.0) t += cycle;

                        UINT acc = 0;
                        int targetFrame = p->currentFrame;
                        for (UINT fi = 0; fi < p->frameCount; ++fi) {
                            UINT delayMs = p->frameDelays ? (std::max)(1u, p->frameDelays[fi]) : 100u;
                            UINT nextAcc = acc + delayMs;
                            if (t < (double)nextAcc) {
                                targetFrame = (int)fi;
                                break;
                            }
                            acc = nextAcc;
                        }
                        if (targetFrame != p->currentFrame) {
                            p->currentFrame = targetFrame;
                            InvalidateRect(hwnd, NULL, FALSE);
                        }
                        SetTimer(hwnd, ID_TIMER_HOVER, tickMs, NULL);
                    } else {
                        KillTimer(hwnd, ID_TIMER_HOVER);
                        g_previewHoverLastTick = 0;
                    }
                } else {
                    KillTimer(hwnd, ID_TIMER_HOVER);
                    g_previewHoverLastTick = 0;
                }
            }
        }
        if (wp == ID_TIMER_TOOLTIP) {
            KillTimer(hwnd, ID_TIMER_TOOLTIP);
            if (g_tooltipTargetIndex != -1 && g_tooltipTargetIndex < (int)g_displayAssets.size()) {
                RECT rc; GetClientRect(hwnd, &rc);
                int listTop = GetListTopY();
                POINT ptClient = {};
                GetCursorPos(&ptClient);
                ScreenToClient(hwnd, &ptClient);
                int listBottom = rc.bottom - GetBottomReservedHeight();
                if (!(ptClient.x >= 0 && ptClient.x < rc.right && ptClient.y > listTop && ptClient.y < listBottom)) {
                    g_tooltipTargetIndex = -1;
                    HideMyTooltip();
                    return 0;
                }

                int itemH = GetListItemHeight();
                int cols = (std::max)(1, (int)rc.right / GetListItemMinWidth());
                int cw = (int)rc.right / cols;
                int row = (ptClient.y - listTop + g_scrollY) / itemH;
                int col = ptClient.x / cw;
                int idx = row * cols + col;
                int itemX = col * cw;
                int itemY = row * itemH + (listTop - g_scrollY);
                bool stillOnTarget = false;
                if (idx == g_tooltipTargetIndex) {
                    if (g_previewThumbOnly) {
                        int thumbW = GetThumbWScaled();
                        int thumbH = GetThumbHScaled();
                        RECT rcThumb = {
                            itemX + (cw - thumbW) / 2,
                            itemY + (itemH - thumbH) / 2,
                            itemX + (cw - thumbW) / 2 + thumbW,
                            itemY + (itemH - thumbH) / 2 + thumbH
                        };
                        stillOnTarget = PtInRect(&rcThumb, ptClient) ? true : false;
                    } else {
                        int right = itemX + cw - 2;
                        int bottom = itemY + itemH - 2;
                        RECT rcMark = { right - 30, bottom - 30, right, bottom };
                        stillOnTarget = PtInRect(&rcMark, ptClient) ? true : false;
                    }
                }
                if (!stillOnTarget) {
                    g_tooltipTargetIndex = -1;
                    HideMyTooltip();
                    return 0;
                }

                Asset* p = g_displayAssets[g_tooltipTargetIndex];
                std::wstring mainT, subT;
                if (g_previewThumbOnly) {
                    mainT = p->name;
                    std::wstring sm = p->isMulti ? L"M (複数オブジェクト)" : (p->isFixedFrame ? L"S (単体/固定)" : L"S (単体/可変)");
                    subT = L"カテゴリ: " + p->category + L"\n仕様: " + sm;
                } else if (p->isFixedFrame) {
                    if (p->isMulti) {
                        mainT = L"複数オブジェクト (Multi Object)";
                        subT = L"構造を維持するため、保存時の長さで固定されます";
                    } else {
                        mainT = L"単体オブジェクト (Single Object)";
                        subT = L"保存時のフレーム数を維持して配置します";
                    }
                } else {
                    if (p->isMulti) {
                        mainT = L"複数オブジェクト (Multi Object)";
                        subT = L"構造を維持するため、保存時の長さで固定されます";
                    } else {
                        mainT = L"単体オブジェクト (Single Object)";
                        subT = L"現在のレイヤー表示の拡大率によって配置の長さが自動で伸縮します";
                    }
                }

                POINT showPt = {};
                if (g_previewThumbOnly) {
                    POINT itemTL = { itemX, itemY };
                    POINT itemBR = { itemX + cw, itemY + itemH };
                    ClientToScreen(hwnd, &itemTL);
                    ClientToScreen(hwnd, &itemBR);
                    int screenW = GetSystemMetrics(SM_CXSCREEN);
                    int tooltipApproxW = 240;
                    if (itemBR.x + tooltipApproxW + 20 <= screenW) {
                        showPt.x = itemBR.x + 4;
                    } else {
                        LONG leftX = itemTL.x - (LONG)tooltipApproxW;
                        showPt.x = (leftX > 0) ? leftX : 0;
                    }
                    showPt.y = itemTL.y;
                } else {
                    GetCursorPos(&showPt);
                }
                ShowMyTooltip(showPt.x, showPt.y, mainT, subT);
            }
        }
        return 0;

    case WM_MOUSEWHEEL: {
        int delta = GET_WHEEL_DELTA_WPARAM(wp);
        UINT lines = 3;
        SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        int step = (lines == WHEEL_PAGESCROLL) ? (GetListItemHeight() * 3) : ((int)lines * SCROLL_SPD);
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
        g_isReorderDrag = false;
        g_reorderFromIndex = -1;
        g_reorderInsertPos = -1;
        g_hoverIndex = -1;
        KillTimer(hwnd, ID_TIMER_TOOLTIP);
        g_tooltipTargetIndex = -1;
        HideMyTooltip(); 
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;

    case WM_CAPTURECHANGED:
        g_isDragCheck = false;
        g_isReorderDrag = false;
        g_reorderFromIndex = -1;
        g_reorderInsertPos = -1;
        return 0;

    case WM_RBUTTONDOWN:
    case WM_KEYDOWN:
    case WM_KEYUP:
        {
        bool cancelByInput = (msg == WM_RBUTTONDOWN) || (msg == WM_KEYDOWN && wp == VK_ESCAPE);
        if (cancelByInput && g_isDragCheck) {
            ReleaseCapture();
            g_isDragCheck = false;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        if (cancelByInput && g_isReorderDrag) {
            ReleaseCapture();
            g_isReorderDrag = false;
            g_reorderFromIndex = -1;
            g_reorderInsertPos = -1;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        if ((msg == WM_KEYDOWN || msg == WM_KEYUP) && wp == VK_CONTROL) {
            InvalidateRect(hwnd, NULL, FALSE);
        }
        if (msg == WM_RBUTTONDOWN) goto RBUTTON_HANDLER; 
        return 0;
        }

    case WM_NCRBUTTONUP:
        if (wp == HTCAPTION) {
            ShowSettingsRefreshMenu(hwnd, GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
            return 0;
        }
        return DefWindowProc(hwnd, msg, wp, lp);

    case WM_RBUTTONUP: {
RBUTTON_HANDLER:
        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        int listTop = GetListTopY();
        if (y >= 0 && y < TITLE_H) {
            POINT pt; GetCursorPos(&pt);
            ShowSettingsRefreshMenu(hwnd, pt.x, pt.y);
        }
        else if (y > listTop && y < rc.bottom - GetBottomReservedHeight()) {
            int cols = (std::max)(1, (int)rc.right / GetListItemMinWidth()); int cw = (int)rc.right / cols;
            int idx = ((y - listTop + g_scrollY) / GetListItemHeight()) * cols + (x / cw);
            HMENU hm = CreatePopupMenu();
            if (idx >= 0 && idx < (int)g_displayAssets.size()) { 
                g_contextTargetIndex = idx; Asset* p = g_displayAssets[idx]; 
                
                // Multiの場合はフレーム固定の切り替えを無効化
                UINT flags = p->isMulti ? MF_GRAYED : MF_STRING;
                if (p->isFixedFrame) flags |= MF_CHECKED;
                
                AppendMenuW(hm, MF_STRING, IDM_EDIT, L"編集"); 
                AppendMenuW(hm, MF_STRING, IDM_FAVORITE, p->isFavorite ? L"お気に入り解除" : L"お気に入り登録"); 
                AppendMenuW(hm, flags, IDM_TOGGLE_FIXED, L"フレーム数を固定する");
                if (g_enableTextStyle && p->isTextStyle) {
                    AppendMenuW(hm, MF_STRING, IDM_APPLY_TEXT_STYLE, L"選択中テキストにスタイル適用");
                }
                if (g_enableTextStyle && g_textStylePending) {
                    AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
                    AppendMenuW(hm, MF_STRING, IDM_TEXT_STYLE_COMMIT, L"テキストスタイルを確定");
                    AppendMenuW(hm, MF_STRING, IDM_TEXT_STYLE_CANCEL, L"テキストスタイルを取消");
                }
                AppendMenuW(hm, MF_STRING, IDM_DELETE, L"削除"); 
            }
            else {
                AppendMenuW(hm, MF_STRING, IDM_SETTINGS, L"設定");
                AppendInfoSubMenu(hm);
                if (g_enableTextStyle && g_textStylePending) {
                    AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
                    AppendMenuW(hm, MF_STRING, IDM_TEXT_STYLE_COMMIT, L"テキストスタイルを確定");
                    AppendMenuW(hm, MF_STRING, IDM_TEXT_STYLE_CANCEL, L"テキストスタイルを取消");
                }
                AppendMenuW(hm, MF_STRING, IDM_OPEN_FOLDER, L"MyAssetフォルダを開く");
            }
            POINT pt; GetCursorPos(&pt); TrackPopupMenu(hm, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL); DestroyMenu(hm);
        } return 0;
    }
    case WM_COMMAND: {
        int id = LOWORD(wp); int code = HIWORD(wp);
        if (id == IDM_SETTINGS) OpenSettings();
        else if (id == ID_BTN_REFRESH) RefreshAssets(true);
        else if (id == IDM_INFO_RELEASES) OpenReleaseNotesInBrowser(hwnd);
        else if (id == IDM_INFO_BUGREPORT) OpenBugReportInBrowser(hwnd);
        else if (id == IDM_OPEN_FOLDER) OpenAssetFolderInExplorer(hwnd);
        else if (id == IDM_EDIT && g_contextTargetIndex != -1) { Asset* p = g_displayAssets[g_contextTargetIndex]; if (p->pImage) { delete p->pImage; p->pImage = nullptr; } g_editOrgPath = p->path; g_editName = p->name; g_editCat = p->category; g_addDialogFromMyAssetAdd = false; OpenAddDialog("", true); }
        else if (id == IDM_FAVORITE && g_contextTargetIndex != -1) { ToggleFavorite(g_displayAssets[g_contextTargetIndex]->path); RefreshAssets(false); }
        else if (id == IDM_APPLY_TEXT_STYLE && g_contextTargetIndex != -1) {
            Asset* p = g_displayAssets[g_contextTargetIndex];
            if (!g_enableTextStyle || !p->isTextStyle) return 0;
            if (g_textStylePending) {
                CancelPendingTextStyle(hwnd);
            }
            ApplyTextStyleFromAssetPath(hwnd, p->path);
        }
        else if (id == IDM_TEXT_STYLE_COMMIT) {
            CommitPendingTextStyle(hwnd);
        }
        else if (id == IDM_TEXT_STYLE_CANCEL) {
            CancelPendingTextStyle(hwnd);
        }
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
                g_textStylePaths.erase(p->path);
                SaveTextStyleFlags();
                std::wstring base = p->path.substr(0, p->path.find_last_of(L"."));
                DeleteFileW((base + L".png").c_str()); DeleteFileW((base + L".gif").c_str());
                RefreshAssets(false); 
            } 
        }
        else if (id == IDM_SORT_NAME || id == IDM_SORT_FAVORITE || id == IDM_SORT_CATEGORY || id == IDM_SORT_CUSTOM) {
            if (id == IDM_SORT_NAME) g_sortMode = 0;
            else if (id == IDM_SORT_FAVORITE) g_sortMode = 1;
            else if (id == IDM_SORT_CATEGORY) g_sortMode = 2;
            else g_sortMode = 3;
            SaveConfig();
            UpdateDisplayList();
        }
        else if (id == ID_COMBO_CATEGORY && code == CBN_SELCHANGE) { int idx = SendMessage(g_hCombo, CB_GETCURSEL, 0, 0); if (idx >= 0 && idx < (int)g_categories.size()) { g_currentCategory = g_categories[idx]; UpdateDisplayList(); } }
        else if (id == ID_EDIT_SEARCH && code == EN_CHANGE) { int len = GetWindowTextLengthW(g_hSearch); std::vector<wchar_t> buf(len + 1); GetWindowTextW(g_hSearch, buf.data(), len + 1); g_searchQuery = buf.data(); UpdateDisplayList(); }
        return 0;
    }
    case WM_PAINT: { PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps); RECT rc; GetClientRect(hwnd, &rc); HDC mc = CreateCompatibleDC(hdc); HBITMAP mb = CreateCompatibleBitmap(hdc, rc.right, rc.bottom); SelectObject(mc, mb); DrawContent(mc, rc.right, rc.bottom); BitBlt(hdc, 0, 0, rc.right, rc.bottom, mc, 0, 0, SRCCOPY); DeleteObject(mb); DeleteDC(mc); EndPaint(hwnd, &ps); return 0; }
    case WM_LBUTTONDOWN: {
        SetFocus(hwnd); 

        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        if (g_enableTextStyle && g_assetViewTab == 1 && y >= rc.bottom - FOOTER_H - STYLE_ACTION_H && y < rc.bottom - FOOTER_H) {
            RECT rcCommit = {10, rc.bottom - FOOTER_H - STYLE_ACTION_H + 6, 100, rc.bottom - FOOTER_H - 6};
            RECT rcCancel = {106, rc.bottom - FOOTER_H - STYLE_ACTION_H + 6, 196, rc.bottom - FOOTER_H - 6};
            RECT rcDetail = {202, rc.bottom - FOOTER_H - STYLE_ACTION_H + 6, 292, rc.bottom - FOOTER_H - 6};
            POINT p = {x, y};
            if (PtInRect(&rcCommit, p)) {
                CommitPendingTextStyle(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            }
            if (PtInRect(&rcCancel, p)) {
                CancelPendingTextStyle(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            }
            if (PtInRect(&rcDetail, p)) {
                OpenTextStyleQuickDialog(hwnd);
                return 0;
            }
        }
        int listTop = GetListTopY();
        if (g_enableTextStyle) {
            RECT rcTabAll = {10, TITLE_H + 4, 130, listTop - 4};
            RECT rcTabStyle = {136, TITLE_H + 4, 276, listTop - 4};
            POINT pt = {x, y};
            if (PtInRect(&rcTabAll, pt)) {
                if (g_assetViewTab != 0) { g_assetViewTab = 0; UpdateDisplayList(); }
                return 0;
            }
            if (PtInRect(&rcTabStyle, pt)) {
                if (g_assetViewTab != 1) { g_assetViewTab = 1; UpdateDisplayList(); }
                return 0;
            }
        }
        if (y > listTop && y < rc.bottom - GetBottomReservedHeight()) {
            int cols = (std::max)(1, (int)rc.right / GetListItemMinWidth()); int cw = (int)rc.right / cols;
            int idx = ((y - listTop + g_scrollY) / GetListItemHeight()) * cols + (x / cw);
            if (idx >= 0 && idx < (int)g_displayAssets.size()) {
                if (g_enableTextStyle && g_assetViewTab == 1) {
                    Asset* p = g_displayAssets[idx];
                    if (!p->isTextStyle) return 0;
                    if (g_textStylePending) {
                        CancelPendingTextStyle(hwnd);
                    }
                    ApplyTextStyleFromAssetPath(hwnd, p->path);
                    InvalidateRect(hwnd, NULL, FALSE);
                    return 0;
                }
                bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                if (ctrlDown && CanStartCustomReorder()) {
                    SetCapture(hwnd);
                    g_selectedIndex = idx;
                    g_isReorderDrag = true;
                    g_reorderFromIndex = idx;
                    g_reorderInsertPos = GetInsertPosFromPoint(x, y, rc);
                    InvalidateRect(hwnd, nullptr, FALSE);
                    return 0;
                }
                SetCapture(hwnd);
                g_selectedIndex = idx;
                g_isDragCheck = true;
                g_dragStartPt = {x, y};
                InvalidateRect(hwnd, nullptr, FALSE);
            }
        } 
        else if (y >= rc.bottom - FOOTER_H) {
            if (x >= rc.right - 70 && x <= rc.right - 10) RefreshAssets(true);
            else if (x >= rc.right - 140 && x <= rc.right - 80) {
                HMENU hmSort = CreatePopupMenu();
                UINT f0 = MF_STRING | (g_sortMode == 0 ? MF_CHECKED : 0);
                UINT f1 = MF_STRING | (g_sortMode == 1 ? MF_CHECKED : 0);
                UINT f2 = MF_STRING | (g_sortMode == 2 ? MF_CHECKED : 0);
                UINT f3 = MF_STRING | (g_sortMode == 3 ? MF_CHECKED : 0);
                AppendMenuW(hmSort, f1, IDM_SORT_FAVORITE, L"お気に入り優先");
                AppendMenuW(hmSort, f0, IDM_SORT_NAME, L"名前順");
                AppendMenuW(hmSort, f2, IDM_SORT_CATEGORY, L"カテゴリ順");
                AppendMenuW(hmSort, f3, IDM_SORT_CUSTOM, L"カスタム");
                POINT pt; GetCursorPos(&pt);
                TrackPopupMenu(hmSort, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
                DestroyMenu(hmSort);
            }
        }
        return 0;
    }
    case WM_MOUSELEAVE: { 
        if (g_hoverIndex != -1) { KillTimer(hwnd, ID_TIMER_HOVER); g_hoverIndex = -1; InvalidateRect(hwnd, NULL, FALSE); } 
        g_previewAllLastTick = 0;
        g_previewHoverLastTick = 0;
        KillTimer(hwnd, ID_TIMER_TOOLTIP);
        g_tooltipTargetIndex = -1;
        g_isMouseTracking = false; 
        HideMyTooltip(); 
        return 0; 
    }
    case WM_LBUTTONUP: { 
        if (g_isReorderDrag) {
            bool moved = ApplyCustomReorderFromInsertPos();
            if (moved && CanStartCustomReorder()) {
                SaveCurrentDisplayOrderForCategory();
            }
            g_isReorderDrag = false;
            g_reorderFromIndex = -1;
            g_reorderInsertPos = -1;
            ReleaseCapture();
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        ReleaseCapture(); 
        g_isDragCheck = false; return 0; 
    }
    case WM_MOUSEMOVE: {
        int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp); RECT rc; GetClientRect(hwnd, &rc);
        if (!g_isMouseTracking) { TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT), TME_LEAVE, hwnd, 0 }; TrackMouseEvent(&tme); g_isMouseTracking = true; }
        bool pausePreviewForReorder = CanStartCustomReorder() && (((GetKeyState(VK_CONTROL) & 0x8000) != 0) || g_isReorderDrag);
        if (pausePreviewForReorder) {
            KillTimer(hwnd, ID_TIMER_HOVER);
            g_previewAllLastTick = 0;
            g_previewHoverLastTick = 0;
        }

        if (g_isReorderDrag) {
            int insertPos = GetInsertPosFromPoint(x, y, rc);
            if (insertPos != g_reorderInsertPos) {
                g_reorderInsertPos = insertPos;
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;
        }
        
        bool onMark = false;
        int listTop = GetListTopY();
        bool mouseInList = (y > listTop && y < rc.bottom - GetBottomReservedHeight() && x >= 0 && x < rc.right);
        if (mouseInList) {
            if (!pausePreviewForReorder && g_previewPlaybackMode == 1) {
                if (g_previewAllLastTick == 0) g_previewAllLastTick = GetTickCount64();
                SetTimer(hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
            }
            int itemH = GetListItemHeight();
            int cols = (std::max)(1, (int)rc.right / GetListItemMinWidth()); int cw = (int)rc.right / cols;
            int row = (y - listTop + g_scrollY) / itemH; int col = x / cw;
            int idx = row * cols + col;
            
            if (idx >= 0 && idx < (int)g_displayAssets.size()) { 
                if (g_hoverIndex != idx) {
                    g_hoverIndex = idx;
                    if (!pausePreviewForReorder && g_previewPlaybackMode == 0) {
                        g_previewHoverLastTick = 0;
                        SetTimer(hwnd, ID_TIMER_HOVER, GetPreviewAllTickMs(), NULL);
                    }
                }

                int itemX = col * cw; int itemY = row * itemH + (listTop - g_scrollY);
                POINT pt = {x, y};

                if (g_previewThumbOnly) {
                    int thumbW = GetThumbWScaled();
                    int thumbH = GetThumbHScaled();
                    RECT rcThumb = {
                        itemX + (cw - thumbW) / 2,
                        itemY + (itemH - thumbH) / 2,
                        itemX + (cw - thumbW) / 2 + thumbW,
                        itemY + (itemH - thumbH) / 2 + thumbH
                    };
                    if (PtInRect(&rcThumb, pt)) {
                        onMark = true;
                        if (g_tooltipTargetIndex != idx) {
                            HideMyTooltip();
                            g_tooltipTargetIndex = idx;
                            SetTimer(hwnd, ID_TIMER_TOOLTIP, 1000, NULL);
                        }
                    }
                } else {
                    int right = itemX + cw - 2; int bottom = itemY + itemH - 2;
                    RECT rcMark = { right - 30, bottom - 30, right, bottom };
                    if (PtInRect(&rcMark, pt)) {
                        onMark = true;
                        if (g_tooltipTargetIndex != idx) {
                            HideMyTooltip();
                            g_tooltipTargetIndex = idx;
                            SetTimer(hwnd, ID_TIMER_TOOLTIP, 1000, NULL);
                        }
                    }
                }
            } else {
                g_hoverIndex = -1;
                if (g_previewPlaybackMode == 0) {
                    KillTimer(hwnd, ID_TIMER_HOVER);
                    g_previewHoverLastTick = 0;
                }
            }
        } else {
            if (g_hoverIndex != -1) g_hoverIndex = -1;
            KillTimer(hwnd, ID_TIMER_HOVER);
            g_previewAllLastTick = 0;
            g_previewHoverLastTick = 0;
        }
        
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
        g_pendingLayerApiCompare = false;
        g_pendingLayerApiCompareObjs.clear();
        g_pendingAddObjSnapshots.clear();
        g_textStylePending = false;
        g_textStylePendingLayer = -1;
        g_textStylePendingStart = -1;
        g_textStylePendingEnd = -1;
        g_textStyleOriginalAlias.clear();
        if (!g_lastTempPath.empty()) DeleteFileW(g_lastTempPath.c_str());
        if(g_hDlg) DestroyWindow(g_hDlg);
        if(g_hSettingDlg) DestroyWindow(g_hSettingDlg);
        if(g_hTextStyleQuickDlg) DestroyWindow(g_hTextStyleQuickDlg);
        if(g_hSnipWnd) DestroyWindow(g_hSnipWnd);
        if(g_hInfoWnd) DestroyWindow(g_hInfoWnd);
        if(g_hTooltip) DestroyWindow(g_hTooltip); 
        if (g_hBrInputBg) { DeleteObject(g_hBrInputBg); g_hBrInputBg = nullptr; }
        if (g_hBrBg) { DeleteObject(g_hBrBg); g_hBrBg = nullptr; }
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

    // API方式:
    // 選択オブジェクトとaliasを一旦保持し、コールバック外で安全に処理する。
    if (g_hwnd && IsWindow(g_hwnd) && edit->get_selected_object && edit->get_object_alias) {
        g_pendingAddObjSnapshots.clear();
        g_pendingAddObjSnapshots.reserve((size_t)count);
        for (int i = 0; i < count; ++i) {
            void* obj = edit->get_selected_object(i);
            if (!obj) continue;
            LPCSTR raw = edit->get_object_alias(obj);
            PendingAddObjSnapshot s = {};
            s.obj = obj;
            if (raw) s.alias = raw;
            g_pendingAddObjSnapshots.push_back(std::move(s));
        }
        if (!g_pendingAddObjSnapshots.empty()) {
            PostMessageW(g_hwnd, WM_APP_BUILD_ADD_ASSET, 0, 0);
            return;
        }
    }

    // ここから下は旧フォールバック経路（APIスナップショット失敗時のみ）。
    g_addDialogRangeValid = false;
    g_addDialogRangeStart = 0;
    g_addDialogRangeEnd = 0;

    int memStart = INT_MAX;
    int memEnd = INT_MIN;
    bool memRangeFound = false;
    for (int i = 0; i < count; i++) {
        void* obj = edit->get_selected_object(i);
        if (!obj) continue;
        int s = 0, e = 0;
        if (TryGetObjectFrameRangeFromMemory(obj, s, e)) {
            memStart = (std::min)(memStart, s);
            memEnd = (std::max)(memEnd, e);
            memRangeFound = true;
        }
    }
    
    std::string bodyData;
    int minLayer = 99999;

    for(int i = 0; i < count; i++) {
        void* obj = edit->get_selected_object(i);
        if (obj) {
            LPCSTR raw = edit->get_object_alias(obj);
            int layer = ResolveObjectLayerForAssetBuild(obj, raw);
            if (layer < minLayer) minLayer = layer;
        }
    }
    if (minLayer == 99999) minLayer = 1;
    
    if (count == 1) {
        void* obj = edit->get_selected_object(0);
        if (obj) {
            LPCSTR raw = edit->get_object_alias(obj);
            if (raw) bodyData = raw;
        }
    } 
    else {
        std::vector<void*> selectedObjs;
        if (g_enableGroupDebugDump) {
            selectedObjs.reserve((size_t)count);
            for (int i = 0; i < count; ++i) {
                selectedObjs.push_back(edit->get_selected_object(i));
            }
        }

        std::vector<int> relativeGroups(count, 0);
        std::vector<int> debugResolvedLayers(count, 0);
        std::map<long long, int> groupIdMap;
        std::map<long long, int> groupCounts;
        std::vector<long long> originalGroups(count, 0);
        std::vector<int> frameStart(count, 0), frameEnd(count, 0);
        std::vector<bool> hasFrame(count, false);
        int nextGroupId = 1;
        bool hasGroupToken = false;
        SYSTEMTIME st = {};
        GetLocalTime(&st);
        if (g_enableGroupDebugDump) {
            std::ostringstream os;
            os << "[OnAddAsset] " << st.wYear << "-" << st.wMonth << "-" << st.wDay
               << " " << st.wHour << ":" << st.wMinute << ":" << st.wSecond
               << " selected=" << count;
            AppendGroupDebugLine(os.str());
            g_pendingLayerApiCompareObjs = selectedObjs;
            DumpGroupMemoryProbe(selectedObjs);
            DumpGroupIdCandidateProbe(selectedObjs);
            DumpLayerOffsetProbe(selectedObjs);
            g_pendingLayerApiCompare = true;
            if (g_hwnd && IsWindow(g_hwnd)) {
                SetTimer(g_hwnd, ID_TIMER_LAYER_API_COMPARE, 1, NULL);
                AppendGroupDebugLine("  layer-api compare trigger=timer");
            } else {
                AppendGroupDebugLine("  layer-api compare trigger skipped: g_hwnd unavailable");
            }
        }

        for (int i = 0; i < count; i++) {
            void* obj = edit->get_selected_object(i);
            if (!obj) continue;
            LPCSTR raw = edit->get_object_alias(obj);
            int aliasLayer = raw ? ExtractHeaderIntValue(raw, "layer=") : 0;
            int memLayer = GetObjectLayerIndex(obj);
            const char* layerSource = nullptr;
            int resolvedLayer = ResolveObjectLayerForAssetBuild(obj, raw, &layerSource);
            debugResolvedLayers[i] = resolvedLayer;

            int aliasGroupId = raw ? ExtractHeaderIntValue(raw, "group=") : 0;
            long long originalGroupId = aliasGroupId;
            int verifiedGroupId = 0;
            if (originalGroupId <= 0) {
                verifiedGroupId = GetVerifiedTimelineGroupId(obj);
                if (verifiedGroupId > 0) originalGroupId = (long long)verifiedGroupId;
            }
            uintptr_t ptr28 = 0;
            bool hasPtr28 = SafeReadUIntPtrAtOffset(obj, 0x28, ptr28);
            if (originalGroupId <= 0 && hasPtr28 && ptr28 != 0) {
                originalGroupId = (long long)ptr28;
            }
            originalGroups[i] = originalGroupId;
            if (raw) {
                hasFrame[i] = ExtractHeaderFrameRange(raw, frameStart[i], frameEnd[i]);
            }
            if (originalGroupId > 0) {
                hasGroupToken = true;
                groupCounts[originalGroupId]++;
            }

            if (g_enableGroupDebugDump) {
                std::ostringstream os;
                os << "  idx=" << i
                   << " layerResolved=" << resolvedLayer
                   << " layerSrc=" << (layerSource ? layerSource : "?")
                   << " layerAlias=" << aliasLayer
                   << " layerMem0x60=" << memLayer
                   << " aliasGroup=" << aliasGroupId;
                os << " verifiedGroup=" << verifiedGroupId;
                if (hasPtr28) {
                    os << " ptr28=0x" << std::hex << (unsigned long long)ptr28 << std::dec;
                } else {
                    os << " ptr28=(unreadable)";
                }
                os
                   << " hasFrame=" << (hasFrame[i] ? 1 : 0)
                   << " frame=" << (hasFrame[i] ? std::to_string(frameStart[i]) + "," + std::to_string(frameEnd[i]) : std::string("-"))
                   << " chosen=" << originalGroupId
                   << " fallback=0";
                AppendGroupDebugLine(os.str());
                DumpObjectMemoryWide(obj, i);
            }

            if (g_enableGroupDebugDump && raw) {
                std::ostringstream head;
                head << "  ---- alias dump begin idx=" << i << " ----";
                AppendGroupDebugLine(head.str());
                std::string rawAlias = std::string(raw);
                AppendGroupDebugText(rawAlias);
                AppendGroupDebugLine("");
                AppendGroupDebugLine("  ---- alias dump end ----");
                DumpApiItemValuesFromAlias(edit, obj, rawAlias, i);
            } else if (g_enableGroupDebugDump) {
                AppendGroupDebugLine("  ---- alias dump skipped (raw=null) ----");
            }
        }

        if (g_enableGroupDebugDump && count > 1) {
            std::set<int> uniq;
            for (int i = 0; i < count; ++i) {
                if (debugResolvedLayers[i] > 0) uniq.insert(debugResolvedLayers[i]);
            }
            std::ostringstream os;
            os << "  layer-check unique=" << uniq.size();
            if (uniq.size() == 1) os << " [warn: all selected objects resolved to same layer]";
            AppendGroupDebugLine(os.str());
        }

        if (!hasGroupToken) {
            if (g_enableGroupDebugDump) {
                AppendGroupDebugLine("  no reliable group token -> no group= will be written (safe mode)");
            }
        } else {
            for (int i = 0; i < count; i++) {
                long long originalGroupId = originalGroups[i];
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
        }

        bool shouldEmitGroupField = false;
        for (int i = 0; i < count; ++i) {
            if (relativeGroups[i] > 0) {
                shouldEmitGroupField = true;
                break;
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
                    int currentLayer = ResolveObjectLayerForAssetBuild(obj, raw);
                    int relativeLayer = (currentLayer - minLayer) + 1;
                    ReplaceStringAll(s, "[Object.", sectionNew);
                    ReplaceStringAll(s, "[Object]", headerNew);
                    RemoveLinesStartingWith(s, "group=");
                    size_t pos = s.find(headerNew);
                    if (pos != std::string::npos) {
                        std::string headerLines = "\r\nlayer=" + std::to_string(relativeLayer);
                        if (relativeGroups[i] > 0) {
                            headerLines += "\r\ngroup=" + std::to_string(relativeGroups[i]);
                        } else if (shouldEmitGroupField) {
                            headerLines += "\r\ngroup=0";
                        }
                        s.insert(pos + headerNew.length(), headerLines);
                    }
                    bodyData += s + "\r\n"; 
                } 
            } 
        }
    }
    int parsedStart = 0, parsedEnd = 0;
    if (TryParseFrameRangeFromAliasData(bodyData, parsedStart, parsedEnd)) {
        g_addDialogRangeValid = true;
        g_addDialogRangeStart = parsedStart;
        g_addDialogRangeEnd = parsedEnd;
    } else if (memRangeFound) {
        // aliasから取れない場合のみメモリ推定値を使う
        g_addDialogRangeValid = true;
        g_addDialogRangeStart = memStart;
        g_addDialogRangeEnd = memEnd;
    }
    g_addDialogFromMyAssetAdd = true;
    OpenAddDialog(bodyData, false);
}

static void EnsureMainWindowVisible(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) return;
    RECT rc = {};
    if (!GetWindowRect(hwnd, &rc)) return;

    HMONITOR mon = MonitorFromRect(&rc, MONITOR_DEFAULTTONULL);
    if (mon) return;

    RECT wa = {};
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa, 0);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    if (w < 200) w = 360;
    if (h < 200) h = 550;
    int x = wa.left + ((wa.right - wa.left) - w) / 2;
    int y = wa.top + ((wa.bottom - wa.top) - h) / 2;
    if (x < wa.left) x = wa.left;
    if (y < wa.top) y = wa.top;
    SetWindowPos(hwnd, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
}

void OnShowMainWindow(EDIT_SECTION_SAFE* edit) {
    {
        std::ostringstream os;
        os << "OnShowMainWindow called: suppress=" << (g_suppressMainShow ? 1 : 0)
           << " g_hwnd=" << (void*)g_hwnd
           << " visible=" << (g_hwnd && IsWindowVisible(g_hwnd) ? 1 : 0);
        AppendUiDebugLine(os.str());
    }
    if (g_suppressMainShow) {
        AppendUiDebugLine("OnShowMainWindow ignored (suppressed)");
        return;
    }
    if (!g_hwnd || !IsWindow(g_hwnd)) {
        CreatePluginWindow();
    }
    if (!g_hwnd || !IsWindow(g_hwnd)) return;

    EnsureMainWindowVisible(g_hwnd);
    ShowWindow(g_hwnd, IsIconic(g_hwnd) ? SW_RESTORE : SW_SHOW);
    SetForegroundWindow(g_hwnd);
}

void CreatePluginWindow() { 
    LoadWindowConfig(); 
    WNDCLASSW wc = {0}; 
    wc.lpfnWndProc = WndProc; 
    wc.hInstance = g_hInst; 
    wc.lpszClassName = L"MyAssetMgr_Modern"; 
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW); 
    wc.hbrBackground = g_hBrBg; 
    wc.style = CS_HREDRAW | CS_VREDRAW; 
    RegisterClassW(&wc); 
    HWND hostMain = FindHostMainWindow();
    g_hwnd = CreateWindowExW(
        WS_EX_TOPMOST, L"MyAssetMgr_Modern", L"My Asset Manager",
        WS_POPUP | WS_CLIPCHILDREN | WS_VSCROLL,
        g_winX, g_winY, g_winW, g_winH,
        hostMain, nullptr, g_hInst, nullptr);
}

extern "C" __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host) { 
    HMODULE hm; GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)RegisterPlugin, &hm); g_hInst = (HINSTANCE)hm;
    g_host = host; host->set_plugin_information(L"My Asset Manager v" MYASSET_VERSION_W); 
    if (host->create_edit_handle) g_editHandle = (EDIT_HANDLE_SAFE*)host->create_edit_handle();
    if (host->f2) {
        auto registerOutput = (void(*)(OUTPUT_PLUGIN_TABLE*))host->f2;
        registerOutput(&g_outputPluginTable);
    }
    RegisterCustomDialogs(); CreatePluginWindow(); 
    if (g_hwnd) host->register_window_client(L"My Asset Manager", g_hwnd); 
    if (host->register_edit_menu) host->register_edit_menu(L"表示 > My Asset Manager", OnShowMainWindow); 
    if (host->register_layer_menu) { host->register_layer_menu(L"MyAsset：追加", OnAddAsset); }
    if (host->register_object_menu) { host->register_object_menu(L"MyAsset：追加", OnAddAsset); }
}
