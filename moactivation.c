// ============================================================================
// MoActivation —— Windows/Office 便捷工具
// 编译：见 build.bat（MSVC: cl + rc）
// ============================================================================
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <dwmapi.h>
#include <string.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "msvcrt.lib")

// ---- 调色板（Win11 扁平白底；COLORREF 0x00BBGGRR） ----
#define C_BG      0x00F6F6F6
#define C_PANEL   0x00FFFFFF
#define C_TEXT    0x001E1E1E
#define C_SUB     0x006B6B6B
#define C_ACCENT  0x000067C0
#define C_BORDER  0x00E8E8E8
#define C_WHITE   0x00FFFFFF
#define C_HOVER   0x00F0F0F0

// ---- 主窗口布局（逻辑单位） ----
#define WIN_W    560
#define WIN_H    300
#define BTN_W    248
#define BTN_H    96
#define BTN_GAP  16
#define SB_W     44   // 顶栏右上角系统按钮宽度

// ---- 组件选择窗布局（双列网格） ----
#define PICK_W        540
#define PICK_H        400
#define PICK_TITLE_H  44
#define APP_BOX       22
#define APP_ROW_H     34
#define PICK_PAD      24
#define PICK_COL_GAP  48
#define APP_COUNT     9

// ---- 消息 ----
#define WM_DONE       (WM_APP + 1)
#define WM_PICK_DONE  (WM_APP + 2)
#define CREATE_NO_WINDOW  0x08000000
#define CREATE_NEW_CONSOLE 0x00000010

// ---- ODT 组件名（与 <ExcludeApp ID> 一致） ----
static const wchar_t *gApps[APP_COUNT] = {
    L"Word", L"Excel", L"PowerPoint", L"Outlook", L"OneNote",
    L"Access", L"Publisher", L"Teams", L"OneDrive",
};

// ---- 全局 UI 状态 ----
static float   g_dpi   = 1.0f;
static HFONT   g_fBase = NULL;
static HFONT   g_fBold = NULL;
static wchar_t gExeDir[MAX_PATH];

// 后台执行状态
static volatile LONG gExecBusy   = 0;
static volatile int  gExecTag    = 0; // 0=Office 安装，1=MAS 激活
static wchar_t       gResult[512] = L"";
static wchar_t       gStatus[256] = L"选择要执行的操作";
static int           gHoverBtn = -1; // 顶栏系统按钮悬停：-1 无，0 最小化，1 最大化/还原，2 关闭
static BOOL          gTracking = FALSE;

// 执行线程传参（仅单个执行存在时访问，busy 保护）
static const wchar_t *gThreadEx[APP_COUNT];
static int gThreadExCount = 0;

// 组件选择窗状态
static BOOL    gPickChecked[APP_COUNT] = { TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE };
static BOOL    gPickConfirmed = FALSE;
static HWND    gPickMain = NULL;
static HWND    gPickHwnd = NULL;
static BOOL    gPickOpen = FALSE;

// ---- 工具函数 ----
static int px(int v) { return (int)(v * g_dpi + 0.5f); }

static COLORREF col(DWORD v) { return (COLORREF)v; }

// 顶栏系统按钮命中：返回 0=最小化 1=最大化/还原 2=关闭 -1=未命中
static int sb_hit(int cw, int bar_h, int x, int y) {
    if (y < 0 || y >= bar_h) return -1;
    int w = px(SB_W);
    int x0 = cw - w * 3;
    if (x < x0 || x >= cw) return -1;
    return (x - x0) / w;
}

// 仅重绘右上角系统按钮区域，避免整窗闪烁
static void inv_sys(HWND hwnd, int bar_h) {
    RECT rc; GetClientRect(hwnd, &rc);
    RECT r = { rc.right - px(SB_W) * 3, 0, rc.right, bar_h };
    InvalidateRect(hwnd, &r, FALSE);
}

static HFONT make_font(int logical, BOOL bold) {
    LOGFONTW lf;
    memset(&lf, 0, sizeof(lf));
    // 字号固定用逻辑值；窗口坐标/按钮仍按 DPI 缩放
    lf.lfHeight = -logical;
    lf.lfWeight = bold ? FW_SEMIBOLD : FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lstrcpyW(lf.lfFaceName, L"Segoe UI");
    return CreateFontIndirectW(&lf);
}

static void fill(HDC hdc, int x, int y, int w, int h, COLORREF c) {
    if (w <= 0 || h <= 0) return;
    RECT r = { x, y, x + w, y + h };
    HBRUSH b = CreateSolidBrush(c);
    FillRect(hdc, &r, b);
    DeleteObject(b);
}

// 画一个 11x11 的方框（用于最大化/还原图标）
static void frame(HDC hdc, int cx, int cy, int s) {
    MoveToEx(hdc, cx - s, cy - s, NULL);
    LineTo(hdc, cx + s, cy - s);
    LineTo(hdc, cx + s, cy + s);
    LineTo(hdc, cx - s, cy + s);
    LineTo(hdc, cx - s, cy - s);
}

// 顶栏右上角：最小化 / 最大化(还原) / 关闭，Win11 扁平风格
static void rounded_fill(HDC hdc, int x, int y, int w, int h, COLORREF c, int r);
static void draw_sysbtns(HDC hdc, HWND hwnd, int cw, int bar_h) {
    int w = px(SB_W);
    int x0 = cw - w * 3;
    int cx0 = x0 + w / 2;
    int cy = bar_h / 2;
    BOOL zoom = IsZoomed(hwnd);

    for (int i = 0; i < 3; i++) {
        int x = x0 + i * w;
        int cx = cx0 + i * w;
        // 悬停背景：贴合图标的小圆角块，居中
        if (gHoverBtn == i) {
            int hw2 = px(15);
            rounded_fill(hdc, cx - hw2, cy - hw2, hw2 * 2, hw2 * 2, col(C_HOVER), px(6));
        }
        HPEN pen = CreatePen(PS_SOLID, 1, col(C_TEXT));
        HPEN op = (HPEN)SelectObject(hdc, pen);
        if (i == 0) { // 最小化：一条横线
            MoveToEx(hdc, cx - 5, cy, NULL);
            LineTo(hdc, cx + 5, cy);
        } else if (i == 1) { // 最大化 / 还原
            if (zoom) { // 还原：两个重叠方框
                frame(hdc, cx, cy, 5);
                MoveToEx(hdc, cx - 3, cy - 3, NULL);
                LineTo(hdc, cx + 3, cy - 3);
                LineTo(hdc, cx + 3, cy + 3);
                LineTo(hdc, cx - 3, cy + 3);
                LineTo(hdc, cx - 3, cy - 3);
            } else {
                frame(hdc, cx, cy, 5);
            }
        } else { // 关闭：叉号
            MoveToEx(hdc, cx - 5, cy - 5, NULL);
            LineTo(hdc, cx + 5, cy + 5);
            MoveToEx(hdc, cx + 5, cy - 5, NULL);
            LineTo(hdc, cx - 5, cy + 5);
        }
        SelectObject(hdc, op);
        DeleteObject(pen);
    }
}

static void rounded_fill(HDC hdc, int x, int y, int w, int h, COLORREF c, int r) {
    if (w <= 0 || h <= 0) return;
    HBRUSH b = CreateSolidBrush(c);
    HPEN   p = CreatePen(PS_SOLID, 1, c);
    HBRUSH ob = (HBRUSH)SelectObject(hdc, b);
    HPEN   op = (HPEN)SelectObject(hdc, p);
    RoundRect(hdc, x, y, x + w, y + h, r, r);
    SelectObject(hdc, op);
    SelectObject(hdc, ob);
    DeleteObject(p);
    DeleteObject(b);
}

// 对勾用圆角几何笔，拐角才平滑
static void draw_check(HDC hdc, int cx, int cy, int s, COLORREF c) {
    UINT width = (UINT)(s / 6.0f + 0.5f);
    if (width < 2) width = 2;
    DWORD style = PS_GEOMETRIC | PS_SOLID | PS_JOIN_ROUND;
    LOGBRUSH lb;
    lb.lbStyle = BS_SOLID;
    lb.lbColor = c;
    lb.lbHatch = 0;
    HPEN p = ExtCreatePen((DWORD)style, width, &lb, 0, NULL);
    HPEN op = (HPEN)SelectObject(hdc, p);
    int m = s / 10;
    MoveToEx(hdc, cx - s / 2 + m * 2, cy - s / 6, NULL);
    LineTo(hdc, cx - s / 2 + m * 5, cy + s / 5);
    LineTo(hdc, cx + s / 2 - m * 2, cy - s / 4);
    SelectObject(hdc, op);
    DeleteObject(p);
}

// 按钮文字精确居中
static void btn_text(HDC hdc, HFONT f, int x, int y, int w, int h, const wchar_t *s, COLORREF c) {
    int len = lstrlenW(s);
    if (len == 0 || w <= 0 || h <= 0) return;
    HFONT of = (HFONT)SelectObject(hdc, f);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, c);
    SIZE sz;
    GetTextExtentPoint32W(hdc, s, len, &sz);
    int tx = x + (w - sz.cx) / 2;
    int ty = y + (h - sz.cy) / 2;
    if (tx < x) tx = x;
    if (ty < y) ty = y;
    TextOutW(hdc, tx, ty, s, len);
    SelectObject(hdc, of);
}

static void status_text(HDC hdc, HFONT f, int x, int y, int w, int h, const wchar_t *s, COLORREF c) {
    int len = lstrlenW(s);
    if (len == 0 || w <= 0) return;
    HFONT of = (HFONT)SelectObject(hdc, f);
    SetTextColor(hdc, c);
    SetBkMode(hdc, TRANSPARENT);
    RECT r = { x, y, x + w, y + h };
    DrawTextW(hdc, (LPWSTR)s, len, &r, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    SelectObject(hdc, of);
}

// ---- 引擎（PowerShell 子进程） ----
static void run_powershell(const wchar_t *psCmd, DWORD flags, BOOL wait) {
    wchar_t cmd[8192];
    wsprintfW(cmd, L"powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"%ls\"", psCmd);

    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    memset(&si, 0, sizeof(si));
    memset(&pi, 0, sizeof(pi));
    si.cb = sizeof(si);
    wchar_t cmdBuf[8192];
    lstrcpyW(cmdBuf, cmd);

    if (!CreateProcessW(NULL, cmdBuf, NULL, NULL, FALSE, flags, NULL, gExeDir[0] ? gExeDir : NULL, &si, &pi)) {
        wsprintfW(gResult, L"启动 PowerShell 失败");
        return;
    }
    if (wait) {
        WaitForSingleObject(pi.hProcess, INFINITE);
        DWORD code = 0;
        GetExitCodeProcess(pi.hProcess, &code);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        wsprintfW(gResult, code == 0 ? L"完成" : L"PowerShell 退出码 %lu", code);
    }
}

static DWORD WINAPI ThreadOffice(LPVOID unused) {
    (void)unused;
    // 生成 ODT config.xml（<ExcludeApp> 排除未勾选项）
    wchar_t cfg[2048];
    wchar_t tmp[128];
    lstrcpyW(cfg, L"<Configuration><Add OfficeClientEdition=\"64\" Channel=\"Current\"><Product ID=\"O365ProPlusRetail\"><Language ID=\"zh-cn\"/>");
    for (int i = 0; i < gThreadExCount; i++) {
        wsprintfW(tmp, L"<ExcludeApp ID=\"%ls\"/>", gThreadEx[i]);
        lstrcatW(cfg, tmp);
    }
    lstrcatW(cfg, L"</Product></Add><Display Level=\"Full\" AcceptEULA=\"TRUE\"/></Configuration>");

    wchar_t ps[10240];
    wsprintfW(ps,
        L"$d = \"$env:TEMP\\odt_install\"; New-Item -ItemType Directory -Force -Path $d | Out-Null; "
        L"Invoke-WebRequest -Uri 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -OutFile \"$d\\setup.exe\" -UseBasicParsing; "
        L"Set-Content -LiteralPath \"$d\\config.xml\" -Value '%ls' -Encoding ASCII; "
        L"Push-Location $d; & .\\setup.exe /configure config.xml; $c = $LASTEXITCODE; Pop-Location; exit $c",
        cfg);

    run_powershell(ps, CREATE_NO_WINDOW, TRUE);

    wchar_t done[256];
    if (!lstrcmpW(gResult, L"完成"))
        lstrcpyW(done, L"安装完成");
    else {
        wsprintfW(done, L"安装失败：%ls", gResult);
    }
    lstrcpyW(gResult, done);
    InterlockedExchange(&gExecBusy, 0);
    PostMessageW(gPickMain, WM_DONE, 0, 0);
    return 0;
}

static DWORD WINAPI ThreadActivate(LPVOID unused) {
    (void)unused;
    run_powershell(L"irm https://gitee.com/cmontage/mas-cn/raw/main/GETMASCN.ps1 | iex",
                   CREATE_NEW_CONSOLE, TRUE);
    wchar_t done[256];
    if (!lstrcmpW(gResult, L"完成"))
        lstrcpyW(done, L"激活窗口已关闭");
    else {
        wsprintfW(done, L"激活失败：%ls", gResult);
    }
    lstrcpyW(gResult, done);
    InterlockedExchange(&gExecBusy, 0);
    PostMessageW(gPickMain, WM_DONE, 0, 0);
    return 0;
}

// 关闭组件选择窗并回传结果
static void finish_picker(HWND pick, BOOL confirmed) {
    gPickConfirmed = confirmed;
    gPickOpen = FALSE;
    ShowWindow(pick, SW_HIDE);
    PostMessageW(gPickMain, WM_PICK_DONE, 0, 0);
}

// ---- 组件选择窗绘制 ----
static void p_draw(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    int cw = rc.right - rc.left, ch = rc.bottom - rc.top;

    fill(hdc, 0, 0, cw, ch, col(C_BG));
    int bar_h = px(PICK_TITLE_H);
    fill(hdc, 0, 0, cw, bar_h, col(C_PANEL));

    // 标题（自绘）+ 右上角系统按钮
    RECT tr = { px(20), 0, cw - px(20) - px(SB_W) * 3, bar_h };
    HFONT of = (HFONT)SelectObject(hdc, g_fBold);
    SetTextColor(hdc, col(C_TEXT));
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, L"选择要安装的 Office 组件", -1, &tr, DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);
    draw_sysbtns(hdc, hwnd, cw, bar_h);

    int y_top = px(PICK_TITLE_H) + px(10);
    int sb_h = px(30);
    int a_x = px(20), a_w = px(64);
    int n_x = a_x + a_w + px(12), n_w = px(76);
    rounded_fill(hdc, a_x, y_top, a_w, sb_h, col(C_ACCENT), 8);
    btn_text(hdc, g_fBase, a_x, y_top, a_w, sb_h, L"全选", col(C_WHITE));
    rounded_fill(hdc, n_x, y_top, n_w, sb_h, col(C_ACCENT), 8);
    btn_text(hdc, g_fBase, n_x, y_top, n_w, sb_h, L"全不选", col(C_WHITE));

    // 组件双列网格
    int row_h = px(APP_ROW_H), box_ = px(APP_BOX);
    int pad = px(PICK_PAD), col_gap = px(PICK_COL_GAP);
    int colw = (cw - pad * 2 - col_gap) / 2;
    int y_rows = y_top + px(46);
    for (int i = 0; i < APP_COUNT; i++) {
        int rx = (i % 2 == 0) ? pad : pad + colw + col_gap;
        int ry = y_rows + (i / 2) * row_h;
        int by = ry + (row_h - box_) / 2;
        if (gPickChecked[i]) {
            fill(hdc, rx, by, box_, box_, col(C_ACCENT));
            draw_check(hdc, rx + box_ / 2, by + box_ / 2, box_, col(C_WHITE));
        } else {
            fill(hdc, rx, by, box_, box_, col(C_PANEL));
            HPEN p = CreatePen(PS_SOLID, 1, col(C_BORDER));
            HPEN op = (HPEN)SelectObject(hdc, p);
            HBRUSH b = CreateSolidBrush(col(C_PANEL));
            HBRUSH ob = (HBRUSH)SelectObject(hdc, b);
            RoundRect(hdc, rx, by, rx + box_, by + box_, 5, 5);
            SelectObject(hdc, ob);
            SelectObject(hdc, op);
            DeleteObject(b);
            DeleteObject(p);
        }
        int lx = rx + box_ + px(10);
        status_text(hdc, g_fBase, lx, ry, colw - box_ - px(10), row_h, gApps[i], col(C_TEXT));
    }

    // 确定 / 取消
    int btn_h = px(36), btn_w = px(104);
    int by2 = ch - btn_h - px(20);
    int bx = (cw - (btn_w * 2 + px(16))) / 2;
    rounded_fill(hdc, bx, by2, btn_w, btn_h, col(C_ACCENT), 8);
    btn_text(hdc, g_fBold, bx, by2, btn_w, btn_h, L"确定", col(C_WHITE));
    rounded_fill(hdc, bx + btn_w + px(16), by2, btn_w, btn_h, col(C_ACCENT), 8);
    btn_text(hdc, g_fBold, bx + btn_w + px(16), by2, btn_w, btn_h, L"取消", col(C_WHITE));

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK pickproc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: p_draw(hwnd); return 0;
    case WM_NCCALCSIZE:
        // 去掉系统标题栏：客户区占满整个窗口
        if (wParam == TRUE) return 0;
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_NCHITTEST: {
        // 顶栏（系统按钮区除外）返回 HTCAPTION，支持拖动
        POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
        ScreenToClient(hwnd, &pt);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left;
        int bar_h = px(PICK_TITLE_H);
        if (pt.y >= 0 && pt.y < bar_h) {
            if (sb_hit(cw, bar_h, pt.x, pt.y) >= 0) return HTCLIENT;
            return HTCAPTION;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    case WM_MOUSEMOVE: {
        if (!gTracking) {
            TRACKMOUSEEVENT tme;
            memset(&tme, 0, sizeof(tme));
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hwnd;
            TrackMouseEvent(&tme);
            gTracking = TRUE;
        }
        int x = (short)LOWORD(lParam);
        int y = (short)HIWORD(lParam);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left;
        int bar_h = px(PICK_TITLE_H);
        int sb = (y >= 0 && y < bar_h) ? sb_hit(cw, bar_h, x, y) : -1;
        if (sb != gHoverBtn) { gHoverBtn = sb; inv_sys(hwnd, px(PICK_TITLE_H)); }
        return 0;
    }
    case WM_MOUSELEAVE: {
        gTracking = FALSE;
        if (gHoverBtn != -1) { gHoverBtn = -1; inv_sys(hwnd, px(PICK_TITLE_H)); }
        return 0;
    }
    case WM_LBUTTONDOWN: {
        int x = (short)LOWORD(lParam);
        int y = (short)HIWORD(lParam);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left, ch = rc.bottom - rc.top;
        int bar_h = px(PICK_TITLE_H);

        // 右上角系统按钮
        int sb = sb_hit(cw, bar_h, x, y);
        if (sb >= 0) {
            if (sb == 0) {
                ShowWindow(hwnd, SW_MINIMIZE);
            } else if (sb == 1) {
                if (IsZoomed(hwnd)) ShowWindow(hwnd, SW_RESTORE);
                else ShowWindow(hwnd, SW_MAXIMIZE);
                gHoverBtn = -1;
                inv_sys(hwnd, px(PICK_TITLE_H));
            } else {
                finish_picker(hwnd, FALSE);
            }
            return 0;
        }

        int y_top = px(PICK_TITLE_H) + px(10);
        int sb_h = px(30);
        int a_x = px(20), a_w = px(64);
        int n_x = a_x + a_w + px(12), n_w = px(76);
        if (y >= y_top && y < y_top + sb_h) {
            if (x >= a_x && x < a_x + a_w) { for (int i = 0; i < APP_COUNT; i++) gPickChecked[i] = TRUE; InvalidateRect(hwnd, NULL, TRUE); return 0; }
            if (x >= n_x && x < n_x + n_w) { for (int i = 0; i < APP_COUNT; i++) gPickChecked[i] = FALSE; InvalidateRect(hwnd, NULL, TRUE); return 0; }
        }

        int row_h = px(APP_ROW_H), pad = px(PICK_PAD), col_gap = px(PICK_COL_GAP);
        int colw = (cw - pad * 2 - col_gap) / 2;
        int y_rows = y_top + px(46);
        for (int i = 0; i < APP_COUNT; i++) {
            int rx = (i % 2 == 0) ? pad : pad + colw + col_gap;
            int ry = y_rows + (i / 2) * row_h;
            if (x >= rx && x < rx + colw && y >= ry && y < ry + row_h) {
                gPickChecked[i] = !gPickChecked[i];
                InvalidateRect(hwnd, NULL, TRUE);
                return 0;
            }
        }

        int btn_h = px(36), btn_w = px(104);
        int by = ch - btn_h - px(20);
        int bx = (cw - (btn_w * 2 + px(16))) / 2;
        if (x >= bx && x < bx + btn_w * 2 + px(16) && y >= by && y < by + btn_h) {
            finish_picker(hwnd, x < bx + btn_w);
            return 0;
        }
        return 0;
    }
    case WM_CLOSE: finish_picker(hwnd, FALSE); return 0;
    default: return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

// ---- 主窗口绘制 ----
static void draw(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    int cw = rc.right - rc.left, ch = rc.bottom - rc.top;
    BOOL busy = gExecBusy;

    fill(hdc, 0, 0, cw, ch, col(C_BG));

    // 顶栏 + 自绘标题 + 系统按钮
    int bar_h = px(56);
    fill(hdc, 0, 0, cw, bar_h, col(C_PANEL));
    RECT tr = { px(24), 0, cw - px(24) - px(SB_W) * 3, bar_h };
    HFONT of = (HFONT)SelectObject(hdc, g_fBold);
    SetTextColor(hdc, col(C_TEXT));
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, L"Windows/Office 便捷工具", -1, &tr, DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);
    draw_sysbtns(hdc, hwnd, cw, bar_h);

    // 两个按钮
    int btn_w = px(BTN_W), btn_h = px(BTN_H), gap = px(BTN_GAP);
    int total = btn_w * 2 + gap;
    int x0 = (cw - total) / 2;
    int y0 = bar_h + px(40);
    rounded_fill(hdc, x0, y0, btn_w, btn_h, col(C_ACCENT), 14);
    btn_text(hdc, g_fBold, x0, y0, btn_w, btn_h, L"一键安装 Office 365", col(C_WHITE));
    rounded_fill(hdc, x0 + btn_w + gap, y0, btn_w, btn_h, col(C_ACCENT), 14);
    btn_text(hdc, g_fBold, x0 + btn_w + gap, y0, btn_w, btn_h, L"Windows / Office 激活", col(C_WHITE));

    // 分隔线 + 状态
    int sy = y0 + btn_h + px(26);
    fill(hdc, px(20), sy, cw - px(40), 1, col(C_BORDER));
    wchar_t status[300];
    if (busy)
        wsprintfW(status, L"（%ls）%ls", gExecTag ? L"激活窗口中" : L"安装中", gStatus);
    else
        lstrcpyW(status, gStatus);
    status_text(hdc, g_fBase, px(24), sy + px(12), cw - px(48), ch - sy - px(12), status, col(C_SUB));

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK wndproc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: draw(hwnd); return 0;
    case WM_NCCALCSIZE:
        // 去掉系统标题栏与边框：客户区占满整个窗口
        if (wParam == TRUE) return 0;
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_NCHITTEST: {
        // 无边框窗口：顶栏（非系统按钮区）返回 HTCAPTION，支持拖动和双击最大化
        POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
        ScreenToClient(hwnd, &pt);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left;
        int ch = rc.bottom - rc.top;
        int bar_h = px(56);

        // 边框已被 NCCALCSIZE 吞掉，四边自绘缩放热区
        const int b = px(5);
        if (pt.y < b) {
            if (pt.x < b) return HTTOPLEFT;
            if (pt.x >= cw - b) return HTTOPRIGHT;
            return HTTOP;
        }
        if (pt.y >= ch - b) {
            if (pt.x < b) return HTBOTTOMLEFT;
            if (pt.x >= cw - b) return HTBOTTOMRIGHT;
            return HTBOTTOM;
        }
        if (pt.x < b) return HTLEFT;
        if (pt.x >= cw - b) return HTRIGHT;

        if (pt.y >= 0 && pt.y < bar_h) {
            if (sb_hit(cw, bar_h, pt.x, pt.y) >= 0) return HTCLIENT;
            return HTCAPTION;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    case WM_MOUSEMOVE: {
        if (!gTracking) {
            TRACKMOUSEEVENT tme;
            memset(&tme, 0, sizeof(tme));
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hwnd;
            TrackMouseEvent(&tme);
            gTracking = TRUE;
        }
        int x = (short)LOWORD(lParam);
        int y = (short)HIWORD(lParam);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left;
        int bar_h = px(56);
        int sb = (y >= 0 && y < bar_h) ? sb_hit(cw, bar_h, x, y) : -1;
        if (sb != gHoverBtn) { gHoverBtn = sb; inv_sys(hwnd, px(56)); }
        return 0;
    }
    case WM_MOUSELEAVE: {
        gTracking = FALSE;
        if (gHoverBtn != -1) { gHoverBtn = -1; inv_sys(hwnd, px(56)); }
        return 0;
    }
    case WM_LBUTTONDOWN: {
        int x = (short)LOWORD(lParam);
        int y = (short)HIWORD(lParam);
        RECT rc; GetClientRect(hwnd, &rc);
        int cw = rc.right - rc.left;
        int bar_h = px(56);
        // 顶栏右上角系统按钮：最小化 / 最大化(还原) / 关闭，始终可用
        int sb = sb_hit(cw, bar_h, x, y);
        if (sb >= 0) {
            if (sb == 0) {
                ShowWindow(hwnd, SW_MINIMIZE);
            } else if (sb == 1) {
                if (IsZoomed(hwnd)) ShowWindow(hwnd, SW_RESTORE);
                else ShowWindow(hwnd, SW_MAXIMIZE);
                gHoverBtn = -1;
                inv_sys(hwnd, px(56));
            } else {
                SendMessageW(hwnd, WM_CLOSE, 0, 0);
            }
            return 0;
        }
        if (gExecBusy || gPickOpen) return 0;
        int btn_w = px(BTN_W), btn_h = px(BTN_H), gap = px(BTN_GAP);
        int total = btn_w * 2 + gap;
        int x0 = (cw - total) / 2;
        int y0 = bar_h + px(40);

        if (x >= x0 && x < x0 + btn_w && y >= y0 && y < y0 + btn_h) {
            gPickMain = hwnd;
            for (int i = 0; i < APP_COUNT; i++) gPickChecked[i] = TRUE;
            gPickConfirmed = FALSE;
            gPickOpen = TRUE;
            // 相对主界面居中弹出
            RECT pr, cr;
            GetWindowRect(hwnd, &pr);
            GetClientRect(gPickHwnd, &cr);
            int pxc = pr.left + (pr.right - pr.left) / 2;
            int pyc = pr.top + (pr.bottom - pr.top) / 2;
            SetWindowPos(gPickHwnd, NULL,
                pxc - (cr.right - cr.left) / 2, pyc - (cr.bottom - cr.top) / 2,
                0, 0, SWP_NOSIZE | SWP_NOZORDER);
            ShowWindow(gPickHwnd, SW_SHOW);
            SetForegroundWindow(gPickHwnd);
        } else if (x >= x0 + btn_w + gap && x < x0 + btn_w + gap + btn_w && y >= y0 && y < y0 + btn_h) {
            InterlockedExchange(&gExecBusy, 1);
            gExecTag = 1;
            lstrcpyW(gStatus, L"正在打开 PowerShell 激活菜单（Windows 与 Office 均可）…");
            CloseHandle(CreateThread(NULL, 0, ThreadActivate, NULL, 0, NULL));
        }
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_PICK_DONE: {
        if (gPickConfirmed) {
            // 未勾选 = 排除安装
            gThreadExCount = 0;
            for (int i = 0; i < APP_COUNT; i++)
                if (!gPickChecked[i]) gThreadEx[gThreadExCount++] = gApps[i];
            if (gThreadExCount == 0)
                lstrcpyW(gStatus, L"正在下载并安装 Office 365 全套…（setup 界面会自动打开）");
            else {
                wchar_t skip[160] = L"";
                for (int i = 0; i < gThreadExCount; i++) {
                    if (i) lstrcatW(skip, L"、");
                    wsprintfW(skip + lstrlenW(skip), L"%ls", gThreadEx[i]);
                }
                wsprintfW(gStatus, L"正在下载并安装 Office 365（跳过：%ls）…", skip);
            }
            InterlockedExchange(&gExecBusy, 1);
            gExecTag = 0;
            CloseHandle(CreateThread(NULL, 0, ThreadOffice, NULL, 0, NULL));
        }
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_DONE: {
        lstrcpyW(gStatus, gResult); // 后台线程已写入结果
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_DESTROY: PostQuitMessage(0); return 0;
    default: return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance; (void)lpCmdLine; (void)nCmdShow;
    // EXE 目录（供子进程定位相对路径）
    GetModuleFileNameW(NULL, gExeDir, MAX_PATH);
    wchar_t *slash = wcsrchr(gExeDir, L'\\');
    if (slash) *slash = 0;

    g_dpi = (float)GetDpiForSystem() / 96.0f;
    g_fBase = make_font(19, FALSE);
    g_fBold = make_font(22, TRUE);

    WNDCLASSW wc;
    memset(&wc, 0, sizeof(wc));
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = wndproc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"MoActivationMain";
    RegisterClassW(&wc);

    memset(&wc, 0, sizeof(wc));
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = pickproc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"MoActivationPick";
    RegisterClassW(&wc);

    // WS_POPUP 完全去掉标题栏与边框；拖动/缩放由 WM_NCHITTEST 自绘接管
    HWND hwnd = CreateWindowExW(0, L"MoActivationMain", L"",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, px(WIN_W), px(WIN_H), NULL, NULL, hInstance, NULL);
    if (!hwnd) return 1;

    // 启动时居中到屏幕工作区中心
    RECT wr;
    GetWindowRect(hwnd, &wr);
    RECT wk;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &wk, 0);
    SetWindowPos(hwnd, NULL,
        wk.left + ((wk.right - wk.left) - (wr.right - wr.left)) / 2,
        wk.top + ((wk.bottom - wk.top) - (wr.bottom - wr.top)) / 2,
        0, 0, SWP_NOSIZE | SWP_NOZORDER);

    gPickHwnd = CreateWindowExW(0, L"MoActivationPick", L"选择组件", WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, px(PICK_W), px(PICK_H), hwnd, NULL, hInstance, NULL);
    if (!gPickHwnd) return 1;

    // 圆角（全自绘白底 + 无标题栏，Mica 无意义，不启用）
    DWORD corner = DWMWCP_ROUND;
    DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, &corner, sizeof(corner));
    DwmSetWindowAttribute(gPickHwnd, DWMWA_WINDOW_CORNER_PREFERENCE, &corner, sizeof(corner));

    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (g_fBase) DeleteObject(g_fBase);
    if (g_fBold) DeleteObject(g_fBold);
    return (int)msg.wParam;
}