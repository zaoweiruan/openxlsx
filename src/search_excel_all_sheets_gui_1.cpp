// excel_search_gui.cpp
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <openxlsx.hpp>
#pragma comment(lib, "Comctl32.lib")

using namespace OpenXLSX;

// 控件 ID
#define IDC_BTN_SEARCH 101
#define IDC_BTN_CLEAR 102
#define IDC_LISTVIEW 103

// UTF-8 转 WideChar
std::wstring UTF8ToWString(const std::string &str) {
    int wlen = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    std::wstring wstr(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &wstr[0], wlen);
    wstr.pop_back(); // remove null
    return wstr;
}

// WideChar 转 UTF-8
std::string WStringToUTF8(const std::wstring &wstr) {
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    std::string str(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], len, NULL, NULL);
    str.pop_back(); // remove null
    return str;
}

// 搜索匹配
struct Match {
    std::wstring sheet;
    std::wstring address;
    std::wstring value;
};

// Excel 搜索函数
std::vector<Match> searchExcel(const std::wstring &path, const std::wstring &keyword, const std::wstring &colLetter) {
    std::vector<Match> results;
    try {
        XLDocument doc;
        doc.open(WStringToUTF8(path));
        auto wb = doc.workbook();
        auto sheets = wb.worksheetNames();

        // 转列字母 -> 索引
        int col = 0;
        for (wchar_t c : colLetter) col = col * 26 + (towupper(c) - L'A' + 1);

        for (auto &s : sheets) {
            auto ws = wb.worksheet(s);
            auto maxRow = 100000; // 最多检测1000行
            for (uint32_t row = 1; row <= maxRow; ++row) {
                auto cell = ws.cell(XLCellReference(row, col));
                if (cell.value().type() == XLValueType::Empty) break;
                auto val = cell.value().getString();
                std::wstring wval = UTF8ToWString(val);
                if (wval.find(keyword) != std::wstring::npos) {
                    results.push_back({ UTF8ToWString(s), UTF8ToWString(cell.cellReference().address()), wval });
                }
            }
        }
        doc.close();
    } catch (...) {
        MessageBox(NULL, L"读取 Excel 文件出错。", L"错误", MB_ICONERROR);
    }
    return results;
}

// 填充 ListView
void FillListView(HWND hList, const std::vector<Match> &data) {
    ListView_DeleteAllItems(hList);
    for (size_t i = 0; i < data.size(); ++i) {
        LVITEM item = { 0 };
        item.mask = LVIF_TEXT;
        item.iItem = (int)i;
        item.pszText = const_cast<LPWSTR>(data[i].sheet.c_str());
        ListView_InsertItem(hList, &item);
        ListView_SetItemText(hList, (int)i, 1, const_cast<LPWSTR>(data[i].address.c_str()));
        ListView_SetItemText(hList, (int)i, 2, const_cast<LPWSTR>(data[i].value.c_str()));
    }
}

// 主窗口过程
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hList, hSearchBtn, hClearBtn;

    switch (msg) {
    case WM_CREATE: {
        InitCommonControls();

        hSearchBtn = CreateWindow(L"BUTTON", L"搜索剪贴板",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 20, 20, 120, 30, hwnd, (HMENU)IDC_BTN_SEARCH, NULL, NULL);

        hClearBtn = CreateWindow(L"BUTTON", L"清空", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            160, 20, 80, 30, hwnd, (HMENU)IDC_BTN_CLEAR, NULL, NULL);

        hList = CreateWindow(WC_LISTVIEW, NULL,
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,
            20, 70, 600, 400, hwnd, (HMENU)IDC_LISTVIEW, NULL, NULL);

        ListView_SetExtendedListViewStyle(hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

        LVCOLUMN col = { 0 };
        col.mask = LVCF_TEXT | LVCF_WIDTH;
        col.pszText = (LPWSTR)L"Sheet";
        col.cx = 120;
        ListView_InsertColumn(hList, 0, &col);
        col.pszText = (LPWSTR)L"单元格";
        col.cx = 100;
        ListView_InsertColumn(hList, 1, &col);
        col.pszText = (LPWSTR)L"内容";
        col.cx = 380;
        ListView_InsertColumn(hList, 2, &col);

        break;
    }
    case WM_COMMAND: {
        if (LOWORD(wParam) == IDC_BTN_SEARCH) {
            if (!OpenClipboard(NULL)) break;
            HANDLE hData = GetClipboardData(CF_UNICODETEXT);
            if (hData) {
                wchar_t* text = (wchar_t*)GlobalLock(hData);
                std::wstring keyword = text;
                GlobalUnlock(hData);
                CloseClipboard();

                std::wstring file = L"20250304.xlsm"; // Excel 文件路径
                std::wstring column = L"A";           // 搜索列
                auto matches = searchExcel(file, keyword, column);
                FillListView(hList, matches);
            }
        }
        else if (LOWORD(wParam) == IDC_BTN_CLEAR) {
            ListView_DeleteAllItems(hList);
        }
        break;
    }
    case WM_DESTROY: PostQuitMessage(0); break;
    default: return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// WinMain - GUI 程序入口
int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nCmdShow) {
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"ExcelSearchWindow";
    RegisterClass(&wc);

    HWND hwnd = CreateWindow(L"ExcelSearchWindow", L"Excel 中文搜索工具",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 660, 550, NULL, NULL, hInst, NULL);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg); DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}
