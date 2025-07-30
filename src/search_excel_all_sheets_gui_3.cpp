// excel_search_gui.cpp
// 一个基于 Win32 API 的 Excel 搜索窗口程序，支持剪贴板自动查询和命令行参数

#include <windows.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <map>
#include <thread>
#include <chrono>
#include <openxlsx.hpp>
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

using namespace OpenXLSX;

LPCWSTR g_szClassName = L"ExcelSearchWndClass";
HWND hwndPath, hwndKeyword, hwndList;
std::wstring g_excelPath;
std::wstring g_lastClipboard;
UINT_PTR TIMER_ID = 1;

struct Match {
    std::wstring sheetName;
    std::wstring cellAddr;
    std::wstring value;
};
std::string WideCharConvertMultiByte(const std::wstring &wstr,UINT acp=936) {
	
	//字符集转换 acp 为激活代码页（输入参数字符集）,consolecp 为控制台字符集（输出字符集)
		
    int wlen = WideCharToMultiByte(acp, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (wlen <= 0) return "";
    std::vector<char> buf(wlen);
    WideCharToMultiByte(acp, 0, wstr.c_str(), -1, buf.data(), wlen, nullptr, nullptr);
  
    return std::string(buf.data());
   
}
std::wstring MultiByteConvertWideChar(const std::string &str,UINT acp=936) {
	
	//字符集转换 acp 为激活代码页（输入参数字符集）,consolecp 为控制台字符集（输出字符集)
		
    int wlen = MultiByteToWideChar(acp, 0, str.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return L"";
    std::vector<wchar_t> buf(wlen);
    MultiByteToWideChar(acp, 0, str.c_str(), -1, buf.data(), wlen);
  
    return std::wstring(buf.data());
   
}

std::vector<Match> searchExcel(const std::wstring& path, const std::wstring& keyword) {
    std::vector<Match> results;
    if (!PathFileExists(path.c_str())) return results;
    try {
        XLDocument doc;
//        doc.open(path);
        doc.open(WideCharConvertMultiByte(path,GetACP()));
        auto wb = doc.workbook();
        auto sheets = wb.worksheetNames();
        for (auto& s : sheets) {
            auto sheet = wb.worksheet(s);
            for (uint32_t row = 1; row < 10000; ++row) {
                auto cell = sheet.cell(XLCellReference(row, 1));
                if (cell.value().type() == XLValueType::Empty) break;
                std::string str = cell.value().get<std::string>();
                std::wstring wstr(str.begin(), str.end());
                if (wstr.find(keyword) != std::wstring::npos) {
                    results.push_back({ std::wstring(s.begin(), s.end()), std::to_wstring(row), wstr });
                }
            }
        }
        doc.close();
    } catch (...) {}
    return results;
}

void updateResultList(HWND list, const std::vector<Match>& matches) {
    SendMessage(list, LB_RESETCONTENT, 0, 0);
    for (const auto& m : matches) {
        std::wstring line = L"[" + m.sheetName + L"] " + m.cellAddr + L" -> " + m.value;
        SendMessage(list, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    }
}

void openExcelFile(const std::wstring& filePath) {
    ShellExecute(NULL, L"open", filePath.c_str(), NULL, NULL, SW_SHOWNORMAL);
}

std::wstring getClipboardText() {
    std::wstring result;
    if (!OpenClipboard(NULL)) return result;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        result = (wchar_t*)GlobalLock(hData);
        GlobalUnlock(hData);
    }
    CloseClipboard();
    return result;
}

void handleSearch(HWND hwnd) {
    wchar_t path[MAX_PATH], keyword[256];
    GetWindowText(hwndPath, path, MAX_PATH);
    GetWindowText(hwndKeyword, keyword, 256);
    auto results = searchExcel(path, keyword);
    updateResultList(hwndList, results);
}

void handleBrowse(HWND hwnd) {
    OPENFILENAME ofn = { sizeof(ofn) };
    wchar_t file[MAX_PATH] = L"";
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"Excel Files\0*.xlsx;*.xlsm\0All Files\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST;
    if (GetOpenFileName(&ofn)) {
        SetWindowText(hwndPath, file);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        CreateWindow(L"STATIC", L"Excel 文件:", WS_VISIBLE | WS_CHILD, 10, 10, 80, 20, hwnd, 0, NULL, NULL);
        hwndPath = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            90, 10, 300, 20, hwnd, 0, NULL, NULL);
        CreateWindow(L"BUTTON", L"浏览", WS_VISIBLE | WS_CHILD, 400, 10, 60, 20, hwnd, (HMENU)1001, NULL, NULL);

        CreateWindow(L"STATIC", L"关键字:", WS_VISIBLE | WS_CHILD, 10, 40, 80, 20, hwnd, 0, NULL, NULL);
        hwndKeyword = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            90, 40, 200, 20, hwnd, 0, NULL, NULL);

        CreateWindow(L"BUTTON", L"搜索", WS_VISIBLE | WS_CHILD, 300, 40, 60, 20, hwnd, (HMENU)1002, NULL, NULL);
        CreateWindow(L"BUTTON", L"清空", WS_VISIBLE | WS_CHILD, 370, 40, 60, 20, hwnd, (HMENU)1003, NULL, NULL);

        hwndList = CreateWindow(L"LISTBOX", NULL, WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL | LBS_NOTIFY,
            10, 70, 460, 280, hwnd, (HMENU)1004, NULL, NULL);

        SetTimer(hwnd, TIMER_ID, 2000, NULL);
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case 1001: handleBrowse(hwnd); break;
        case 1002: handleSearch(hwnd); break;
        case 1003: SendMessage(hwndList, LB_RESETCONTENT, 0, 0); break;
        case 1004:
            if (HIWORD(wParam) == LBN_DBLCLK) openExcelFile(g_excelPath);
            break;
        }
        return 0;

    case WM_TIMER:
        if (wParam == TIMER_ID) {
            std::wstring text = getClipboardText();
            if (!text.empty() && text != g_lastClipboard) {
                g_lastClipboard = text;
                SetWindowText(hwndKeyword, text.c_str());
                handleSearch(hwnd);
            }
        }
        return 0;

    case WM_DESTROY:
        KillTimer(hwnd, TIMER_ID);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR lpCmdLine, int nCmdShow) {
    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = g_szClassName;
    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(0, g_szClassName, L"Excel 搜索工具",
        WS_OVERLAPPEDWINDOW ^ WS_MAXIMIZEBOX ^ WS_THICKFRAME,
        CW_USEDEFAULT, CW_USEDEFAULT, 500, 400,
        nullptr, nullptr, hInstance, nullptr);

    std::map<std::wstring, std::wstring> args;
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(lpCmdLine, &argc);
    for (int i = 0; i < argc; ++i) {
        std::wstring arg = argv[i];
        auto pos = arg.find(L'=');
        if (pos != std::wstring::npos) {
            args[arg.substr(0, pos)] = arg.substr(pos + 1);
        }
    }
    if (args.count(L"filename")) {
        g_excelPath = args[L"filename"];
        SetWindowText(hwndPath, g_excelPath.c_str());
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}
