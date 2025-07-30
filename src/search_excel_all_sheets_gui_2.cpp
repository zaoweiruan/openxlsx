// excel_gui_clipboard_search.cpp
// GUI 程序：支持剪贴板自动监视 Excel，搜索并展示结果，支持文件选择、点击打开、清空等功能

#include <windows.h>
#include <commctrl.h>
#include <shobjidl.h> 
#include <shellapi.h>
#include <vector>
#include <string>
#include <fstream>
#include <algorithm>
#include <openxlsx.hpp>
#include <codecvt>
#include <locale>
#include <optional>
using namespace OpenXLSX;

HINSTANCE hInst;
HWND hList, hEditSearch, hBtnSelect, hBtnSearch, hBtnClear, hChkAuto;
std::wstring g_excelFile;
bool g_autoSearch = true;
std::wstring g_lastClipboard;

struct Match {
    std::wstring sheet;
    std::wstring address;
    std::wstring value;
   	Match(const std::wstring& sheet, const std::wstring& addr,const std::wstring&_val)
        : sheet(sheet), address(addr),value(_val) {}
};

std::vector<Match> g_currentMatches;

std::wstring GetClipboardText() {
    std::wstring result;
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return result;
    if (!OpenClipboard(NULL)) return result;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        wchar_t* pText = static_cast<wchar_t*>(GlobalLock(hData));
        if (pText) {
            result = pText;
            GlobalUnlock(hData);
        }
    }
    CloseClipboard();
    return result;
}

std::optional<uint16_t> column_letter_to_index(const std::wstring& colLetter) {
    if (colLetter.empty() || colLetter.size() > 2) return std::nullopt;
    uint16_t col = 0;
    for (wchar_t c : colLetter) {
        if (!iswalpha(c)) return std::nullopt;
        col = col * 26 + (towupper(c) - L'A' + 1);
    }
    return col;
}
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
std::vector<Match> searchExcel(const std::wstring& filePath, const std::wstring& keyword, const std::wstring& col = L"A") {
    std::vector<Match> results;
    try {
        XLDocument doc;
        doc.open(WideCharConvertMultiByte(filePath,GetACP()));
        auto wb = doc.workbook();
        auto colIndexOpt = column_letter_to_index(col);
        if (!colIndexOpt) return results;
        uint16_t colIndex = colIndexOpt.value();

        for (auto sheetName : wb.worksheetNames()) {
            auto sheet = wb.worksheet(sheetName);
            for (uint32_t row = 1; row <= 100000; ++row) {
                auto cell = sheet.cell(XLCellReference(row, colIndex));
                if (cell.value().type() == XLValueType::Empty) break;
                std::wstring sheetname=	MultiByteConvertWideChar(sheet.name(),CP_UTF8);
                std::wstring address=MultiByteConvertWideChar(cell.cellReference().address(),CP_UTF8);
                std::wstring val = MultiByteConvertWideChar(cell.value().get<std::string>(),CP_UTF8);
                if (val.find(keyword) != std::wstring::npos) {
                    results.push_back({ sheetname,address , val });
                }
            }
        }
        doc.close();
    } catch (...) {}
    return results;
}

void FillListView(HWND hList, const std::vector<Match>& matches) {
    ListView_DeleteAllItems(hList);
    for (int i = 0; i < matches.size(); ++i) {
        const auto& m = matches[i];
        LVITEM lvi = { 0 };
        lvi.mask = LVIF_TEXT;
        lvi.iItem = i;
        lvi.pszText = const_cast<LPWSTR>(m.sheet.c_str());
        ListView_InsertItem(hList, &lvi);
        ListView_SetItemText(hList, i, 1, const_cast<LPWSTR>(m.address.c_str()));
        ListView_SetItemText(hList, i, 2, const_cast<LPWSTR>(m.value.c_str()));
    }
}

void DoSearch() {
    wchar_t keyword[256] = {};
    GetWindowText(hEditSearch, keyword, 255);
    auto matches = searchExcel(g_excelFile, keyword);
    g_currentMatches = matches;
    FillListView(hList, matches);
}

void SelectExcelFile(HWND hwnd) {
    IFileOpenDialog* pFileOpen;
    if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pFileOpen)))) {
        COMDLG_FILTERSPEC rgSpec[] = {
            { L"Excel Files", L"*.xlsx;*.xlsm" },
            { L"All Files", L"*.*" }
        };
        pFileOpen->SetFileTypes(ARRAYSIZE(rgSpec), rgSpec);
        if (SUCCEEDED(pFileOpen->Show(hwnd))) {
            IShellItem* pItem;
            if (SUCCEEDED(pFileOpen->GetResult(&pItem))) {
                PWSTR pszFilePath;
                if (SUCCEEDED(pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath))) {
                    g_excelFile = pszFilePath;
                    SetWindowText(hBtnSelect, pszFilePath);
                    CoTaskMemFree(pszFilePath);
                }
                pItem->Release();
            }
        }
        pFileOpen->Release();
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        InitCommonControls();
        hEditSearch = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            10, 10, 300, 25, hwnd, NULL, hInst, NULL);

        hBtnSearch = CreateWindow(L"BUTTON", L"搜索", WS_CHILD | WS_VISIBLE,
            320, 10, 80, 25, hwnd, (HMENU)1, hInst, NULL);

        hBtnClear = CreateWindow(L"BUTTON", L"清空", WS_CHILD | WS_VISIBLE,
            410, 10, 80, 25, hwnd, (HMENU)2, hInst, NULL);

        hBtnSelect = CreateWindow(L"BUTTON", L"选择文件", WS_CHILD | WS_VISIBLE,
            500, 10, 160, 25, hwnd, (HMENU)3, hInst, NULL);

        hChkAuto = CreateWindow(L"BUTTON", L"自动剪贴板搜索", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            670, 10, 160, 25, hwnd, (HMENU)4, hInst, NULL);
        SendMessage(hChkAuto, BM_SETCHECK, BST_CHECKED, 0);

        hList = CreateWindow(WC_LISTVIEW, NULL,
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,
            10, 50, 800, 500,
            hwnd, NULL, hInst, NULL);

        ListView_SetExtendedListViewStyle(hList, LVS_EX_FULLROWSELECT);
        LVCOLUMN lvc = { 0 };
        lvc.mask = LVCF_TEXT | LVCF_WIDTH;
        lvc.cx = 150;
        lvc.pszText = L"工作表";
        ListView_InsertColumn(hList, 0, &lvc);
        lvc.pszText = L"单元格";
        ListView_InsertColumn(hList, 1, &lvc);
        lvc.cx = 400;
        lvc.pszText = L"值";
        ListView_InsertColumn(hList, 2, &lvc);

        SetTimer(hwnd, 1, 2000, NULL);
    }
    break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case 1: DoSearch(); break;
        case 2: SetWindowText(hEditSearch, L""); ListView_DeleteAllItems(hList); break;
        case 3: SelectExcelFile(hwnd); break;
        case 4: g_autoSearch = (IsDlgButtonChecked(hwnd, 4) == BST_CHECKED); break;
        }
        break;
    case WM_TIMER:
        if (g_autoSearch && !g_excelFile.empty()) {
            std::wstring clip = GetClipboardText();
            if (!clip.empty() && clip != g_lastClipboard) {
                g_lastClipboard = clip;
                SetWindowText(hEditSearch, clip.c_str());
                auto matches = searchExcel(g_excelFile, clip);
                g_currentMatches = matches;
                FillListView(hList, matches);
            }
        }
        break;
    case WM_NOTIFY:
        if (((LPNMHDR)lParam)->hwndFrom == hList && ((LPNMHDR)lParam)->code == NM_DBLCLK) {
            int iSel = ListView_GetNextItem(hList, -1, LVNI_SELECTED);
            if (iSel >= 0 && iSel < g_currentMatches.size()) {
                std::wstring path = g_excelFile;
                ShellExecute(NULL, L"open", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
            }
        }
        break;
    case WM_DESTROY:
        KillTimer(hwnd, 1);
        PostQuitMessage(0);
        break;
    default: return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    hInst = hInstance;
    const wchar_t CLASS_NAME[] = L"ExcelSearchWin";
    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClass(&wc);

    HWND hwnd = CreateWindow(CLASS_NAME, L"Excel 关键字搜索", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 850, 620, NULL, NULL, hInstance, NULL);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}
