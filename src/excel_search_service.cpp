// excel_search_service.cpp
// 一个 Windows 服务程序，监听剪贴板内容，搜索 Excel 文件中是否包含该字符串
// 支持中文路径、UTF-8 输出，适配 CMD 环境

#include <windows.h>
#include <tchar.h>
#include <io.h>
#include <fcntl.h>
#include <iostream>
#include <fstream>
#include <thread>
#include <atomic>
#include <openxlsx.hpp> // OpenXLSX 头文件
#include <optional>
#include <cctype>
#include <cstring>
#include <codecvt>
using namespace OpenXLSX;
	static UINT code_page = GetConsoleOutputCP(); //显示字符集
	static UINT acp = GetACP(); //输入字符集
#pragma comment(lib, "Advapi32.lib")
std::string CharsetConvert(const char* utf8_str,UINT acp=936,UINT consolecp=CP_UTF8) {
	
	//字符集转换 acp 为激活代码页（输入参数字符集）,consolecp 为控制台字符集（输出字符集)
    // 1. 输入参数 转 unicode
    if(acp==consolecp){return std::string(utf8_str);}
    int wlen = MultiByteToWideChar(acp, 0, utf8_str, -1, nullptr, 0);
    if (wlen <= 0) return "";
    std::vector<wchar_t> wbuf(wlen);
    MultiByteToWideChar(acp, 0, utf8_str, -1, wbuf.data(), wlen);
    //unicode 转 输出字符
    int len = WideCharToMultiByte(consolecp, 0, wbuf.data(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::vector<char> buf(len);
    WideCharToMultiByte(consolecp, 0, wbuf.data(), -1, buf.data(), len, nullptr, nullptr);
    
    return std::string(buf.data());
   
}
// ==== UTF-8 控制台设置 ====
void setup_utf8_console() {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    _setmode(_fileno(stdout), _O_U8TEXT);
}

struct Match {
    std::string sheetName;
    std::string cellAddress;
		std::string val;
    Match(const std::string& sheet, const std::string& addr,const std::string&_val)
        : sheetName(sheet), cellAddress(addr),val(_val) {}
};

// 将列字母（如"A", "B"）转换为列号（1, 2,...）
std::optional<uint16_t> column_letter_to_index(const std::string& colLetter) {
    if (colLetter.empty() || colLetter.size() > 2) return std::nullopt;

    uint16_t col = 0;
    for (char c : colLetter) {
        if (!std::isalpha(c)) return std::nullopt;
        col = col * 26 + (std::toupper(c) - 'A' + 1);
    }
    return col;
}

uint32_t detect_max_row(XLWorksheet& sheet, uint16_t col = 1) {
    uint32_t maxRow = 1;
    for (uint32_t row = 1; row <= 2000000; ++row) {
        auto cell = sheet.cell(XLCellReference(row, col));
        if (cell.value().type() != XLValueType::Empty) {
            maxRow = row;
        }
        else 
        {
	        	break;
        }
    }
    return maxRow;
}

std::vector<Match> find_string_in_sheet(XLWorksheet& sheet, const std::string& target, uint16_t col) {
    std::vector<Match> matches;
    uint32_t maxRow = detect_max_row(sheet, col);

    for (uint32_t row = 1; row <= maxRow; ++row) {
        auto cell = sheet.cell(XLCellReference(row, col));
        if (cell.value().type() == XLValueType::Empty) continue;

        std::string val = cell.value().get<std::string>();
        if (val.find(target) != std::string::npos) {
            matches.emplace_back(sheet.name(), cell.cellReference().address(),val);
        }
    }

    return matches;
}

std::vector<Match> search_excel_all_sheets(const std::string& filepath, const std::string& keyword,const std::string& column) {
    std::vector<Match> results;
    OpenXLSX::XLDocument doc;
    doc.open(filepath);

    auto colOpt = column_letter_to_index(column);
    if (!colOpt) {
        std::cerr << CharsetConvert("无效的列字母: ",CP_UTF8,code_page) << column << "\n";
        return results;
    }
    uint16_t colIndex = colOpt.value();
 
        auto wb = doc.workbook();
        auto sheets = wb.worksheetNames();

        

        for (const auto& sheetName : sheets) {
            auto sheet = wb.worksheet(sheetName);
            auto matches = find_string_in_sheet(sheet, keyword, colIndex);
            results.insert(results.end(), matches.begin(), matches.end());
        }

        doc.close();
     return results;
}

// ==== 剪贴板内容获取 ====
std::wstring get_clipboard_text() {
    std::wstring result;
    if (!OpenClipboard(nullptr)) return result;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        wchar_t* text = static_cast<wchar_t*>(GlobalLock(hData));
        if (text) result = text;
        GlobalUnlock(hData);
    }
    CloseClipboard();
    return result;
}

// ==== 日志输出 ====
void log_matches(const std::vector<Match>& matches, const std::wstring& keyword) {
    std::wofstream log("search_log.txt", std::ios::app);
    log.imbue(std::locale(log.getloc(), new std::codecvt_utf8<wchar_t>()));
    log << L"=== 搜索关键词：" << keyword << L" ===\n";
    for (const auto& match : matches) {
        log << L"[" << match.sheetName << L"] " << match.cellAddress << L" -> " << match.value << L"\n";
    }
    log << L"==============================\n";
    log.close();
}

// ==== 服务相关 ====
std::atomic<bool> g_Running{ true };

void run_service_loop(const std::wstring& excel_path) {
    std::wstring last;
    while (g_Running) {
        std::wstring clip = get_clipboard_text();
        if (!clip.empty() && clip != last) {
            last = clip;
            auto matches = search_excel_all_sheets(excel_path, clip);
            log_matches(matches, clip);
        }
        Sleep(2000);
    }
}

SERVICE_STATUS g_ServiceStatus;
SERVICE_STATUS_HANDLE g_StatusHandle = nullptr;
std::thread g_WorkerThread;

void WINAPI ServiceCtrlHandler(DWORD CtrlCode) {
    if (CtrlCode == SERVICE_CONTROL_STOP) {
        g_ServiceStatus.dwCurrentState = SERVICE_STOP_PENDING;
        SetServiceStatus(g_StatusHandle, &g_ServiceStatus);
        g_Running = false;
        if (g_WorkerThread.joinable()) g_WorkerThread.join();
        g_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
        SetServiceStatus(g_StatusHandle, &g_ServiceStatus);
    }
}

void WINAPI ServiceMain(DWORD, LPTSTR*) {
    g_StatusHandle = RegisterServiceCtrlHandler(L"ExcelSearchService", ServiceCtrlHandler);
    g_ServiceStatus = { SERVICE_WIN32_OWN_PROCESS, SERVICE_START_PENDING, 0, 0, 0, 0, 0 };
    SetServiceStatus(g_StatusHandle, &g_ServiceStatus);

    g_ServiceStatus.dwCurrentState = SERVICE_RUNNING;
    SetServiceStatus(g_StatusHandle, &g_ServiceStatus);

    setup_utf8_console();
    std::wstring excelFile = L"20250304.xlsm"; // TODO: 替换为你的实际路径
    g_WorkerThread = std::thread(run_service_loop, excelFile);
}

int wmain(int argc, wchar_t* argv[]) {
    SERVICE_TABLE_ENTRY ServiceTable[] = {
        { (LPWSTR)L"ExcelSearchService", ServiceMain },
        { nullptr, nullptr }
    };

    if (!StartServiceCtrlDispatcher(ServiceTable)) {
        std::wcerr << L"无法启动服务控制调度器。\n";
    }
    return 0;
}
