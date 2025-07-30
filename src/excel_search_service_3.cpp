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
#include <cstdlib>  // 包含 _wtoi（Windows 专用）
#include <map>
#include <cwchar>  // 包含 _wcsicmp（Windows 专用）
using namespace OpenXLSX;
// 自定义带参数的断言宏
#define ASSERT_MSG(condition, ...) \
    do { \
        if (!(condition)) { \
            fprintf(stderr, "Assertion failed: %s, file %s, line %d: ", \
                    #condition, __FILE__, __LINE__); \
            fprintf(stderr, __VA_ARGS__); \
            fprintf(stderr, "\n"); \
            abort(); \
        } \
    } while (0)

 
//#pragma comment(lib, "Advapi32.lib")
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
std::map<std::wstring, std::wstring> parse_args(int argc, wchar_t* argv[]) {
    std::map<std::wstring, std::wstring> args;
    for (int i = 1; i < argc; ++i) {
        std::wstring arg = argv[i];
        auto pos = arg.find(L'=');
        if (pos != std::wstring::npos) {
            std::wstring key = arg.substr(0, pos);
            std::wstring value = arg.substr(pos + 1);
            args[key] = value;
        }
    }
    return args;
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
		std::string value;
    Match(const std::string& sheet, const std::string& addr,const std::string&_val)
        : sheetName(sheet), cellAddress(addr),value(_val) {}
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

//        std::string val = cell.value().get<std::string>();
        std::string val = cell.value().getString();
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
        std::cerr << CharsetConvert("无效的列字母: ",CP_UTF8,GetConsoleOutputCP()) << column << "\n";
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
std::string get_clipboard_text() {
    std::string result;
    if (!OpenClipboard(nullptr)) return result;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        wchar_t* text = static_cast<wchar_t*>(GlobalLock(hData));
        if (text) result = WideCharConvertMultiByte(std::wstring(text),CP_UTF8);
        GlobalUnlock(hData);
    }
    CloseClipboard();
    return result;
}

// ==== 日志输出 ====
void log_console(const std::vector<Match>& matches,std::string filename, const std::string& keyword,bool detail=FALSE,std::string column="A") {
	std::wostream *log = &std::wcout;  // 默认输出到控制台
      
    	if (detail)
      {for (const auto& match : matches) {
      	*log << L"[" <<MultiByteConvertWideChar(match.sheetName,CP_UTF8) << L"] "\
      	<< MultiByteConvertWideChar(match.cellAddress,CP_UTF8)<< L" -> " \
      	<< MultiByteConvertWideChar(match.value,CP_UTF8) << "\n";}
      }
     *log << L"=== 搜索 " <<MultiByteConvertWideChar(column,CP_UTF8)<<L" 列 \n"<<L" 关键词：" << MultiByteConvertWideChar(keyword,CP_UTF8) << L" 匹配结果,共 ："<<matches.size()<<L" 条记录"<<" \n";		

    *log << L"==============================\n";
  
}
// ==== 日志输出 ====
void log_file(const std::vector<Match>& matches,std::string filename, const std::string& keyword,bool detail=FALSE,std::string colume="A") {
    std::ofstream log("search_log.txt", std::ios::app);
//    log.imbue(std::locale(log.getloc(), new std::codecvt_utf8<wchar_t>()));
 
    	if (detail)
      {for (const auto& match : matches) {
      	log << "[" << match.sheetName << "] " << match.cellAddress << " -> " << match.value << "\n";}
      }
     log <<  "=== 搜索关键词：" << keyword << " 匹配结果,共 ："<<matches.size()<<" 条记录 ===\n";		

    log << "==============================\n";
    log.close();
    }
// ==== 服务相关 ====
std::atomic<bool> g_Running{ true };

void run_service_loop(const std::wstring& excel_path,bool detail=FALSE,std::wstring column=L"A",bool tofile=FALSE) {
    std::string last;
    std::string excel_file=WideCharConvertMultiByte(excel_path,GetACP());
    std::string column_str=WideCharConvertMultiByte(column);
//    	ASSERT_MSG(excel_file.empty(),"字符串为空 %s",excel_file);
    while (g_Running) {
        std::string clip = get_clipboard_text();
        if (!clip.empty() && clip != last) {
            last = clip;
            auto matches = search_excel_all_sheets(excel_file, CharsetConvert(clip.c_str(),GetACP(),CP_UTF8),column_str);
            if (tofile) log_file(matches,excel_file,clip,detail,column_str);
         
            log_console(matches,excel_file,clip,detail,column_str);
        
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
//	    	std::cerr << "开始启动循环。\n";
    g_StatusHandle = RegisterServiceCtrlHandler(L"ExcelSearchService", ServiceCtrlHandler);
    g_ServiceStatus = { SERVICE_WIN32_OWN_PROCESS, SERVICE_START_PENDING, 0, 0, 0, 0, 0 };
    SetServiceStatus(g_StatusHandle, &g_ServiceStatus);

    g_ServiceStatus.dwCurrentState = SERVICE_RUNNING;
    SetServiceStatus(g_StatusHandle, &g_ServiceStatus);

    setup_utf8_console();
    std::wstring excelFile = L"20250304.xlsm"; // TODO: 替换为你的实际路径

    g_WorkerThread = std::thread(run_service_loop, excelFile,TRUE,L"A",TRUE);
}

int wmain(int argc, wchar_t* argv[]) {
	    setup_utf8_console();
    auto args = parse_args(argc, argv);
    std::wstring filename = (args.find(L"filename") != args.end()) ?args[L"filename"]:L"20250304.xlsm";
//    std::wstring keyword = args[L"keyword"];
    std::wstring column = (args.find(L"column") != args.end()) ?args[L"column"]:L"A"; //  默认 A 列
    std::wstring mode = (args.find(L"mode") != args.end()) ?args[L"mode"]:L"debug"; //  
    std::wstring outdetail = (args.find(L"outdetail") != args.end()) ?args[L"outdetail"]:L"1"; // 
    std::wstring tofile = (args.find(L"tofile") != args.end()) ?args[L"tofile"]:L"1"; // 

//    	 std::wcout<<filename<<"\n"<<column<<"\n"<<mode<<"\n"<<outdetail<<"\n"<<tofile<<"\n";
  
   if (!mode.empty() && _wcsicmp(mode.c_str(),L"debug") == 0) {
        // ===== 调试模式 =====  
        
        run_service_loop(filename,_wtoi(outdetail.c_str()),column,_wtoi(tofile.c_str())); 
        return 0;
    }

    SERVICE_TABLE_ENTRY ServiceTable[] = {
        { (LPWSTR)L"ExcelSearchService", ServiceMain },
        { nullptr, nullptr }
    };

    if (!StartServiceCtrlDispatcher(ServiceTable)) {
        std::cerr << "无法启动服务控制调度器。\n";
    }
    return 0;
}
