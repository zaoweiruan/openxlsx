#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <map>
//#include <io.h>
#include <fcntl.h>
//#include <exception>
#include <codecvt>
#include <locale>
#include <stdexcept> 
#include <optional>
#include <cstdint>
#include <openxlsx.hpp> // OpenXLSX 头文件
#include <fstream>
#include <sstream>
#include <iomanip>
#include <ctime>
using namespace OpenXLSX;
struct Config {
    std::wstring output_path = L"search_log.txt"; // 默认输出文件
    std::wstring input_path = L"20250304.xlsm"; // 默认输入文件
    std::wstring column = L"A"; // 默认输入文件
    bool detail = false;                        // 是否显示详细日志
    bool logTofile=false;					//是否将查找到的书籍目录写入到文件
    bool show_help = false;                       // 是否显示帮助信息
} cfg;

// 解析宽字符命令行参数
void parse_command_line(LPWSTR lpCmdLine,Config &cfg) {
    //Config cfg;
    std::vector<std::wstring> args;

    // 分割 lpCmdLine 为单个参数（简易分割，处理空格分隔的参数）
    wchar_t* token = wcstok(lpCmdLine, L" ");
    while (token != nullptr) {
        args.emplace_back(token);
        token = wcstok(nullptr, L" ");
    }

    // 解析参数
    for (size_t i = 0; i < args.size(); ++i) {
        if (args[i] == L"--file" || args[i] == L"-f") {
			if (i + 1 < args.size()) {
				cfg.input_path = args[++i]; // 取后续参数作为输入路径
			} else {
				cfg.show_help = true;
				std::wcerr << L"错误：--file 需要指定已下载书籍文件路径" << std::endl;
			}
		} else if (args[i] == L"--detail" || args[i] == L"-d") {
			cfg.detail = true; // 启用详细日志
		} else if (args[i] == L"--logTofile" || args[i] == L"-l") {
			cfg.logTofile = true; // 启用日志到文件

             if (i + 1 < args.size()) {
                cfg.output_path = args[++i]; // 取后续参数作为输出路径
              }
		}else if (args[i] == L"--help" || args[i] == L"-h") {
            cfg.show_help = true;
        } else if (args[i] == L"--column" || args[i] == L"-c") {
            if (i + 1 < args.size()) {
				cfg.column = args[++i]; // 取后续参数作为列字母
			} else {
				cfg.show_help = true;
				std::wcerr << L"错误：--column 需要指定列字母" << std::endl;
			}

        }

		else {
            cfg.show_help = true;
            std::wcerr << L"错误：未知参数 " << args[i] << std::endl;
        }
    }

    return ;
}

// 显示帮助信息
void show_help(HINSTANCE hInstance) {
    std::wcout << L"剪贴板监听查找书籍工具" << std::endl;
    std::wcout << L"用法：search_book_service [选项]" << std::endl;
    std::wcout << L"选项：" << std::endl;
    std::wcout << L"  --file <路径> (-f)  已下载书籍文件路径（默认：20250304.xlsm）" << std::endl;
    std::wcout << L"  --logTofile <路径> (-l)  是否保存查找结果及路径（默认：不保存，search_log.txt）" << std::endl;
    std::wcout << L"  --column <路径> (-c)  查找列（默认：A列-书名，其它列：B列-文件长度、C列-文件日期）" << std::endl;
    std::wcout << L"  --detail (-d)        显示详细日志" << std::endl;
    std::wcout << L"  --help (-h)           显示此帮助信息" << std::endl;
}
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

std::string Utf16ToUtf8(const std::wstring &utf16)
	{
		 std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> ucs2conv;
			return ucs2conv.to_bytes(utf16);
	}
// 自定义宽字符异常类
#if 0
class WideException : public std::runtime_error {
private:
    std::wstring wmsg; // 存储宽字符异常信息
//    mutable std::string msg; 
    	// 缓存窄字符版本（用于 what()）

public:
    // 构造函数：接受宽字符字符串
    explicit WideException(const std::wstring& message) :std::runtime_error(Utf16ToUtf8(message)),wmsg(message) {}
     WideException(const std::exception& e) : std::runtime_error(e.what()),wmsg(Utf8ToUtf16(e.what())) {}

    // 重写 what()：返回窄字符版本（需转换）
    const char* what() const noexcept override {
        // 宽字符转窄字符（UTF-8）
        std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> converter;
        msg = converter.to_bytes(wmsg); // 转换失败会抛出，但 what() 不允许抛异常，需处理
        return msg.c_str();
    }

    // 提供宽字符信息获取接口
    const std::wstring& wide_what() const noexcept {
        return wmsg;
    }
};
#endif
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
std::string tmToString(const std::tm& timeStruct, const std::string& format = "%Y-%m-%d %H:%M:%S") {
    std::ostringstream oss;
    oss << std::put_time(&timeStruct, format.c_str());
    return oss.str();
}
std::tm excelDateToTm(double excelDate) {
    // Excel date starts from 1900-01-01
    const time_t baseDate = -2209161600; // Unix timestamp for 1900-01-01
    time_t timestamp = baseDate + static_cast<time_t>(excelDate * 86400); // Convert days to seconds
    std::tm dateTime = *std::localtime(&timestamp);
    return dateTime;
}
bool isValidExcelDate(double value) {
    // Excel date range: 1 (1900-01-01) to 2958465 (9999-12-31)
    return value >= 1.0 && value <= 2958465.0;
}
std::string  processDatetimeCell(const XLCellValue& cellValue) {

        double excelDate = cellValue.get<double>();
        std::tm dateTime = excelDateToTm(excelDate);
        return tmToString(dateTime, "%Y-%m-%d %H:%M:%S");
}
std::vector<Match> find_string_in_sheet(XLWorksheet& sheet, const std::string& target, uint16_t col) {
    std::vector<Match> matches;
    uint32_t maxRow = detect_max_row(sheet, col);
    std::string val="";
    for (uint32_t row = 1; row <= maxRow; ++row) {
        auto cell = sheet.cell(XLCellReference(row, col));
        if (cell.value().type() == XLValueType::Empty) continue;

//        std::string val = cell.value().get<std::string>();
		if (cell.value().type() == XLValueType::Float&&isValidExcelDate(cell.value().get<double>())) {

			val=processDatetimeCell(cell.value());
			}

		else
		{
			val = cell.value().getString();
		}
//        std::string val = cell.get<std::string>(format::string);
        if (val.find(target) != std::string::npos) {
            matches.emplace_back(sheet.name(), cell.cellReference().address(),val);
        }
    }

    return matches;
}

std::vector<Match> search_excel_all_sheets(const std::string& filepath, const std::string& keyword,const std::string& column) {
	
	  std::vector<Match> results;
   	
 /*   std::string excel_file=WideCharConvertMultiByte(filepath,GetACP());
    std::string column_str=WideCharConvertMultiByte(column,GetACP());	
    std::string keyword_str=WideCharConvertMultiByte(keyword,GetACP());	*/
    

    OpenXLSX::XLDocument doc;
 try{
    doc.open(filepath);
    }
    catch (std::runtime_error &r)
	  {
			throw std::runtime_error ("打开文件: "+filepath+" 出错,错误为："+r.what());
		}
    auto colOpt = column_letter_to_index(column);
    if (!colOpt) {
    	
        throw std::runtime_error ("无效的列字母: "+column);
//        return results;
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


// ==== 日志输出 ====
void log_console(const std::vector<Match>& matches,std::wstring filename, const std::wstring& keyword,bool detail=FALSE,std::wstring column=L"A") {
	std::wostream *log = &std::wcout;  // 默认输出到控制台
      
    	if (detail)
      {for (const auto& match : matches) {
      	*log << L"[" <<MultiByteConvertWideChar(match.sheetName,CP_UTF8) << L"] "\
      	<< MultiByteConvertWideChar(match.cellAddress,CP_UTF8)<< L" -> " \
      	<< MultiByteConvertWideChar(match.value,CP_UTF8) << "\n";}
      }
     *log << L"=== 搜索: " <<column<<L" 列 "<<L" 关键词：" <<keyword << L" 匹配结果共 ："<<matches.size()<<L" 条记录=== "<<" \n";		

  
}
// ==== 日志输出 ====
void log_file(const std::vector<Match>& matches,std::string filename, const std::string& keyword,std::string column="A",bool detail=FALSE) {
    std::ofstream log(filename, std::ios::app);
//    log.imbue(std::locale(log.getloc(), new std::codecvt_utf8<wchar_t>()));
 
    	if (detail)
      {for (const auto& match : matches) {
      	log << "[" << match.sheetName << "] " << match.cellAddress << " -> " << match.value << "\n";}
      }
     log <<  "=== 搜索: " <<column<< " 列   关键词：" << keyword << " 匹配结果共 ："<<matches.size()<<" 条记录 ===\n";		
    log.close();
    }

void search_keyword(const std::wstring& excel_path,const std::wstring &output_path,const std::wstring &keyword,std::wstring column=L"A",bool detail=TRUE,bool tofile=FALSE) {
   
    std::string excel_file=WideCharConvertMultiByte(excel_path,GetACP());
    std::string output_file=WideCharConvertMultiByte(output_path,GetACP());
    std::string column_str=WideCharConvertMultiByte(column,GetACP());
    std::string keyword_str=WideCharConvertMultiByte(keyword,GetACP());
//    	ASSERT_MSG(excel_file.empty(),"字符串为空 %s",excel_file);
    if (!keyword.empty()){
           
            auto matches = search_excel_all_sheets(excel_file,keyword_str,column_str);
            if (tofile) log_file(matches,output_file,keyword_str,column_str,detail);
         
            log_console(matches,excel_path,keyword,detail,column);
        
        }
       
    }


// 全局窗口类名和窗口句柄
const wchar_t* g_ClassName = L"ClipboardListenerClass";
HWND g_Hwnd = nullptr;

// 读取剪贴板中的文本内容（返回 Unicode 字符串）
std::wstring GetClipboardText() {
    std::wstring text;

    // 打开剪贴板（需传入窗口句柄）
    if (!OpenClipboard(g_Hwnd)) {
        std::wcerr << L"OpenClipboard 失败，错误码：" << GetLastError() << std::endl;
        return text;
    }

    // 优先读取 Unicode 文本（CF_UNICODETEXT）
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData != nullptr) {
        // 锁定内存块并获取文本指针
        wchar_t* pText = static_cast<wchar_t*>(GlobalLock(hData));
        if (pText != nullptr) {
            text = pText; // 复制文本内容
            GlobalUnlock(hData); // 解锁内存
            
        }
    } else {
        // 若没有 Unicode 文本，尝试读取 ANSI 文本（CF_TEXT）并转换为 Unicode
        HANDLE hAnsiData = GetClipboardData(CF_TEXT);
        if (hAnsiData != nullptr) {
            char* pAnsiText = static_cast<char*>(GlobalLock(hAnsiData));
            if (pAnsiText != nullptr) {
                // ANSI 转 Unicode（使用 MultiByteToWideChar）
                int len = MultiByteToWideChar(CP_ACP, 0, pAnsiText, -1, nullptr, 0);
                if (len > 0) {
                    std::vector<wchar_t> buf(len);
                    MultiByteToWideChar(CP_ACP, 0, pAnsiText, -1, buf.data(), len);
                    text = buf.data();
                }
                GlobalUnlock(hAnsiData);
            }
        }
    }
    CloseClipboard(); // 关闭剪贴板
    return text;
  
    
}

// 窗口过程（处理剪贴板更新消息）
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	    static std::wstring last=L"";

	try{
    switch (msg) {
        case WM_CLIPBOARDUPDATE:
            // 剪贴板内容更新，读取内容
           { std::wstring clipboardText = GetClipboardText();
            if (!clipboardText.empty()&&last!=clipboardText) {
    						last=clipboardText;
    					search_keyword(cfg.input_path,cfg.output_path,clipboardText,cfg.column,cfg.detail,cfg.logTofile);
    								
            }            
            return 0;
					}
        case WM_DESTROY:
        	{
            // 注销剪贴板监听并退出
            RemoveClipboardFormatListener(hwnd);
            PostQuitMessage(0);
            return 0;
					}
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    }
    catch(std::runtime_error &e)
    {
    	std::wcout<<MultiByteConvertWideChar(e.what(),GetACP())<<"\n";
    		    RemoveClipboardFormatListener(hwnd);
            PostQuitMessage(0);
            return 0;
    }
}

// 注册窗口并启动消息循环
void StartClipboardListener() {
    // 1. 注册窗口类
    WNDCLASSEX wc = {0};
    wc.cbSize        = sizeof(WNDCLASSEX);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = GetModuleHandle(nullptr);
    wc.lpszClassName = g_ClassName;

    if (!RegisterClassEx(&wc)) {
        std::cerr << "窗口类注册失败" << std::endl;
        return;
    }

    // 2. 创建隐藏窗口（无需显示，仅用于接收消息）
    g_Hwnd = CreateWindowEx(
        0,
        g_ClassName,
        L"剪贴板监听窗口",
        0, // 隐藏窗口（无样式）
        CW_USEDEFAULT, CW_USEDEFAULT, 0, 0,
        HWND_MESSAGE, // 消息窗口（不显示在任务栏）
        nullptr,
        GetModuleHandle(nullptr),
        nullptr
    );

    if (!g_Hwnd) {
        std::cerr << "窗口创建失败" << std::endl;
        return;
    }

    // 3. 注册剪贴板监听（接收 WM_CLIPBOARDUPDATE 消息）
    if (!AddClipboardFormatListener(g_Hwnd)) {
        std::cerr << "注册剪贴板监听失败" << std::endl;
        DestroyWindow(g_Hwnd);
        return;
    }

    // 4. 启动消息循环
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

//int main() {
int WINAPI wWinMain(
    HINSTANCE hInstance,    // 当前实例句柄
    HINSTANCE hPrevInstance,// 前一个实例句柄（已废弃，通常为 NULL）
    LPWSTR lpCmdLine,       // 宽字符命令行参数
    int nCmdShow            // 窗口显示方式（如 SW_SHOW）
) 
{
	setup_utf8_console();
	parse_command_line(lpCmdLine,cfg);


    // 3. 处理帮助信息
    if (cfg.show_help) {
        show_help(hInstance);
        return 0;
    }

    // 4. 输出参数配置（演示用）
    std::wcout << L"===== 配置信息 =====" << std::endl;
    std::wcout << L"书名文件路径：" << cfg.input_path << std::endl;
    std::wcout << L"输出日志到文件：" << (cfg.logTofile ? L"开启" : L"关闭") << std::endl;
    std::wcout << L"输出文件路径：" << cfg.output_path << std::endl;
    std::wcout << L"查找列：" << cfg.column << std::endl;
    std::wcout << L"详细日志：" << (cfg.detail ? L"开启" : L"关闭") << std::endl;
    std::wcout << L"开始监听剪贴板更新（按 Ctrl+C 退出）..." << std::endl;
    StartClipboardListener();
    return 0;
}
