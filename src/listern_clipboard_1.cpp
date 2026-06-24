#include <winsock2.h>
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
#include <cstdlib>

#include <cstdio>
#include <memory>
#include <array>

#include <set>
#include <regex>
#include <algorithm>

#include <ws2tcpip.h>


using namespace OpenXLSX;


struct Config {
    std::wstring output_path = L"search_log.txt"; // 默认输出文件
    std::wstring input_path = L"20250304.xlsm"; // 默认输入文件
    std::wstring column = L"A"; // 默认搜索列
    std::wstring config_path = L".\\config.ini"; // 默认配置文件路径
    std::set<wchar_t> illegalChars; // 默认非法字符集合（从配置文件加载）
    std::map < std::wstring, std::wstring > extraConfig; // 其他配置项;
    bool detail = false;                        // 是否显示详细日志
    bool logTofile=false;					//是否将查找到的书籍目录写入到文件
    bool show_help = false;                       // 是否显示帮助信息
    bool autorefresh=false;						// 是否自动refreshE:\My Kindle Content目录文件
    std::wstring errMsg=L"";
} cfg;


// UTF-8 文件读取函数（MinGW codecvt 不可靠，使用 Windows API）
inline std::wstring ReadUtf8TextFile(const std::wstring& path) {
    // 调试输出
    std::wcerr << L"[DBG ReadUtf8] 尝试打开: " << path << std::endl;
    
    // 使用 Windows API 打开宽字符串路径的文件
    FILE* fp = nullptr;
    errno_t err = _wfopen_s(&fp, path.c_str(), L"rb");
    if (err != 0 || !fp) {
        std::wcerr << L"[DBG ReadUtf8] _wfopen_s 失败, err=" << err << std::endl;
        return L"";
    }
    
    std::string content;
    char buffer[4096];
    size_t bytesRead;
    while ((bytesRead = fread(buffer, 1, sizeof(buffer), fp)) > 0) {
        content.append(buffer, bytesRead);
    }
    fclose(fp);
    
    if (content.empty()) return L"";
    
    // 跳过 UTF-8 BOM（EF BB BF）
    size_t start = 0;
    if (content.size() >= 3 && 
        static_cast<unsigned char>(content[0]) == 0xEF &&
        static_cast<unsigned char>(content[1]) == 0xBB &&
        static_cast<unsigned char>(content[2]) == 0xBF) {
        start = 3;
    }
    
    // 转换为 UTF-16 LE
    int wlen = MultiByteToWideChar(CP_UTF8, 0, content.data() + start, -1, nullptr, 0);
    if (wlen <= 0) return L"";
    std::wstring result(wlen - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, content.data() + start, -1, &result[0], wlen);
    return result;
}

bool LoadIllegalChars(Config &cfg) {
    std::set<wchar_t> illegalChars;

    // 调试：输出尝试读取的路径
    std::wcerr << L"[DBG] 尝试打开非法字符文件: " << cfg.extraConfig[L"illegalcharsfile"] << std::endl;
    std::wstring content = ReadUtf8TextFile(cfg.extraConfig[L"illegalcharsfile"]);
    std::wcerr << L"[DBG] 读取结果长度: " << content.size() << std::endl;
    if (content.empty()) {
    	cfg.errMsg+= L"非法字符文件打开失败: " +cfg.extraConfig[L"illegalcharsfile"]+L'\n';
    	return false;
    }
    std::wistringstream sin(content);
    wchar_t ch;
    while (sin >> ch) {
    	cfg.illegalChars.insert(ch);
        sin.ignore(256, L'\n');
    }
    return true;
}
bool LoadConfig(struct Config &cfg) {
    // 使用平台函数读取 UTF-8 配置文件（MinGW codecvt 不可靠）
    std::wstring content = ReadUtf8TextFile(cfg.config_path);
    if (content.empty()) {
    	cfg.errMsg=L"配置文件打开失败: " + cfg.config_path + L'\n';
        //std::wcerr << L"配置文件打开失败: " << cfg.config_path << L'\n';
        return false;
    }

    cfg.errMsg=L"配置文件加载成功: " + cfg.config_path + L'\n';

    std::wistringstream fin(content);
    std::wstring line;
    while (std::getline(fin, line)) {
        auto pos = line.find(L'=');
        if (pos != std::wstring::npos) {
            std::wstring key = line.substr(0, pos);
            std::wstring value = line.substr(pos + 1);
            // 去掉 value 中的注释（; 开头）
            size_t commentPos = value.find(L';');
            if (commentPos != std::wstring::npos) {
                value = value.substr(0, commentPos);
            }
            // 去掉尾部 \r（CRLF 文件用 getline 会保留 carriage return）
            if (!key.empty() && key.back() == L'\r') key.pop_back();
            if (!value.empty() && value.back() == L'\r') value.pop_back();
            cfg.extraConfig[key] = value;
        }
    }

    return LoadIllegalChars(cfg);
}
// 判断 text 是否包含非法字符
inline bool ContainsIllegalChar(const std::wstring &text,
		const std::set<wchar_t> &illegalChars) {
	for (wchar_t c : text) {
		//if (illegalChars.count(c) || iswspace(c)) {
		if (illegalChars.count(c)) {
			return true;
		}
	}
	return false;
}
// 判断是否为实数
bool isReal(const std::wstring& text) {
	double outValue;
	std::wistringstream iss(text);
	    // 尝试转换
	    if (!(iss >> outValue)) {
	        return false;
	    }
	    // 确保没有多余字符（避免 "123abc" 这种被误转）
	    wchar_t c;
	    if (iss >> c) {
	        return false;
	    }
	    return true;
}
// 判断是否为 CJK 字符（中文、日文、韩文等）
inline bool is_cjk_char(wchar_t wc) {
    return (wc >= 0x4E00 && wc <= 0x9FFF) ||    // 基本汉字
           (wc >= 0x3400 && wc <= 0x4DBF) ||    // 扩展A
           (wc >= 0x3000 && wc <= 0x30FF) ||    // 日文平假名/片假名等
           (wc >= 0xFF00 && wc <= 0xFFEF);      // 全角
}

// 判断整个 std::wstring 是否**包含**中文字符
inline bool has_chinese(const std::wstring& ws) {
    for (wchar_t wc : ws) {
        if (is_cjk_char(wc)) {
            return true;
        }
    }
    return false;
}
// 判断是否为有效的 IPv4 地址


inline std::wstring trim(const std::wstring& str)
{
    size_t first = str.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos)
        return L"";

    size_t last = str.find_last_not_of(L" \t\r\n");
    return str.substr(first, last - first + 1);
}


bool isValidIPv4(const std::wstring& ip)
{
    sockaddr_in sa;
    return InetPtonW(AF_INET, ip.c_str(), &(sa.sin_addr)) == 1;
}

bool domainExists(const std::wstring& domain)
{
    ADDRINFOW hints = {0};
    ADDRINFOW* result = nullptr;

    hints.ai_family = AF_UNSPEC;

    int ret = GetAddrInfoW(domain.c_str(), nullptr, &hints, &result);

    if (ret == 0)
    {
        FreeAddrInfoW(result);
        return true;
    }
    return false;
}

// 前向声明
void AppendLogLine(const std::wstring& line);

// 调用 Python 脚本并显示结果
void RunPythonAndShowResult() {

	std::wstring cmd =
			L"(python.exe "+cfg.extraConfig[L"autorefreshfile"]+L")";
			//LR"(E:\Python312\python.exe D:\mvnworkspace\example\openxlsx\src\rungetfilenamefrompath.py)";

    AppendLogLine(L"=== 刷新目录: 开始 ===");

    FILE* pipe = _wpopen(cmd.c_str(), L"r");
    if (!pipe){

		DWORD err = GetLastError();
		std::wcout << L"无法启动 Python 脚本" << std::endl;
        std::wcout << L"错误码: " << err << std::endl;
        AppendLogLine(L"错误：无法启动 Python 脚本，错误码: " + std::to_wstring(err));
        wchar_t msgBuf[512];
        FormatMessageW(
            FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            err,
            0,
            msgBuf,
            sizeof(msgBuf) / sizeof(wchar_t),
            NULL
        );

		return;
	}
    std::array<char, 512> buffer;
    while (fgets(buffer.data(), buffer.size(), pipe)) {
        // 将多字节转换为宽字符
        int wlen = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), -1, NULL, 0);
        if (wlen > 0) {
            std::wstring wbuffer(wlen, 0);
            MultiByteToWideChar(CP_UTF8, 0, buffer.data(), -1, &wbuffer[0], wlen);
            std::wcout << wbuffer.c_str();
        }
    }
	_pclose(pipe);
    AppendLogLine(L"=== 刷新目录: 完成 ===");
}



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
		}
		else if (args[i] == L"--autorefresh" || args[i] == L"-a") {
					cfg.autorefresh = true; // 自动refreshE:\My Kindle Content目录文件
				}
		else if (args[i] == L"--logTofile" || args[i] == L"-l") {
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
    std::wcout << L"  --autorefresh (-a)        自动refresh E:\\My Kindle Content目录文件" << std::endl;

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

inline std::string Utf16ToUtf8(const std::wstring &utf16)
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
inline std::string tmToString(const std::tm& timeStruct, const std::string& format = "%Y-%m-%d %H:%M:%S") {
    std::ostringstream oss;
    oss << std::put_time(&timeStruct, format.c_str());
    return oss.str();
}
inline std::tm excelDateToTm(double excelDate) {
    // Excel date starts from 1900-01-01
    const time_t baseDate = -2209161600; // Unix timestamp for 1900-01-01
    time_t timestamp = baseDate + static_cast<time_t>(excelDate * 86400); // Convert days to seconds
    std::tm dateTime = *std::localtime(&timestamp);
    return dateTime;
}
inline bool isValidExcelDate(double value) {
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

// 辅助：当 logTofile 启用时追加一行到日志文件（UTF-8）
void AppendLogLine(const std::wstring& line) {
    if (!cfg.logTofile) return;
    std::string outPath = WideCharConvertMultiByte(cfg.output_path, CP_UTF8);
    std::ofstream log(outPath, std::ios::app);
    if (!log.is_open()) return;
    std::string utf8Line = WideCharConvertMultiByte(line, CP_UTF8);
    log << utf8Line.c_str() << std::endl;
    log.close();
}

// 全局窗口类名和窗口句柄
const wchar_t* g_ClassName = L"ClipboardListenerClass";
HWND g_Hwnd = nullptr;

// 读取剪贴板中的原始文本内容（不作任何过滤）
std::wstring GetRawClipboardText() {
    std::wstring text;

    // 打开剪贴板（需传入窗口句柄）

    if (!OpenClipboard(g_Hwnd)) {
    	DWORD err = GetLastError();
        //std::wcerr << L"OpenClipboard 失败，错误码：" << err << std::endl;
    	LPWSTR buffer = nullptr;
    	FormatMessageW(
    	    FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
    	    NULL,
    	    err,
    	    MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
    	    (LPWSTR)&buffer,
    	    0,
    	    NULL
    	);

    	std::wcout << L"OpenClipboard 失败，错误码：" << err << L"\n错误信息：" << buffer << std::endl;

    	LocalFree(buffer);
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

// 对文本应用搜索专用的过滤链（非法字符、IP/域名等）
std::wstring FilterForSearch(const std::wstring& text) {
    if (text.empty()) return L"";

    std::wstring result = text;

    // 字符校验：全空、全数字、包含非法字符、IP地址、域名等都视为无效输入，返回空字符串
    if (std::all_of(result.begin(), result.end(), [](wchar_t c) {
        return iswspace(c)||iswdigit(static_cast<wint_t>(c))||c == L'.'||c == L'-';
    }) || ContainsIllegalChar(result, cfg.illegalChars)) {
        return L"";
    }
    if (has_chinese(result)) {
        result = trim(result);
    }
    else if (isValidIPv4(result)||domainExists(result)) {
        return L"";
    }
    return result;
}

// 窗口过程（处理剪贴板更新消息）
// 前向声明：批量重命名函数
void BatchRename(const std::wstring& dirPath, const std::wstring& rnpPath);

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	    static std::wstring last=L"";

	try{
    switch (msg) {
        case WM_CLIPBOARDUPDATE:
            // 剪贴板内容更新，读取原始文本（不含过滤）
            { std::wstring rawText = GetRawClipboardText();
            if (!rawText.empty()&&last!=rawText) {
    						last=rawText;
                // 诊断日志：输出收到的原始剪贴板内容（未过滤，便于调试）
                std::wcout << L"[DEBUG] 收到剪贴板: " << rawText << std::endl;
                AppendLogLine(L"[DEBUG] 收到剪贴板: " + rawText);

                // 特殊命令检查：必须在过滤之前，避免 `\` 等字符被非法字符过滤丢弃
                if (rawText == L"刷新My Kindle Content") {
                    // 调用 Python 脚本
                	RunPythonAndShowResult();
                } else if ([&]() -> bool {
                    // 从 config 读取批量重命名触发前缀，支持空值禁用
                    std::wstring prefix = L"批量重命名 ";
                    auto it = cfg.extraConfig.find(L"renamerprefix");
                    if (it != cfg.extraConfig.end()) {
                        if (it->second.empty()) return false; // 空值 = 禁用
                        prefix = it->second;
                    }
                    if (rawText.find(prefix) != 0) {
                        AppendLogLine(L"[DBG] 前缀不匹配: prefix='" + prefix + L"'");
                        return false;
                    }
                    AppendLogLine(L"[DBG] 前缀匹配成功，prefix='" + prefix + L"'");
    std::wstring dirPath = trim(rawText.substr(prefix.size()));
    dirPath = trim(dirPath);
    // 自动剥离路径两端双引号（用户从资源管理器复制路径时可能带引号）
    if (dirPath.size() >= 2 && dirPath.front() == L'"' && dirPath.back() == L'"') {
        dirPath = dirPath.substr(1, dirPath.size() - 2);
    }
                    AppendLogLine(L"[DBG] dirPath 解析结果: '" + dirPath + L"', 长度=" + std::to_wstring(dirPath.size()) + L", 为空=" + (dirPath.empty() ? L"是" : L"否"));
    if (!dirPath.empty()) {
        std::wcout << L"批量重命名指令已识别，目标: " << dirPath << std::endl;
                        AppendLogLine(L"批量重命名指令已识别，目标: " + dirPath);
                        auto it = cfg.extraConfig.find(L"renamerpreset");
                        if (it != cfg.extraConfig.end() && !it->second.empty()) {
                            wchar_t absPath[MAX_PATH] = {0};
                            GetFullPathNameW(it->second.c_str(), MAX_PATH, absPath, nullptr);
                            AppendLogLine(L"规则文件: " + std::wstring(absPath));
                            BatchRename(dirPath, absPath);
                        } else {
                            std::wcout << L"错误：config.ini 中未配置 renamerpreset" << std::endl;
                            AppendLogLine(L"错误：config.ini 中未配置 renamerpreset");
                        }
                    } else {
                        AppendLogLine(L"[DBG] dirPath 为空，放弃批量重命名");
                    }
                    return true;
                }()) {
                    // lambda 已处理，条件成功
                } else {
                    // 普通搜索路径：此时才应用过滤链
                    std::wstring searchText = FilterForSearch(rawText);
                    if (!searchText.empty()) {
                        search_keyword(cfg.input_path, cfg.output_path, searchText, cfg.column, cfg.detail, cfg.logTofile);
                    }
                }
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

// ============================================================
// 批量重命名功能 - ReNamer .rnp 预设解析与重命名引擎
// ============================================================

enum class RnpRuleType { Remove, Replace, RegEx };

struct RnpRule {
    RnpRuleType type = RnpRuleType::Remove;
    bool enabled = true;
    std::wstring matchText;
    std::wstring replaceText;
    bool skipExtension = true;
    bool caseSensitive = false;
    bool useWildcards = false;
    bool wholeWordsOnly = false;
    int which = 3; // 1=first, 2=last, 3=all
};

// URL 解码：将 %XX 编码的 UTF-8 序列解码为宽字符串
std::wstring UrlDecode(const std::wstring& input) {
    if (input.empty()) return L"";
    std::string utf8bytes;
    utf8bytes.reserve(input.size());
    for (size_t i = 0; i < input.size(); ++i) {
        if (input[i] == L'%' && i + 2 < input.size()) {
            wchar_t hex[3] = { input[i+1], input[i+2], 0 };
            wchar_t* end;
            unsigned long byteVal = wcstoul(hex, &end, 16);
            if (end == hex + 2) {
                utf8bytes.push_back(static_cast<char>(byteVal));
                i += 2; continue;
            }
        } else if (input[i] == L'+') {
            utf8bytes.push_back(' '); continue;
        }
        char utf8Buf[4] = {0};
        int len = WideCharToMultiByte(CP_UTF8, 0, &input[i], 1, utf8Buf, 4, nullptr, nullptr);
        for (int j = 0; j < len; ++j) utf8bytes.push_back(utf8Buf[j]);
    }
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8bytes.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return L"";
    std::wstring result(wlen - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8bytes.c_str(), -1, &result[0], wlen);
    return result;
}

// 解析 .rnp 预设文件为规则列表
std::vector<RnpRule> ParseRnpRules(const std::wstring& rnpPath) {
    std::vector<RnpRule> rules;
    // 使用平台函数读取 UTF-8 .rnp 文件
    std::wstring content = ReadUtf8TextFile(rnpPath);
    if (content.empty()) {
        std::wcout << L"错误：无法打开 .rnp 文件: " << rnpPath << std::endl;
        return rules;
    }
    std::wistringstream fin(content);
    std::wstring line;
    std::map<std::wstring, std::wstring> ruleKV;
    bool inRule = false;

    auto finishRule = [&]() {
        if (!inRule || !ruleKV.count(L"ID")) return;
        RnpRule rule;
        std::wstring id = ruleKV[L"ID"];
        if      (id == L"Remove")  rule.type = RnpRuleType::Remove;
        else if (id == L"Replace") rule.type = RnpRuleType::Replace;
        else if (id == L"RegEx")   rule.type = RnpRuleType::RegEx;
        else return;
        if (ruleKV.count(L"Marked")) rule.enabled = (ruleKV[L"Marked"] == L"1");
        if (ruleKV.count(L"Config")) {
            std::wistringstream cs(ruleKV[L"Config"]);
            std::wstring seg;
            while (std::getline(cs, seg, L';')) {
                auto colon = seg.find(L':');
                if (colon == std::wstring::npos) continue;
                std::wstring k = seg.substr(0, colon);
                std::wstring v = UrlDecode(seg.substr(colon + 1));
                if      (k == L"TEXT" || k == L"TEXTWHAT" || k == L"EXPRESSION") rule.matchText = v;
                else if (k == L"TEXTWITH" || k == L"REPLACE")                    rule.replaceText = v;
                else if (k == L"WHICH")                                          rule.which = std::stoi(v);
                else if (k == L"SKIPEXTENSION")  rule.skipExtension  = (v == L"1");
                else if (k == L"CASESENSITIVE")  rule.caseSensitive  = (v == L"1");
                else if (k == L"USEWILDCARDS")   rule.useWildcards   = (v == L"1");
                else if (k == L"WHOLEWORDSONLY") rule.wholeWordsOnly = (v == L"1");
            }
        }
        rules.push_back(rule);
    };

    while (std::getline(fin, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        if (line.empty()) continue;
        if (line.front() == L'[' && line.back() == L']') { finishRule(); ruleKV.clear(); inRule = true; continue; }
        if (inRule) {
            auto eq = line.find(L'=');
            if (eq != std::wstring::npos) ruleKV[line.substr(0, eq)] = line.substr(eq + 1);
        }
    }
    finishRule();
    return rules;
}

// 对单个文件基名应用所有启用规则
std::wstring ApplyRules(const std::wstring& baseName, const std::vector<RnpRule>& rules) {
    if (rules.empty()) return baseName;
    std::wstring result = baseName;

    for (const auto& rule : rules) {
        if (!rule.enabled || rule.matchText.empty()) continue;
        try {
            if (rule.type == RnpRuleType::Remove || rule.type == RnpRuleType::Replace) {
                bool isReplace = (rule.type == RnpRuleType::Replace);
                auto ciSearch = [&](const std::wstring& s, size_t start) -> std::wstring::size_type {
                    auto it = std::search(s.begin() + start, s.end(),
                        rule.matchText.begin(), rule.matchText.end(),
                        [](wchar_t a, wchar_t b) { return towlower(a) == towlower(b); });
                    return (it != s.end()) ? (it - s.begin()) : std::wstring::npos;
                };
                auto ciRSearch = [&](const std::wstring& s) -> std::wstring::size_type {
                    std::wstring ls = s, lm = rule.matchText;
                    std::transform(ls.begin(), ls.end(), ls.begin(), towlower);
                    std::transform(lm.begin(), lm.end(), lm.begin(), towlower);
                    return ls.rfind(lm);
                };
                if (rule.which == 1) { // first
                    auto pos = rule.caseSensitive ? result.find(rule.matchText) : ciSearch(result, 0);
                    if (pos != std::wstring::npos) {
                        if (isReplace) result.replace(pos, rule.matchText.size(), rule.replaceText);
                        else           result.erase(pos, rule.matchText.size());
                    }
                } else if (rule.which == 2) { // last
                    auto pos = rule.caseSensitive ? result.rfind(rule.matchText) : ciRSearch(result);
                    if (pos != std::wstring::npos) {
                        if (isReplace) result.replace(pos, rule.matchText.size(), rule.replaceText);
                        else           result.erase(pos, rule.matchText.size());
                    }
                } else { // all (3)
                    std::wstring::size_type start = 0;
                    while (true) {
                        auto pos = rule.caseSensitive ? result.find(rule.matchText, start) : ciSearch(result, start);
                        if (pos == std::wstring::npos) break;
                        if (isReplace) { result.replace(pos, rule.matchText.size(), rule.replaceText); start = pos + rule.replaceText.size(); }
                        else           { result.erase(pos, rule.matchText.size());                   start = pos; }
                    }
                }
            } else if (rule.type == RnpRuleType::RegEx) {
                std::wregex::flag_type flags = std::wregex::ECMAScript;
                if (!rule.caseSensitive) flags |= std::wregex::icase;
                std::wregex re(rule.matchText, flags);
                result = (rule.which == 1)
                    ? std::regex_replace(result, re, rule.replaceText, std::regex_constants::format_first_only)
                    : std::regex_replace(result, re, rule.replaceText);
            }
        } catch (const std::exception& e) {
            std::wcout << L"  规则警告: ";
            const char* msg = e.what();
            int wlen = MultiByteToWideChar(CP_UTF8, 0, msg, -1, nullptr, 0);
            if (wlen > 0) { std::wstring wb(wlen, 0); MultiByteToWideChar(CP_UTF8, 0, msg, -1, &wb[0], wlen); std::wcout << wb; }
            std::wcout << std::endl;
        }
    }
    return result;
}

// 批量重命名：遍历目录内所有文件，应用 .rnp 规则并输出对照表
void BatchRename(const std::wstring& dirPath, const std::wstring& rnpPath) {
    std::wcout << L"正在解析规则文件: " << rnpPath << std::endl;
    AppendLogLine(L"=== 批量重命名: " + dirPath + L" ===");
    AppendLogLine(L"规则文件: " + rnpPath);
    std::vector<RnpRule> rules = ParseRnpRules(rnpPath);
    if (rules.empty()) {
        std::wcout << L"错误：未解析到任何有效规则" << std::endl;
        AppendLogLine(L"错误：未解析到任何有效规则");
        return;
    }

    int enabledCnt = 0;
    for (auto& r : rules) if (r.enabled) enabledCnt++;
    std::wcout << L"共加载 " << rules.size() << L" 条规则，启用 " << enabledCnt << L" 条" << std::endl;
    AppendLogLine(L"共加载 " + std::to_wstring(rules.size()) + L" 条规则，启用 " + std::to_wstring(enabledCnt) + L" 条");

    std::wstring searchPath = dirPath;
    if (searchPath.back() != L'\\') searchPath += L'\\';
    searchPath += L'*';

    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) {
        std::wcout << L"错误：无法打开目录 " << dirPath << std::endl;
        AppendLogLine(L"错误：无法打开目录 " + dirPath);
        return;
    }

    std::wcout << L"----- 重命名对照表 -----" << std::endl;
    int ok = 0, skip = 0;

    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;

        std::wstring orig = fd.cFileName;
        std::wstring base = orig, ext;
        size_t dot = orig.rfind(L'.');
        if (dot != std::wstring::npos && dot > 0) { base = orig.substr(0, dot); ext = orig.substr(dot); }

        std::wstring newBase = ApplyRules(base, rules);
        std::wstring newName = newBase + ext;
        if (newName == orig) continue;

        std::wstring newPath = dirPath + L'\\' + newName;
        if (GetFileAttributesW(newPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
            std::wcout << L"  跳过（已存在）: " << orig << L" -> " << newName << std::endl;
            AppendLogLine(L"  跳过（已存在）: " + orig + L" -> " + newName);
            skip++; continue;
        }

        std::wstring oldPath = dirPath + L'\\' + orig;
        if (MoveFileW(oldPath.c_str(), newPath.c_str())) {
            std::wcout << L"  OK " << orig << L"\n    -> " << newName << std::endl;
            AppendLogLine(L"  OK " + orig + L" -> " + newName);
            ok++;
        } else {
            std::wcout << L"  FAIL: " << orig << L" -> " << newName << L" (error: " << GetLastError() << L")" << std::endl;
            AppendLogLine(L"  FAIL: " + orig + L" -> " + newName + L" (error: " + std::to_wstring(GetLastError()) + L")");
        }
    } while (FindNextFileW(hFind, &fd));

    FindClose(hFind);
    std::wcout << L"----- Done: " << ok << L" ok, " << skip << L" skipped -----" << std::endl;
    AppendLogLine(L"----- Done: " + std::to_wstring(ok) + L" ok, " + std::to_wstring(skip) + L" skipped -----");
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
    //SetConsoleOutputCP(CP_UTF8);
   // SetConsoleCP(CP_UTF8);
	parse_command_line(lpCmdLine,cfg);


    // 3. 处理帮助信息
    if (cfg.show_help) {
        show_help(hInstance);
        return 0;
    }

	if 	(!LoadConfig(cfg))
	{
		AppendLogLine(L"[DBG] LoadConfig 失败: " + cfg.errMsg);
		std::wcout << cfg.errMsg << std::endl;
		return 0;
	} // 加载配置文件（包括非法字符列表）


    // 诊断：打印重命名配置
    {
        auto itPrefix = cfg.extraConfig.find(L"renamerprefix");
        std::wcout << L"[配置] renamerprefix=" << (itPrefix != cfg.extraConfig.end() ? itPrefix->second : L"(未配置)") << std::endl;
        auto itPreset = cfg.extraConfig.find(L"renamerpreset");
        std::wcout << L"[配置] renamerpreset=" << (itPreset != cfg.extraConfig.end() ? itPreset->second : L"(未配置)") << std::endl;
    }

    if (cfg.autorefresh){
    	RunPythonAndShowResult();
    }
    // 4. 输出参数配置（演示用）
    std::wcout << L"===== 配置信息 =====" << std::endl;
    std::wcout << L"书名文件路径：" << cfg.input_path << std::endl;
    std::wcout << L"输出日志到文件：" << (cfg.logTofile ? L"开启" : L"关闭") << std::endl;
    if (cfg.logTofile) std::wcout << L"输出文件路径：" << cfg.output_path << std::endl;
    std::wcout << L"查找列：" << cfg.column << std::endl;
    std::wcout << L"详细日志：" << (cfg.detail ? L"开启" : L"关闭") << std::endl;
    // 打印批量重命名配置
    auto rpIt = cfg.extraConfig.find(L"renamerprefix");
    std::wcout << L"批量重命名前缀：" << (rpIt != cfg.extraConfig.end() ? rpIt->second : L"(未设置)") << std::endl;
    auto rrIt = cfg.extraConfig.find(L"renamerpreset");
    std::wcout << L"规则文件路径：" << (rrIt != cfg.extraConfig.end() ? rrIt->second : L"(未设置)") << std::endl;
    std::wcout << L"开始监听剪贴板更新（按 Ctrl+C 退出）..." << std::endl;
    WSADATA wsaData;
    WSAStartup(MAKEWORD(2,2), &wsaData);
    StartClipboardListener();
    WSACleanup();
    return 0;
}
