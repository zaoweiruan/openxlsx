#include <OpenXLSX.hpp>
#include <iostream>
#include <string>
#include <vector>
#include <optional>
#include <cctype>
#include <windows.h>      // SetConsoleOutputCP
#include <locale>
#include <codecvt>
#include <io.h>
#include <fcntl.h>
#include <cstring>
using namespace OpenXLSX;



// 获取 Windows 终端输出编码（控制台编码）
std::string get_terminal_encoding_windows() {
    // 获取控制台输出代码页（影响 printf/cout 输出）
    UINT code_page = GetConsoleOutputCP();
    
    // 常见代码页对应
    switch (code_page) {
        case 65001:  return "UTF-8";
        case 936:    return "GBK/CP936";
        case 1252:   return "Latin-1/CP1252";
        case 437:    return "IBM PC (US)/CP437";
        default:     return "Unknown (CodePage: " + std::to_string(code_page) + ")";
    }
}
std::string GBKToUTF8(const char* gbk_str) {
    if (!gbk_str) return "";
    // 先将 GBK 转换为 UTF-16
    int wlen = MultiByteToWideChar(CP_ACP, 0, gbk_str, -1, nullptr, 0);
    if (wlen <= 0) return "";
    std::vector<wchar_t> wbuf(wlen);
    MultiByteToWideChar(CP_ACP, 0, gbk_str, -1, wbuf.data(), wlen);
    // 再将 UTF-16 转换为 UTF-8
    int len = WideCharToMultiByte(CP_UTF8, 0, wbuf.data(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::vector<char> buf(len);
    WideCharToMultiByte(CP_UTF8, 0, wbuf.data(), -1, buf.data(), len, nullptr, nullptr);
    return std::string(buf.data());
}

std::string UTF8ToGBK(const std::string& utf8_str) {
    // 1. UTF-8 转 UTF-16
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return "";
    std::vector<wchar_t> wbuf(wlen);
    MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), -1, wbuf.data(), wlen);
    
    // 2. UTF-16 转 GBK
    int len = WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::vector<char> buf(len);
    WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, buf.data(), len, nullptr, nullptr);
    
    return std::string(buf.data());
}

void setup_utf8_console() {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    _setmode(_fileno(stdout), _O_TEXT);
    _setmode(_fileno(stdin),  _O_TEXT);
}
std::string utf8_from_wide(const std::wstring& wstr) {
    std::wstring_convert<std::codecvt_utf8<wchar_t>> conv;
    return conv.to_bytes(wstr);
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
// 辅助函数：判断字符串是否为 UTF-8 编码
bool is_utf8(const char* bytes) {
    if (!bytes) return false;
    size_t len = std::strlen(bytes);
    size_t i = 0;
    while (i < len) {
        if ((bytes[i] & 0x80) == 0x00) {  // 单字节：0xxxxxxx
            i++;
            continue;
        }
        // 多字节：第一个字节的高位连续 1 的数量表示字节数
        if ((bytes[i] & 0xE0) == 0xC0) {  // 双字节：110xxxxx
            if (i + 1 >= len || (bytes[i+1] & 0xC0) != 0x80) return false;
            i += 2;
        } else if ((bytes[i] & 0xF0) == 0xE0) {  // 三字节：1110xxxx
            if (i + 2 >= len || (bytes[i+1] & 0xC0) != 0x80 || (bytes[i+2] & 0xC0) != 0x80) return false;
            i += 3;
        } else if ((bytes[i] & 0xF8) == 0xF0) {  // 四字节：11110xxx
            if (i + 3 >= len || (bytes[i+1] & 0xC0) != 0x80 || (bytes[i+2] & 0xC0) != 0x80 || (bytes[i+3] & 0xC0) != 0x80) return false;
            i += 4;
        } else {  // 无效 UTF-8 字节
            return false;
        }
    }
    return true;
}

// 辅助函数：GBK 转 UTF-8（Windows 平台）
std::string gbk_to_utf8(const char* gbk_str) {
    if (!gbk_str) return "";
    // 1. GBK 转 UTF-16
    int wlen = MultiByteToWideChar(CP_ACP, 0, gbk_str, -1, nullptr, 0);
    if (wlen <= 0) return "";
    std::vector<wchar_t> wbuf(wlen);
    MultiByteToWideChar(CP_ACP, 0, gbk_str, -1, wbuf.data(), wlen);
    // 2. UTF-16 转 UTF-8
    int len = WideCharToMultiByte(CP_UTF8, 0, wbuf.data(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::vector<char> buf(len);
    WideCharToMultiByte(CP_UTF8, 0, wbuf.data(), -1, buf.data(), len, nullptr, nullptr);
    return std::string(buf.data());
}

// 处理命令行参数，转换为 UTF-8
std::vector<std::string> process_cmd_args(int argc, char* argv[]) {
    std::vector<std::string> utf8_args;
    for (int i = 0; i < argc; ++i) {
        std::string arg = argv[i];
#ifdef _WIN32
        // Windows 平台：默认按 GBK 处理，若检测为 UTF-8 则直接使用
        if (is_utf8(arg.c_str())) {
            utf8_args.push_back(arg);  // 已为 UTF-8
        } else {
            utf8_args.push_back(gbk_to_utf8(arg.c_str()));  // GBK 转 UTF-8
        }
#else
        // Linux/macOS 平台：默认按 UTF-8 处理（命令行默认编码）
        utf8_args.push_back(arg);
#endif
    }
    return utf8_args;
}
//int wmain(int argc, wchar_t* argv[]) {
	int main(int argc, char* argv[]) {
   if (argc<3){
   	//std::cout<<get_terminal_encoding_windows()<<std::endl;
   		// 根据终端编码输出中文
		std::string terminal_encoding = get_terminal_encoding_windows();
		if (terminal_encoding.find("UTF-8") != std::string::npos) {
		std::cout <<"terminal code is "<<terminal_encoding<<" UTF-8 中文测试" << std::endl;  // 输出 UTF-8
		} else if (terminal_encoding.find("GBK") != std::string::npos) {
		// 转换为 GBK 后输出（需使用 WideCharToMultiByte 等函数）
		std::cout <<"terminal code is "<<terminal_encoding<< "  "<<UTF8ToGBK("UTF-8 中文测试" ) << std::endl;
		}
   	return 0;
   }
    //setup_utf8_console();
		std::string filename=argv[1];
    if (!is_utf8(argv[1])) {
        filename = GBKToUTF8(argv[1]);
     }   
    std::string keyword  = argv[2];
    if (!is_utf8(argv[2])) {
     keyword  = GBKToUTF8(argv[2]);
     }
     std::string column   =  "A"; 
    if(argc>3&&!is_utf8(argv[3]))  {
      column   = GBKToUTF8(argv[3]);
    }
    

    std::cout << "搜索文件: " << filename << "，关键词: " << keyword <<"，列: " << column << std::endl;


    auto colOpt = column_letter_to_index(column);
    if (!colOpt) {
        std::cerr << "无效的列字母: " << column << "\n";
        return 2;
    }
    uint16_t colIndex = colOpt.value();

    try {
        XLDocument doc;
        doc.open(filename);
        auto wb = doc.workbook();
        auto sheets = wb.worksheetNames();

        std::vector<Match> allMatches;

        for (const auto& sheetName : sheets) {
            auto sheet = wb.worksheet(sheetName);
            auto matches = find_string_in_sheet(sheet, keyword, colIndex);
            allMatches.insert(allMatches.end(), matches.begin(), matches.end());
        }

        doc.close();

        std::cout << "搜索列: " << column << " (第 " << colIndex << " 列)\n";
        std::cout << "搜索内容: \"" << keyword << "\"\n";

        if (allMatches.empty()) {
            std::cout << "未在任何单元格中找到匹配内容。\n";
        } else {
            
            for (const auto& m : allMatches) {
                std::cout << "  工作表: " << m.sheetName \
                	<< "，单元格: " << m.cellAddress \
                	<< "] 内容: \"" << m.val << "\"\n";
            
            }
            std::cout << "匹配结果,共 ："<<allMatches.size()<<" 条记录\n";    		
        }

    } catch (const std::exception& ex) {
        std::cerr << "读取 Excel 文件失败: " << ex.what() << "\n";
        return 3;
    }

    return 0;
}
