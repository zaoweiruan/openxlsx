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
	static UINT code_page = GetConsoleOutputCP(); //显示字符集
	static UINT acp = GetACP(); //输入字符集
#define MY_ASSERT(condition, message) \
    do { \
        if (!(condition)) { \
            std::cerr << "Assertion failed: " << message \
                      << ", file " << __FILE__ \
                      << ", line " << __LINE__ << std::endl; \
            abort(); \
        } \
    } while (0)

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

//int wmain(int argc, wchar_t* argv[]) {
	int main(int argc, char* argv[]) {
		
		UINT code_page = GetConsoleOutputCP(); //显示字符集
		UINT acp = GetACP(); //输入字符集
try {
		ASSERT_MSG(argc>2, "arg must >=2, argc =%d",argc);
		std::string filename=CharsetConvert(argv[1],acp,code_page);
			
		std::string keyword=CharsetConvert(argv[2],acp,CP_UTF8);
//		std::string keyword=argv[2];
		std::string keyword_str=CharsetConvert(argv[2],acp,code_page);
		std::string 	column=(argc<4?"A":argv[3]);
			
		std::string outstr=CharsetConvert("搜索文件: ",CP_UTF8,code_page)\
			+filename+CharsetConvert("，关键词: ",CP_UTF8,code_page)\
				+keyword_str\
				+CharsetConvert("，列: " ,CP_UTF8,code_page)\
					+column;

    std::cout << outstr << std::endl;


    auto colOpt = column_letter_to_index(column);
    if (!colOpt) {
        std::cerr << CharsetConvert("无效的列字母: ",CP_UTF8,code_page) << column << "\n";
        return 2;
    }
    uint16_t colIndex = colOpt.value();

    
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

        std::cout << CharsetConvert("搜索列: ",CP_UTF8,code_page) << column << CharsetConvert(" (第 ",CP_UTF8,code_page) << colIndex << CharsetConvert(" 列)",CP_UTF8,code_page)<<std::endl;
        std::cout << CharsetConvert("搜索内容: \"",CP_UTF8,code_page) << keyword_str << "\"\n";

        if (allMatches.empty()) {
            std::cout << CharsetConvert("未在任何单元格中找到匹配内容",CP_UTF8,code_page)<<std::endl;
        } else {
            
            for (const auto& m : allMatches) {
                std::cout << CharsetConvert("  工作表: ",CP_UTF8,code_page) << m.sheetName \
                	<< CharsetConvert("，单元格: ",CP_UTF8,code_page) << CharsetConvert(std::string(m.cellAddress).c_str(),CP_UTF8,code_page) \
                	<< CharsetConvert("] 内容: \"",CP_UTF8,code_page) << CharsetConvert(std::string(m.val).c_str(),CP_UTF8,code_page) << "\"\n";
            
            }
            std::cout <<  CharsetConvert("匹配结果,共 ：",CP_UTF8,code_page)<<allMatches.size()<<CharsetConvert(" 条记录",CP_UTF8,code_page)<<std::endl;		
        }

    } catch (const std::exception& ex) {
        std::cerr << CharsetConvert("读取 Excel 文件失败: ",CP_UTF8,code_page)<<std::endl;
        return 3;
    }

    return 0;
}
