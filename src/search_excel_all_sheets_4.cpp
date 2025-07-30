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
std::wstring MlutibytesConvertWildebytes(const std::string& str) {
	
	//ansi多字节字符集转换宽字符集unicode
    // 1. 输入参数 转 unicode
    UINT acp = GetACP(); //输入字符集
    int wlen = MultiByteToWideChar(acp, 0, str.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return L"";
    std::vector<wchar_t> wbuf(wlen);
    MultiByteToWideChar(acp, 0, str.c_str(), -1, wbuf.data(), wlen);
    return std::wstring(wbuf.data());
   
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
		auto cell = sheet.cell(1,1);
		std::string val;
    for (uint32_t row = 1; row <= maxRow; ++row) {
        cell = sheet.cell(XLCellReference(row, col));
//        auto cell = sheet.cell(XLCellReference(row, col));
       // if (cell.value().type() == XLValueType::Empty) continue;

//        std::string val = cell.value().get<std::string>();
        val = cell.value().get<std::string>();
        if (val.find(target) != std::string::npos) {
            matches.emplace_back(sheet.name(), cell.cellReference().address(),val);
        }
    }

    return matches;
}

//int wmain(int argc, wchar_t* argv[]) {
	int main(int argc, char* argv[]) {
		
//		UINT code_page = GetConsoleOutputCP(); //显示字符集
		
		_setmode(_fileno(stdout), _O_U16TEXT); //输出宽字符集
try {
		ASSERT_MSG(argc>2, "arg must >=2, argc =%d",argc);
		
//		std::string filename=CharsetConvert(argv[1],acp,code_page);

		std::string filename=argv[1];
			UINT acp = GetACP(); //输入字符集	
					//文件默认以utf8编码需要将输入搜索关键字转为utf8		
		std::string keyword=CharsetConvert(argv[2],acp,CP_UTF8); 

		std::string 	column=(argc<4?"A":argv[3]);
			
		std::wstring column_wstr=MlutibytesConvertWildebytes(column);	
		std::wstring keyword_wstr=MlutibytesConvertWildebytes(std::string(argv[2]));
		std::wstring filename_wstr=MlutibytesConvertWildebytes(filename);
			
		std::wstring outstr=L"搜索文件: "\
												+filename_wstr\
												+L"，关键词: "\
												+keyword_wstr\
												+L"，列: " \
												+column_wstr;

    std::wcout << outstr << std::endl;


    auto colOpt = column_letter_to_index(column);
    if (!colOpt) {
    	_setmode(_fileno(stdout),_O_TEXT); //输出ansi字符集
    	
        std::cerr << "无效的列字母: " << column << "\n";
        ASSERT_MSG(colOpt<1, "colOpt error =%d",colOpt);		
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

        std::wcout << L"搜索列: "<< column_wstr << L" (第 "<< colIndex << L" 列)"<<std::endl;
        std::wcout << L"搜索内容: \"" << keyword_wstr << L"\"\n";

        if (allMatches.empty()) {
            std::wcout << L"未在任何单元格中找到匹配内容"<<std::endl;
        } else {
            
            for (const auto& m : allMatches) {
                std::wcout <<L"  工作表: " << MlutibytesConvertWildebytes(m.sheetName) \
                	<< L"，单元格: " << MlutibytesConvertWildebytes(std::string(m.cellAddress)) \
                	<< L"] 内容: \""<< MlutibytesConvertWildebytes(std::string(m.val)) << L"\"\n";
            
            }
            std::wcout << L"匹配结果,共 ："<<allMatches.size()<<L" 条记录"<<std::endl;		
        }

    } catch (const std::exception& ex) {
    	  _setmode(_fileno(stdout),_O_TEXT); //输出ansi字符集
        std::cerr <<"读取el 文件失败: "<<ex.what()<<std::endl;
  
        return 3;
    }
	 _setmode(_fileno(stdout),_O_TEXT); //输出ansi字符集
    return 0;
}
