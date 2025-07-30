#include <OpenXLSX.hpp>
#include <iostream>
#include <string>
#include <vector>
#include <optional>
#include <cctype>
#include <windows.h>      // SetConsoleOutputCP
using namespace OpenXLSX;

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
    for (uint32_t row = 1; row <= 10000; ++row) {
        auto cell = sheet.cell(XLCellReference(row, col));
        if (cell.value().type() != XLValueType::Empty) {
            maxRow = row;
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

int main(int argc, char* argv[]) {
	SetConsoleOutputCP(CP_UTF8);
    if (argc < 3 || argc > 4) {
        std::cerr << "用法: " << argv[0] << " <xlsx文件路径> <搜索字符串> [列字母，默认A]\n";
        return 1;
    }

    std::string filename = argv[1];
    std::string searchStr = argv[2];
    std::string colLetter = (argc == 4) ? argv[3] : "A";

    auto colOpt = column_letter_to_index(colLetter);
    if (!colOpt) {
        std::cerr << "无效的列字母: " << colLetter << "\n";
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
            auto matches = find_string_in_sheet(sheet, searchStr, colIndex);
            allMatches.insert(allMatches.end(), matches.begin(), matches.end());
        }

        doc.close();

        std::cout << "搜索列: " << colLetter << " (第 " << colIndex << " 列)\n";
        std::cout << "搜索内容: \"" << searchStr << "\"\n";

        if (allMatches.empty()) {
            std::cout << "未在任何单元格中找到匹配内容。\n";
        } else {
            std::cout << "匹配结果：\n";
            for (const auto& m : allMatches) {
                std::cout << "  工作表: " << m.sheetName \
                	<< "，单元格: " << m.cellAddress \
                	<< "] 内容: \"" << m.val << "\"\n";
            }
        }

    } catch (const std::exception& ex) {
        std::cerr << "读取 Excel 文件失败: " << ex.what() << "\n";
        return 3;
    }

    return 0;
}
