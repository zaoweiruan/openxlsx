#include <OpenXLSX.hpp>
#include <boost/algorithm/string.hpp>
#include <iostream>
#include <vector>
#include <optional>
#include <cctype>

using namespace OpenXLSX;

struct Match {
    std::string sheetName;
    std::string cellAddress;
    	    Match(const std::string& sheet, const std::string& address)
        : sheetName(sheet), cellAddress(address) {}
};

// 字母列名 -> 数字列号 (A=1, B=2, ..., Z=26, AA=27, AB=28...)
uint16_t column_name_to_index(const std::string& colName) {
    uint16_t result = 0;
    for (char c : colName) {
        if (std::isalpha(c)) {
            result = result * 26 + (std::toupper(c) - 'A' + 1);
        }
    }
    return result;
}

// 检测某列最大非空行
uint32_t detect_max_row(XLWorksheet& sheet, uint16_t col) {
    uint32_t maxRow = 1;
    for (uint32_t row = 1; row <= 10000; ++row) {
        auto cell = sheet.cell(XLCellReference(row, col));
        if (cell.value().type() != XLValueType::Empty)
      {
            maxRow = row;
        }
    }
    return maxRow;
}

std::vector<Match> find_string_in_sheet(XLWorksheet& sheet, const std::string& target, std::optional<uint16_t> onlyColumn = std::nullopt) {
    std::vector<Match> matches;

    uint16_t startCol = 1;
    uint16_t endCol = 1000;
    uint32_t maxRow = 10000;

    if (onlyColumn.has_value()) {
        startCol = endCol = onlyColumn.value();
        maxRow = detect_max_row(sheet, startCol);
    }

    for (uint32_t row = 1; row <= maxRow; ++row) {
        for (uint16_t col = startCol; col <= endCol; ++col) {
            XLCell cell = sheet.cell(XLCellReference(row, col));
            if (cell.value().type() == XLValueType::Empty) continue;

            if (cell.value().type() == XLValueType::String) {
                std::string val = cell.value().get<std::string>();
                if (boost::iequals(val, target)) {
                   
                    matches.emplace_back(sheet.name(), cell.cellReference().address());

                }
            }
        }
    }

    return matches;
}

int main(int argc, char* argv[]) {
    if (argc < 3) {
        std::cerr << "Usage: " << argv[0] << " <ExcelFile.xlsx> <SearchString> [Column]\n";
        return 1;
    }

    std::string fileName = argv[1];
    std::string search = argv[2];
    std::optional<uint16_t> onlyColumn = std::nullopt;

    if (argc >= 4) {
        onlyColumn = column_name_to_index(argv[3]);
        if (onlyColumn == 0) {
            std::cerr << "Invalid column name: " << argv[3] << "\n";
            return 1;
        }
    }

    try {
        XLDocument doc;
        doc.open(fileName);
        auto workbook = doc.workbook();

        std::vector<Match> allMatches;

        for (const auto& sheetName : workbook.worksheetNames()) {
            XLWorksheet sheet = workbook.worksheet(sheetName);
            auto matches = find_string_in_sheet(sheet, search, onlyColumn);
            allMatches.insert(allMatches.end(), matches.begin(), matches.end());
        }

        doc.close();

        if (allMatches.empty()) {
            std::cout << "No matches found for \"" << search << "\" in \"" << fileName << "\".\n";
        } else {
            std::cout << "Found \"" << search << "\" in \"" << fileName << "\" at:\n";
            for (const auto& match : allMatches) {
                std::cout << "  - Sheet: " << match.sheetName
                          << ", Cell: " << match.cellAddress << "\n";
            }
        }
    } catch (const std::exception& ex) {
        std::cerr << "Error: " << ex.what() << "\n";
        return 1;
    }

    return 0;
}
