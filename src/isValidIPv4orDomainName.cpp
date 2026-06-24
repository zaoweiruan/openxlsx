/*
 * isValidIPv4orDomainName.cpp
 *
 *  Created on: 2026年2月12日
 *      Author: dsm
 */



#include <string>
#include <cctype>
#include <cwctype>
#include <optional>
#include <algorithm>
#include <windows.h>     // 提供 SetConsoleOutputCP、SetConsoleCP、CP_UTF8 等
#include <fcntl.h>       // _setmode、_fileno
#include <io.h>          // _fileno、_setmode 的另一個必要標頭
// 判断是否全部是 ASCII 字符（域名和IPv4通常都要求ASCII）
inline bool is_ascii(const std::wstring& s) {
    return std::all_of(s.begin(), s.end(), [](wchar_t c) { return c <= 127; });
}

// 辅助：把 wstring 安全转为 string（假设已经是ASCII）
inline std::string to_narrow(const std::wstring& ws) {
    return std::string(ws.begin(), ws.end());
}


std::string to_narrow_strict(const std::wstring& ws) {
    std::string result;
    result.reserve(ws.size());

    for (wchar_t wc : ws) {
        if (wc > 127 || wc < 0) {
           return "";   // 看你業務要怎麼處理
        }
        result.push_back(static_cast<char>(wc));
    }
    return result;
}
// 判断是否像合法的 IPv4 地址（最严格模式）
bool is_valid_ipv4(const std::wstring& s) {
    if (s.empty() || s.size() > 15) return false;           // 最长 "255.255.255.255"
    if (!is_ascii(s)) return false;

    //std::string str = to_narrow_strict(s);
    std::string str = to_narrow(s);
    if (str.empty()) return false; // 转换失败（非ASCII）

    int dot_count = 0;
    int num_start = 0;

    for (size_t i = 0; i <= str.size(); ++i) {
        if (i == str.size() || str[i] == '.') {
            if (i == (size_t)num_start) return false;               // 连续两个点 或 结尾是点
            if (i - num_start > 3) return false;            // 超过3位

            std::string part = str.substr(num_start, i - num_start);
            if (part.size() > 1 && part[0] == '0') return false;  // 禁止前导0（除了0本身）

            int num = std::stoi(part);
            if (num < 0 || num > 255) return false;

            dot_count++;
            num_start = i + 1;
        }
        else if (!std::isdigit(static_cast<unsigned char>(str[i]))) {
            return false;
        }
    }

//    return dot_count == 4 && num_start <= str.size();
    return dot_count == 4;
}

// 判断是否像域名（比较宽松的常用判断）
bool looks_like_domain(const std::wstring& s) {
    if (s.empty() || s.size() > 253) return false;          // RFC 1035 域名最大253字节
    if (!is_ascii(s)) return false;

    std::string str = to_narrow(s);

    // 最基本的非法字符过滤
    static const std::string allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.";
    if (str.find_first_not_of(allowed) != std::string::npos) {
        return false;
    }

    // 不能以点开头或结尾
    if (str.front() == '.' || str.back() == '.') {
        return false;
    }

    // 不能有连续的点
    if (str.find("..") != std::string::npos) {
        return false;
    }

    // 每个 label 不能超过63字节，且不能以-开头或结尾
    size_t pos = 0;
    while (pos < str.size()) {
        size_t next = str.find('.', pos);
        if (next == std::string::npos) next = str.size();

        size_t len = next - pos;
        if (len == 0 || len > 63) return false;
        if (str[pos] == '-' || str[next - 1] == '-') return false;

        pos = next + 1;
    }

    // 至少要有一个点（最简域名 a.b）
    return str.find('.') != std::string::npos;
}

// 综合判断：是 IPv4 还是 域名
enum class AddressType {
    Invalid,
    IPv4,
    Domain
};

AddressType classify_address(const std::wstring& s) {
    if (is_valid_ipv4(s)) {
        return AddressType::IPv4;
    }
    if (looks_like_domain(s)) {
        return AddressType::Domain;
    }
    return AddressType::Invalid;
}

// 使用示例
#include <iostream>

int wmain() {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    _setmode(_fileno(stdout), _O_U8TEXT);
    std::wstring tests[] = {
        L"192.168.1.1",
        L"255.255.255.255",
        L"10.0.0.01",           // 前导0 → 非法
        L"256.1.2.3",           // 256非法
        L"1.2.3.4.5",           // 多一段
        L"www.example.com",
        L"api-v2.internal.local",
        L"中文.域名.com",        // 非ASCII → 非法（按传统域名算）
        L"localhost",
        L"127.0.0.1",
        L"1.1.1",
        L"1.1.1.1."
    };

    for (const auto& t : tests) {
        auto type = classify_address(t);
        std::wcout << t << L" → ";

        if (type == AddressType::IPv4)      std::wcout << L"IPv4\n";
        else if (type == AddressType::Domain) std::wcout << L"Domain\n";
        else                                std::wcout << L"Invalid\n";
    }

    return 0;
}
