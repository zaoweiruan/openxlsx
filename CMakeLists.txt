cmake_minimum_required (VERSION 3.30)
set(CMAKE_VERBOSE_MAKEFILEON ON)

project (search_excel)
#set(CMAKE_BUILD_TYPE RelWithDebInfo)
#set(CMAKE_BUILD_TYPE Release)
set(CMAKE_BUILD_TYPE Debug)

add_library(boost_algorithm_lib INTERFACE)
# 添加头文件所在目录（相对路径或绝对路径）
target_include_directories(boost_algorithm_lib INTERFACE
    "D:/boost_1_88_0"  # 库的头文件所在目录
)
set(OpenXLSX_DIR "E:/OpenXLSX/lib/cmake/OpenXLSX")
find_package(OpenXLSX REQUIRED)
find_package(unofficial-sqlite3 CONFIG REQUIRED)

#add_library(openxlsx_lib STATIC IMPORTE)
# 添加头文件所在目录（相对路径或绝对路径）
#target_include_directories(boost_algorithm_lib INTERFACE
#    "E:/OpenXLSX/include/OpenXLSX"  # 库的头文件所在目录
#)

#set_target_properties(openxlsx_lib
#             PROPERTY IMPORTED_LOCATION "E:/OpenXLSX/lib/libOpenXLSX.a") 
# Boost 静态构建（禁用动态 Boost.DLL）
#add_definitions(-DBOOST_ALL_NO_LIB)
#add_definitions(-DBOOST_ALL_STATIC_LINK)

#find_package(Boost  CONFIG REQUIRED  json )

add_executable(search_excel)
add_executable(isValidIPv4orDomainName)
add_executable(validProxy)


set_target_properties(search_excel PROPERTIES SOURCES "${PROJECT_SOURCE_DIR}/src/listern_clipboard_1.cpp")
set_target_properties(isValidIPv4orDomainName PROPERTIES SOURCES "${PROJECT_SOURCE_DIR}/src/isValidIPv4orDomainName.cpp")
set_target_properties(validProxy PROPERTIES SOURCES "${PROJECT_SOURCE_DIR}/src/validProxy.cpp")


set_target_properties(search_excel PROPERTIES OUTPUT_NAME "search_book_service")

set_target_properties(search_excel PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${PROJECT_SOURCE_DIR}/bin")
set_target_properties(isValidIPv4orDomainName PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${PROJECT_SOURCE_DIR}/bin")
set_target_properties(validProxy PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${PROJECT_SOURCE_DIR}/bin")

target_link_libraries(search_excel  OpenXLSX::OpenXLSX boost_algorithm_lib Comctl32 Shlwapi user32 ws2_32)
target_link_options(isValidIPv4orDomainName PRIVATE -mconsole -municode)
target_link_libraries(validProxy PRIVATE unofficial::sqlite3::sqlite3 Comctl32  ws2_32)
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -lstdc++ -std=c++17 -g -Wall  -O2 -mconsole -municode ")
#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -lstdc++ -std=c++17 -g -Wall  -O2 -mwindows -municode ")
