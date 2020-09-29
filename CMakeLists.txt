cmake_minimum_required(VERSION 3.15)

project(tarfldr VERSION 1.0.0)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)

add_definitions(-DUNICODE -D_UNICODE)

set(SRC_FILES tarfldr.cpp
	folder.cpp
	enum.cpp
	item.cpp
	menu.cpp
	stream.cpp
	${CMAKE_CURRENT_BINARY_DIR}/tarfldr.rc
	tarfldr.def)

configure_file(tarfldr.rc.in tarfldr.rc)

add_library(tarfldr SHARED ${SRC_FILES})

find_package(fmt REQUIRED)
find_package(LibArchive REQUIRED)

target_include_directories(tarfldr PUBLIC "${LibArchive_INCLUDE_DIR}")

set_target_properties(tarfldr PROPERTIES PREFIX "")

target_link_libraries(tarfldr comctl32 shlwapi)
target_link_libraries(tarfldr fmt::fmt-header-only)
target_link_libraries(tarfldr libarchive.a libbz2.a libxml2.a libz.a liblzma.a)
target_link_libraries(tarfldr ws2_32)
target_link_options(tarfldr PUBLIC -static-libgcc)

if(CMAKE_BUILD_TYPE STREQUAL "Debug")
	add_definitions(-D_DEBUG)
endif()

target_compile_options(tarfldr PUBLIC -gcodeview)
target_link_options(tarfldr PUBLIC -gcodeview)
