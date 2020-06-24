cmake_minimum_required(VERSION 3.16)

project(tarfldr)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)

set(SRC_FILES tarfldr.cpp
	folder.cpp
	tarfldr.def)

add_library(tarfldr SHARED ${SRC_FILES})

find_package(fmt REQUIRED)

set_target_properties(tarfldr PROPERTIES PREFIX "")

target_link_libraries(tarfldr comctl32)
target_link_libraries(tarfldr fmt::fmt-header-only)
target_link_options(tarfldr PUBLIC -static-libgcc)

target_compile_options(tarfldr PUBLIC -gcodeview)
target_link_options(tarfldr PUBLIC -gcodeview)
