cmake_minimum_required(VERSION 3.16)

project(tarfldr)

set(SRC_FILES tarfldr.cpp
	tarfldr.def)

add_library(tarfldr SHARED ${SRC_FILES})

set_target_properties(tarfldr PROPERTIES PREFIX "")

# target_link_libraries(explorer advapi32 gdi32 user32 comctl32 ole32 oleaut32 shell32 shlwapi shdocvw version uxtheme msvcrt kernel32 ntdll)

target_compile_options(tarfldr PUBLIC -gcodeview)
target_link_options(tarfldr PUBLIC -gcodeview)
