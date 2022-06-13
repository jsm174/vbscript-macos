cmake_minimum_required(VERSION 3.22)

project(vbstest)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_C_STANDARD 99)

set(CMAKE_C_FLAGS "-Wno-pragma-pack -Wno-format")

add_compile_definitions(
   _WIN64
   __WINESRC__
   WINE_UNICODE_NATIVE
)

add_executable(vbstest
   wine/dlls/oleaut32/oleaut.c
   wine/dlls/oleaut32/variant.c
   wine/dlls/oleaut32/varformat.c
   wine/dlls/oleaut32/vartype.c
   wine/dlls/oleaut32/safearray.c

   wine/dlls/vbscript/compile.c
   wine/dlls/vbscript/global.c
   wine/dlls/vbscript/interp.c
   wine/dlls/vbscript/lex.c
   wine/dlls/vbscript/regexp.c
   wine/dlls/vbscript/utils.c
   wine/dlls/vbscript/vbdisp.c
   wine/dlls/vbscript/vbregexp.c
   wine/dlls/vbscript/vbscript.c
   wine/dlls/vbscript/vbscript_main.c
   wine/dlls/vbscript/parser.tab.c

   wine/dlls/kernelbase/locale.c

   wine_stubs.c

   main.c
)

target_include_directories(vbstest PUBLIC
   wine/include
   wine/dlls/vbscript
   wine/dlls/oleaut32
   wine/dlls/kernelbase
)

target_link_libraries(vbstest PUBLIC
)

set_target_properties(vbstest PROPERTIES
   RUNTIME_OUTPUT_NAME "vbstest"
   SKIP_BUILD_RPATH TRUE
   LINK_FLAGS "-Wl,-rpath,@executable_path"
)