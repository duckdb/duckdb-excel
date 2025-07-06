cmake_minimum_required(VERSION 2.8.12...3.29)
set(TARGET_NAME excel)
set(EXTENSION_NAME ${TARGET_NAME}_extension)
set(LOADABLE_EXTENSION_NAME ${TARGET_NAME}_loadable_extension)

project(ExcelExtension)

if (EMSCRIPTEN)
  # _LINKED_LIBS influences only Wasm compilation
  # it's unclear why this is needed, but avoid globals to be incorrectly not linked-in
  # this solves by basically re-linking (at the moment the Wasm binary is actually produced)
  set (DUCKDB_EXTENSION_EXCEL_LINKED_LIBS "../../vcpkg_installed/wasm32-emscripten/lib/lib*.a")
endif()

# Dependencies from VCPKG
find_package(EXPAT REQUIRED)
find_package(ZLIB REQUIRED)
find_package(minizip-ng CONFIG REQUIRED)

include_directories(src/excel/numformat/include)
include_directories(src/excel/include)
add_subdirectory(src/excel/numformat)

set(EXTENSION_SOURCES src/excel/excel_extension.cpp src/excel/xlsx/zip_file.cpp
                      src/excel/xlsx/read_xlsx.cpp src/excel/xlsx/copy_xlsx.cpp)

build_static_extension(${TARGET_NAME} ${EXTENSION_SOURCES}
                       ${NUMFORMAT_OBJECT_FILES})
set(PARAMETERS "-warnings")
build_loadable_extension(${TARGET_NAME} ${PARAMETERS} ${EXTENSION_SOURCES}
                         ${NUMFORMAT_OBJECT_FILES})

target_link_libraries(${EXTENSION_NAME} EXPAT::EXPAT MINIZIP::minizip-ng
                      ZLIB::ZLIB)
target_link_libraries(${LOADABLE_EXTENSION_NAME} EXPAT::EXPAT
                      MINIZIP::minizip-ng ZLIB::ZLIB)

install(
  TARGETS ${EXTENSION_NAME}
  EXPORT "${DUCKDB_EXPORT_SET}"
  LIBRARY DESTINATION "${INSTALL_LIB_DIR}"
  ARCHIVE DESTINATION "${INSTALL_LIB_DIR}")
