cmake_minimum_required(VERSION 3.2)

set(INCLUDE_DIRS "${CMAKE_CURRENT_SOURCE_DIR}")

add_compile_options(${CMAKE_CXX_FLAGS})

add_subdirectory("QXlsx")

set(SAPR_3RDPARTY_INCLUDE_DIRECTORIES "${INCLUDE_DIRS}" PARENT_SCOPE)
