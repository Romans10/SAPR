# CMakeLists.txt for QXlsx Library

cmake_minimum_required(VERSION 3.14)

project(QXlsx)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt5Core REQUIRED)
find_package(Qt5Gui REQUIRED)

file(GLOB QXLSX_CPP "${CMAKE_CURRENT_SOURCE_DIR}/*.cpp")
file(GLOB QXLSX_H "${CMAKE_CURRENT_SOURCE_DIR}/*.h")

add_library(QXlsx STATIC ${QXLSX_CPP} ${QXLSX_H})
target_link_libraries(QXlsx PUBLIC Qt5::Core Qt5::GuiPrivate)
target_include_directories(QXlsx PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})
