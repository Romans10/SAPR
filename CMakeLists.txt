cmake_minimum_required(VERSION 3.2)

set (this SAPR)

project(${this})

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt5Core REQUIRED)
find_package(Qt5Gui REQUIRED)

include_directories(SYSTEM
	"${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/"
)

add_subdirectory("3rdparty")

add_executable(SAPR
    src/main.cpp
)

target_link_libraries(SAPR 
	Qt5::Core
	Qt5::GuiPrivate
	QXlsx
)

target_include_directories(SAPR
	PRIVATE
	"${SAPR_3RDPARTY_INCLUDE_DIRECTORIES}"
	"${CMAKE_CURRENT_SOURCE_DIR}"
)
