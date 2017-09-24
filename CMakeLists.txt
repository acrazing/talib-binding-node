cmake_minimum_required(VERSION 3.7)
project(talib_binding_node)

set(CMAKE_CXX_STANDARD 11)

set(SOURCE_FILES
        src/talib-binding.generated.cc)

add_executable(talib_binding_node ${SOURCE_FILES})