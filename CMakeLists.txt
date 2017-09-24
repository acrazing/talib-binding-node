cmake_minimum_required(VERSION 3.7)
project(talib_binding_node)

set(CMAKE_CXX_STANDARD 11)

include_directories(node/deps/v8/include)
include_directories(node/src)
include_directories(node_modules/nan)
include_directories(ta-lib/c/include)

set(SOURCE_FILES
        src/talib-binding.generated.cc)

add_executable(talib_binding_node ${SOURCE_FILES})