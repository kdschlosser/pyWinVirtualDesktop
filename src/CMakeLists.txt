cmake_minimum_required(VERSION 3.9)
project(untitled1)

set(CMAKE_CXX_STANDARD 14)

add_library(stdafx stdafx.cpp stdafx.h)
add_library(dllmain dllmain.cpp dllmain.h)
include_directories(C:\\Stackless27\\include)
include_directories(C:\\Stackless27\\include\\stackless)