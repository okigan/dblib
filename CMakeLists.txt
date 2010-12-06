cmake_minimum_required (VERSION 2.6)

PROJECT(dblib)

option (USE_SQLITE_2 "Use SQLite version 2" OFF) 
option (USE_SQLITE_3 "Use SQLite version 3" OFF) 

INCLUDE_DIRECTORIES(${dblib_SOURCE_DIR}/include )

SUBDIRS(libsrc test)

