cmake_minimum_required (VERSION 2.6)

PROJECT(dblib)

INCLUDE_DIRECTORIES(${dblib_SOURCE_DIR}/include )

SUBDIRS(libsrc test)

#INSTALL_TARGETS(${INSTALL_LIBDIR} dcmdata)
