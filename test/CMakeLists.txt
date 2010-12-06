# declare additional include directories
INCLUDE_DIRECTORIES(${dblib_SOURCE_DIR}/include )

# declare directories containing used libraries
LINK_DIRECTORIES(${dblib_BINARY_DIR} )

# declare executables
ADD_EXECUTABLE(dbtest dbtest.cpp)

# make sure executables are linked to the corresponding libraries
FOREACH(PROGRAM dbtest)
  TARGET_LINK_LIBRARIES(${PROGRAM} dblib)
ENDFOREACH(PROGRAM)
