#Describes the xll 
BUILD=DEBUG
LIBRARY = LoggerDemo
LIBTYPE=SHARE
LIBPREFIX=
EXT_SHARE=xll

#Describes the Linker details
LIBDIRS = ../../../lib
ifeq ($(BUILD),DEBUG)
LIBS=xlw-gcc-s-gd-5DEV
else
LIBS=xlw-gcc-s-5DEV
endif 

#Describes the Compiler details
INCLUDE_DIR =../common_source  ../../../include
CXXFLAGS = 


#The source
SRC_DIR=../common_source
LIBSRC = xlwLogger.cpp \
		 Test.cpp \
		 xlwTest.cpp 
		
MAKEDIR = ../../../make
include $(MAKEDIR)/make.rules
include $(MAKEDIR)/make.targets