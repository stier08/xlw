#Describes the xll 
LIBRARY = uBLAS
LIBTYPE=SHARE
LIBPREFIX=
EXT_SHARE=xll

#Describes the Linker details
ifeq ($(PLATFORM), x64)
LIBDIRS = $(subst $(strip \),/,$(XLW))/xlw/lib/x64
else
LIBDIRS = $(subst $(strip \),/,$(XLW))/xlw/lib
endif
ifeq ($(BUILD),DEBUG)
LIBS=xlw-gcc-s-gd-5DEV
else
LIBS=xlw-gcc-s-5DEV
endif 

#Describes the Compiler details
INCLUDE_DIR =../common_source  $(subst $(strip \),/,$(XLW))/xlw/include
CXXFLAGS = 


#The source
SRC_DIR=../common_source
LIBSRC = Test.cpp \
		 xlwTest.cpp 
		
MAKEDIR = $(subst $(strip \),/,$(XLW))/xlw/make
include $(MAKEDIR)/make.rules
include $(MAKEDIR)/make.targets