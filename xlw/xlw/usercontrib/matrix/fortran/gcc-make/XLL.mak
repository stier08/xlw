#Describes the xll 
LIBRARY = Fortran
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
LIBS=xlw-gcc-s-gd-5_0_1f0
else
LIBS=xlw-gcc-s-5_0_1f0
endif 

#Describes the Compiler details
INCLUDE_DIR =../common_source  $(subst $(strip \),/,$(XLW))/xlw/include
CXXFLAGS = 


#The source
SRC_DIR=../common_source
LIBSRC = Test.cpp
		
MAKEDIR = $(subst $(strip \),/,$(XLW))/xlw/make
include $(MAKEDIR)/make.rules
include $(MAKEDIR)/make.targets