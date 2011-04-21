include ./XLWVERSION.mak

ifeq ($(BUILD),DEBUG)
LIBRARY = xlw-gcc-s-gd-$(XLWVERSION)
endif
ifeq ($(BUILD),RELEASE)
LIBRARY = xlw-gcc-s-$(XLWVERSION)
endif


INSTALL_DIR = ../../lib



LIBTYPE=STATIC



INCLUDE_DIR = ../../include
CXXFLAGS = -DBUILDING_DLL=1  -fexceptions


SRC_DIR = ../../src
LIBSRC = ArgList.cpp \
         CellMatrix.cpp \
         Dispatcher.cpp \
         DoubleOrNothing.cpp \
         FileConverter.cpp \
         HiResTimer.cpp \
         NCmatrices.cpp \
         MyContainers.cpp \
         PascalStringConversions.cpp \
         PathUpdater.cpp \
         TempMemory.cpp \
         Win32StreamBuf.cpp \
         XlFunctionRegistration.cpp \
         XlfAbstractCmdDesc.cpp \
         XlfArgDesc.cpp \
         XlfArgDescList.cpp \
         XlfCmdDesc.cpp \
         XlfExcel.cpp \
         XlfFuncDesc.cpp \
         XlfOperImpl.cpp \
         XlfOperProperties.cpp \
         XlfRef.cpp \
         XlfStr.cpp \
         xlarray.cpp \
         xlcall.cpp \
         XlOpenClose.cpp   

MAKEDIR = ../../make
include $(MAKEDIR)/make.rules
include $(MAKEDIR)/make.targets
