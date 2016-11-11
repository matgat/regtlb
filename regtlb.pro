TEMPLATE = app
CONFIG += c++14
CONFIG += console
CONFIG -= app_bundle
CONFIG -= qt
#CONFIG += shared

#include(deployment.pri)
#qtcAddDeployment()

#message(Building $$TARGET)

SOURCES += main.cpp \
    unt_OleUts.cpp


HEADERS += \
    unt_OleUts.h

#DEFINES += "UNICODE"

win32 {
CONFIG += embed_manifest_exe
#QMAKE_LFLAGS_WINDOWS += /MANIFESTUAC:level=\'requireAdministrator\'
#QMAKE_LFLAGS += /MANIFESTUAC:\"level=\'requireAdministrator\' uiAccess=\'false\'\"
# Link to Windows OLE libraries
LIBS += -lole32 -loleaut32
}



#message($$QMAKESPEC)
win32-g++ {
    # Remove MinGW libstdc++-6.dll dependency and other dlls
    QMAKE_LFLAGS += -static -static-libgcc -static-libstdc++
    #  -fno-exceptions
    QMAKE_LFLAGS += -fno-rtti
}
