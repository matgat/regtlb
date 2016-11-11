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

# Link to Windows OLE libraries
win32: LIBS += -lole32 -loleaut32


#message($$QMAKESPEC)
win32-g++ {
    # Remove MinGW libstdc++-6.dll dependency and other dlls
    QMAKE_LFLAGS += -static -static-libgcc -static-libstdc++
    #  -fno-exceptions
    QMAKE_LFLAGS += -fno-rtti
}
