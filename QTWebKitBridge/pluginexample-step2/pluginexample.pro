#-------------------------------------------------
#
# Project created by QtCreator 2010-08-19T09:53:29
#
#-------------------------------------------------

QT += core \
      gui \
      webkit

TARGET = pluginexample
TEMPLATE = app

Debug:DEFINES += DEBUG

SOURCES += src/main.cpp\
        src/pluginexample.cpp \
    src/webpluginfactory.cpp \
    src/examplewidget.cpp \
    src/descriptionhelper.cpp

HEADERS  += src/pluginexample.h \
    src/webpluginfactory.h \
    src/examplewidget.h \
    src/descriptionhelper.h

RESOURCES += \
    pluginexample.qrc

FORMS += \
    src/examplewidget.ui
