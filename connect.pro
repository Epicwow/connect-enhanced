#-------------------------------------------------
#
# Project created by QtCreator 2014-11-27T21:08:50
#
#-------------------------------------------------

QT       += core gui
QT       += sql
QT       += network

include(3rdparty/qtxlsx/src/xlsx/qtxlsx.pri)

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = connect
TEMPLATE = app

QMAKE_MAC_SDK = macosx10.11


SOURCES += main.cpp\
        connect.cpp

HEADERS  += connect.h

FORMS    += connect.ui

RESOURCES += \
    img.qrc
