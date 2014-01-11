#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QtGui/QMainWindow>
#include <QtCore/QDir>
#include <QtCore/QUrl>

#include <QtGui/QWidget>
#include <QtGui/QVBoxLayout>
#include <QtGui/QFrame>
#include <QtGui/QDesktopServices>

#include <QtWebKit/QWebView>
#include <QtWebKit/QWebPage>
#include <QtWebKit/QWebFrame>
#include <QtWebKit/QWebSettings>

#include <QtCore/QDebug>

#include "webpluginfactory.h"
#include "examplewidget.h"

class QWebView;

class PluginExample : public QMainWindow
{
    Q_OBJECT

public:
    PluginExample(QWidget *parent = 0);
    ~PluginExample();



private:
    QWebView* m_webView;

    QWebView* createWebView();
};


#endif // MAINWINDOW_H
