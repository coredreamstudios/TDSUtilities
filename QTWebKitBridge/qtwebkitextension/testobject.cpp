#include <qdebug.h>
#include <qwebview.h>
#include <qwebframe.h>
#include <QMessageBox>

#include "testobject.h"
#include "dialog.h"

MyApi::MyApi( QObject *parent ) : QObject( parent )
{
    Dialog d;
    d.exec();
}

void MyApi::setWebView( QWebView *view )
{
    QWebPage *page = view->page();
    frame = page->mainFrame();

    attachObject();
    connect( frame, SIGNAL(javaScriptWindowObjectCleared()), this, SLOT(attachObject()) );
}

void MyApi::attachObject()
{
    frame->addToJavaScriptWindowObject( QString("MyApi"), this );
}

void MyApi::doSomething( const QString &param )
{
    qDebug() << "doSomething called with parameter " << param;

    QMessageBox m;
    m.setText("doSomething called with parameter : " + param);
    m.exec();
}

int MyApi::doSums( int a, int b )
{
    return a + b;
}
