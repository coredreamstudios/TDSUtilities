#ifndef EXAMPLEWIDGET_H
#define EXAMPLEWIDGET_H

#include <QtGui/QWidget>
#include <QEvent>
#include <QObject>
#include <QUrl>
#include <QTreeWidgetItem>
#include <QDebug>
#include <QTime>
#include <QWebFrame>

#include "descriptionhelper.h"

namespace Ui {
    class ExampleWidget;
}

class ExampleWidget : public QWidget
{
    Q_OBJECT

public:
    explicit ExampleWidget(DescriptionHelper* descriptionHelper, 
                           QWidget *parent = 0);
    ~ExampleWidget();

protected:
    void changeEvent(QEvent *e);

private:
    Ui::ExampleWidget *ui;
    DescriptionHelper* _descriptionHelper;

public slots:
    void _cancelled();
    void _accepted();

};

#endif // EXAMPLEWIDGET_H
