#include "dialog.h"
#include "ui_dialog.h"

Dialog::Dialog(QWidget *parent) : QDialog(parent), ui(new Ui::Dialog)
{
    ui->setupUi(this);

    ui->webView->load(QUrl("qrc:/testing.html"));
}

Dialog::~Dialog()
{
    delete ui;
}


