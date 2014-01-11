#ifndef WEBPLUGINFACTORY_H
#define WEBPLUGINFACTORY_H

#include <QWebPluginFactory>

#include <QLabel>

#include <QDebug>
#include <QUrl>
#include <QWebView>
#include <QWebFrame>

#include "examplewidget.h"


class WebPluginFactory : public QWebPluginFactory
{
    Q_OBJECT
public:
    explicit WebPluginFactory(DescriptionHelper *descriptionHelper, 
                              QObject *parent = 0);
    QObject * create(const QString & mimeType,
                     const QUrl & url,
                     const QStringList & argumentNames,
                     const QStringList & argumentValues) const;
    QList<QWebPluginFactory::Plugin> plugins () const;

signals:

public slots:

private:
    DescriptionHelper* _descriptionHelper;
    QList<QWebPluginFactory::Plugin> _plugins;
};

#endif // WEBPLUGINFACTORY_H
