#include "webpluginfactory.h"

/**
 * Constructor of a WebPluginFactory with a capability of creating plug-ins
 * of MIME type 'application/x-qt-widget'.
 * The descriptionHelper is passed to the created exampleWidget.
 */
WebPluginFactory::WebPluginFactory(DescriptionHelper *descriptionHelper, 
                                   QObject *parent) :
    QWebPluginFactory(parent),
    _descriptionHelper(descriptionHelper)
{
    // @TODO STEP 4.2
    /*
     * Make a new plugin structure, and initialise it with proper
     * values.
     * Add the MIME type 'application/x-qt-examplewidget' to the 
     * created plug-in structure.
     * Append the created plugin structure into the _plugins list.
     */
}

/**
 * Returns an instance of a plug-in, here, ExampleWidget.
 * Note that this plug-in object is deleted when the corresponding HTML
 * element is hidden.
 */
QObject * WebPluginFactory::create(const QString & mimeType,
                 const QUrl & url,
                 const QStringList & argumentNames,
                 const QStringList & argumentValues) const
{
    // @TODO STEP 4.3
    /*
     * Remove unnecessary Q_UNUSED macros.
     * If the MIME type is not 'application/x-qt-examplewidget',
     * return QObject. Else, return an ExampleWidget instance instead,
     * and provide _descriptionHelper for its constructor.
     */
    Q_UNUSED(url);

    Q_UNUSED(mimeType);

    Q_UNUSED(argumentNames);
    Q_UNUSED(argumentValues);

    return new QObject();
}

/**
 * Returns supported plug-ins.
 * Currently, this function is only called when JavaScript
 * programs access the global plug-ins or MIME type objects.
 */
QList<QWebPluginFactory::Plugin> WebPluginFactory::plugins () const
{
    return _plugins;
}
