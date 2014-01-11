#include "pluginexample.h"

/**
 * Plug-in example main class.
 */
PluginExample::PluginExample(QWidget *parent) : QMainWindow(parent)
{
    // Create the central widget and set it.
    QFrame* cW = new QFrame(this);
    setCentralWidget(cW);

    // Set the layout to the central widget.
    QVBoxLayout* layout = new QVBoxLayout(cW);
    cW->setLayout(layout);
    layout->setMargin(0);
    layout->setSpacing(0);

    // Create the webview which will be used to display the page.
    m_webView = createWebView();

    // Add it to the layout.
    layout->addWidget(m_webView);

    m_webView->show();
}

PluginExample::~PluginExample()
{

}

/**
 * Creates a new webview
 */
QWebView* PluginExample::createWebView()
{
    QWebSettings* defaultSettings = QWebSettings::globalSettings();
    // We use JavaScript, so set it to be enabled.
    defaultSettings->setAttribute(QWebSettings::JavascriptEnabled, true);
    // Plug-ins must be set to be enabled to use plug-ins.
    defaultSettings->setAttribute(QWebSettings::PluginsEnabled,true);

    /*
     * Let's enable the developer extras.
     * The DEBUG flag is defined in the .pro file.
     * The developer extras are available in the desktop version
     * when opening the context menu and choosing Inspect.
     */
#if defined(DEBUG)
    defaultSettings->setAttribute(QWebSettings::DeveloperExtrasEnabled,
                                  true);
#endif

    QWebView* webView = new QWebView(this);

    /*
     * Let's add the description Helper here, to make sure that
     * when the webview is initialised, window.descriptionHelper
     * can be used from the JavaScript side.
     */
    // @TODO STEP 3.2
    /*
     * Create a description widget and add it to the js window object
     * by using the addToJavaScriptWindowObject method of the frame.
     */

    /*
     * We also pass the web plug-in factory to the webview.
     */
    // @TODO STEP 4.1
    /*
     * Create and add a plug-in factory with the webview page's
     * setPluginFactory of web view's page.
     */

    /*
     * Developer extras need the context menu,
     * but let's disable it in the release mode.
     */
#if !defined(DEBUG)
    webView->setContextMenuPolicy(Qt::NoContextMenu);
#endif

    webView->load(QUrl("qrc:/html/index.html"));
    return webView;
}
