#ifndef DESCRIPTIONHELPER_H
#define DESCRIPTIONHELPER_H

#include <QObject>
#include <QDebug>

class DescriptionHelper : public QObject
{
    Q_OBJECT
public:
    explicit DescriptionHelper(QObject *parent = 0);
    void doDescriptionChange(const QString descriptionValue);
    void doCancel();

signals:
    void descriptionWasChanged(QString descriptionFormId, 
                               QString descriptionValue);
    void descriptionWasNotChanged(QString descriptionFormId);
    void openDescriptionWidget();

public slots:
    void descriptionNeedsToBeChanged(const QString descriptionFormId, 
                                     const QString oldValue);

private:
    QString currentFormId;
    QString currentOldValue;
};

#endif // DESCRIPTIONHELPER_H
