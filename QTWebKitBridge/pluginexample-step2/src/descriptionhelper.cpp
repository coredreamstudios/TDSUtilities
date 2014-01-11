#include "descriptionhelper.h"

/**
 * A very simple class to pass signals to the HTML world
 * and provide the slot descriptionNeedsToBeChanged to inform
 * that the description must be changed.
 */
DescriptionHelper::DescriptionHelper(QObject *parent) :
    QObject(parent),
    currentFormId(QString()),
    currentOldValue(QString())
{
}

/**
 * Method which is called when the description has been changed.
 * Emits the doDescriptionChange signal.
 */
void DescriptionHelper::doDescriptionChange(const QString descriptionValue)
{
    // @TODO STEP 3.1
    /*
     * Remove the unnecessary Q_UNUSED macros.
     * Emit the descriptionWasChanged signal with proper parameters.
     */
    Q_UNUSED(descriptionValue);
}

/**
 * Method which is called when the description change has been cancelled.
 * Emits the descriptionWasNotChanged signal.
 */
void DescriptionHelper::doCancel()
{
    // @TODO STEP 3.1
    /*
     * Clear the currentFormId and currentOldValue values.
     * Emit the descriptionWasNotChanged signal.
     */
}

/**
 * Slot which is called when the description needs to be changed.
 * Emits the openDescriptionWidget signal.
 */
void DescriptionHelper::descriptionNeedsToBeChanged(
                        const QString descriptionFormId, 
                        const QString oldValue)
{
    // @TODO STEP 3.1
    /*
     * Remove the unnecessary Q_UNUSED macros.
     * Store descriptionFormId and oldValue
     * into currentFormId and currentOldValue.
     * Emit the openDescriptionWidget signal.
     */
    Q_UNUSED(descriptionFormId);
    Q_UNUSED(oldValue);
}
