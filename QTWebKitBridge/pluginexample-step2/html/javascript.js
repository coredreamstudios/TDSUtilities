window.onload = init;

function init() {
    // @TODO STEP 3.3
    /*
     * Call initFormHandler for each form element.
     * Based on the comments, connect descriptionHelper's signals 
     * to appropriate JavaScript functions, defined by the comments.
     */
  
    // It is a good convention to use 'if' when using the Qt objects.

    // When the openDescriptionWidget signal is received,
    // the openDescriptionView function is called.

    // When the descriptionWasChanged signal is received,
    // descriptionReceiver gets the new description.

    // When the descriptionWasNotChanged signal is received,
    // the cancelDescriptionChange function is called.
}

// Note: This code block must be declared in a separate function.
// Otherwise, the formElement which is passed into the closure
// would point to the formElement variable of the *init* function,
// which would be equal to the last form element of the loop
// after the init function has finished.
function initFormHandler(formElement) {
    // @TODO STEP 3.3
    /*
     * Based on the comments, implement the onclick handler for
     * formElement according to the comments.
     */
    // Attach an onclick event handler to the a element inside formElement
    // Conditionally, since using the Qt objects,
    // emit the descriptionNeedsToBeChanged signal
}

// Shows the descriptionView and hides the mainView.
function openDescriptionView() {
    document.getElementById("mainView").style.display = "none";
    document.getElementById("descriptionView").style.display = "block";
    document.getElementById("body").setAttribute("class", "descriptionView");
}

// Shows the mainView and hides the descriptionView.
function openMainView() {
    document.getElementById("mainView").style.display = "block";
    document.getElementById("descriptionView").style.display = "none";
    document.getElementById("body").setAttribute("class", "mainView");
}

// Function to receive the new description.
function descriptionReceiver(descriptionElementId,newDescription) {
    document.getElementById(descriptionElementId)
        .getElementsByTagName("input")[0].value = newDescription;
    openMainView();
}

// Cancels the description change.
function cancelDescriptionChange(descriptionElementId) {
    openMainView();  
}

// Functionality to check results and show result correctness.
// The parameter for the ready() function is called when the body 
// element is loaded, possibly before the 'onload' event.
// @TODO STEP 7.3 STARTS
/** Based on the comments, add the required JavaScript code. **/
    // attach an onclick event handler for #checkButton

    // for each form of the describeform class:
        // get the given and the correct value,
        // then get the next p element with the result class, 
        // change its class and contents according to the 
        // correctness, and hide them.
    // Finally, slide all the handled result elements down 
    // with a minor delay.
// @TODO STEP 7.3 ENDS
