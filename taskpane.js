Office.onReady(function() {
    // Attach event listener for document selection
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onDocumentSelectionChanged
    );

    // Add global keydown event listener
    document.addEventListener("keydown", function(event) {
        if (event.ctrlKey && event.shiftKey) {
            // Handle your desired action here
            activateAddIn();
        }
    });
});

function onDocumentSelectionChanged(eventArgs) {
    // Check if the selected text is non-empty
    if (eventArgs.getOoxml().trim() !== "") {
        // Enable or show your task pane here
        // For this example, we'll update the HTML content
        var container = document.getElementById("pronunciationResult");
        container.innerHTML = `
            <p>Selected Text: ${eventArgs.getOoxml()}</p>
        `;
    }
}

function activateAddIn() {
    // Handle the activation of your add-in
    // Show your task pane or perform the desired action
    var container = document.getElementById("pronunciationResult");
    container.innerHTML = `
        <p>Add-in Activated!</p>
    `;
}
