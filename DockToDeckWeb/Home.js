
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#urlInputsubmit').click(getUrlValue);
            $('#textInputsubmit').click(getTextValue);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function getUrlValue() {
        var inputVal = document.getElementById("urlInput").value;
        // use this value to call python script and then display the summary points in the existing or opened ppt
        // Displaying the value
        alert(inputVal);
    }

    function getTextValue() {
        var inputVal = document.getElementById("textInput").value;
        // use this value to call python script and then display the summary points in the existing or opened ppt
        // Displaying the value
        alert(inputVal);
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
