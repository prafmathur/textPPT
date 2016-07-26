/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var Globals = { activeViewHandler: 0, firstSlideId: 0 };
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            getSelectedRange();
            setInterval(makeRequest, 5000)
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

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function getSelectedRange() {
        // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
        Globals.firstSlideId = 0;

        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                Globals.firstSlideId = asyncResult.value.slides[0].id;
                showNotification(JSON.stringify(asyncResult.value));
            }
        });
    }


    function makeRequest() {
        $.ajax({
            url: "http://localhost:3000",
            crossDomain: true,
            dataType: "jsonp",
            success: function (data) {
                console.log(data)
            },
            error: function (e) {
                console.log(e)
            }
        });
    }



    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showNotification("Navigation successful");
            }
        });
    }

    function goToSlideByIndex() {
        var goToFirst = Office.Index.First;
        var goToLast = Office.Index.Last;
        var goToPrevious = Office.Index.Previous;
        var goToNext = Office.Index.Next;

        Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showNotification("Navigation successful");
            }
        });
    }


})();
