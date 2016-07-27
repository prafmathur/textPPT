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


    function makeRequest() {
        $.ajax({
            //url: "http://52.161.27.50:3000",
            url: "http://localhost:3000",
            crossDomain: true,
            dataType: "jsonp",
            success: function (data) {
                console.log("Recieved command")
                interpretCommand("next");
            },
            error: function (e) {
                console.log("Error:")
                console.log(e)
            }
        });
    }

    function interpretCommand(command) {
        command = command.replace(/\s+/g, '').toLowerCase();
        if (command == "next") {
            console.log("Going to next slide");
            goToSlide(Office.Index.Next)
        }
        if (command == "previous") {
            console.log("Going to previous slide");
            goToSlide(Office.Index.Previous)
        }
        //Office.Index.First;
        //Office.Index.Last;
    }


    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
            if (asyncResult.status == "failed") {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
            else {
                console.log("Navigation successful");
            }
        });
    }


    function goToSlide(whichSlide) {
        Office.context.document.goToByIdAsync(whichSlide, Office.GoToType.Index, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showNotification("Navigation successful");
            }
        });
    }


    //function goToNextSlide() {

    //    Office.context.document.goToByIdAsync(4, Office.GoToType.Index, function (asyncResult) {
    //        if (asyncResult.status == "failed") {
    //            showNotification("Action failed with error: " + asyncResult.error.message);
    //        }
    //        else {
    //            showNotification("Navigation successful");
    //        }
    //    });
    //    //Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, function (asyncResult) {
    //    //    if (asyncResult.status == "failed") {
    //    //        showNotification("Action failed with error: " + asyncResult.error.message);
    //    //    }
    //    //    else {
    //    //        showNotification("Navigation successful");
    //    //    }
    //    //});


    //}

    //function goToSlideByIndex() {
    //    var goToFirst = Office.Index.First;
    //    var goToLast = Office.Index.Last;
    //    var goToPrevious = Office.Index.Previous;
    //    var goToNext = Office.Index.Next;

    //    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
    //        if (asyncResult.status == "failed") {
    //            showNotification("Action failed with error: " + asyncResult.error.message);
    //        }
    //        else {
    //            showNotification("Navigation successful");
    //        }
    //    });
    //}


})();
