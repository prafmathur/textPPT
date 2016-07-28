/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var Globals = { activeViewHandler: 0, firstSlideId: 0 };
    var firstTime = true;
    var recentTimestamp;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        makeRequest()
        $(document).ready(function () {
            setInterval(makeRequest, 2000)
        });
    };

    function makeRequest() {
        $.ajax({
            url: "https://textppt.herokuapp.com",
            crossDomain: true,
            dataType: "jsonp",
            success: function (data) {
                if (firstTime) {
                    firstTime = false
                }
                else if (recentTimestamp != data.timestamp) {
                    interpretCommand(data.command);
                }
                recentTimestamp = data.timestamp;
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


    function goToSlide(whichSlide) {
        Office.context.document.goToByIdAsync(whichSlide, Office.GoToType.Index, function (asyncResult) {
            if (asyncResult.status == "failed") {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
            else {
                console.log("Navigation successful");
            }
        });
    }

})();
