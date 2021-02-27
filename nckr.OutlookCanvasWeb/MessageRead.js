(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
        var appId = "PUT YOUR APP ID HERE";
        $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId + "?source=iframe");
    });
  };

})();