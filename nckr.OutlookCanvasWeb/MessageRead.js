(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
        var appId = "85c5af77-76ca-4945-95d3-a862b0e57946";
        $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId + "?source=iframe");
    });
  };

})();