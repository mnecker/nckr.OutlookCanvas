(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
        var appId = "YOUR APP HERE";
        var item = Office.context.mailbox.item;
        var parameters =
            "&mailid=" + encodeURIComponent(item.itemId) +
            "&from=" + item.from.emailAddress +
            "&fromname=" + item.from.displayName +
            "&subject=" + item.subject +
            "&dateTimeReceived" + item.dateTimeCreated;
        $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId + "?source=iframe" + parameters);
    });
  };

})();