(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
        var appId = "5e0af466-a56a-4762-8231-6b411af34c5a";
        var item = Office.context.mailbox.item;
        var parameters = "&mailid=" + item.itemId +
            "&from=" + item.from.emailAddress +
            "&fromname=" + item.from.displayName +
            "&subject=" + item.subject; 
        $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId + "?source=iframe" + parameters);
    });
  };

})();