(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var appId = "c2650fc8-e65a-464e-a7b7-3862bc2213f1";
            var item = Office.context.mailbox.item;
            //console.log("itemid:"+item.itemId);
            var parameters =
                "&subject=" + item.subject +
                "&dateTimeReceived=" + item.dateTimeCreated +
                "&conversationId=" + item.conversationId +
                "&messageId=" + encodeURIComponent(item.itemId) +
                "&subject=" + item.normalizedSubject +
                "&sender=" + item.sender.emailAddress +
                "&to=" + buildEmailAddressesString(item.to) +
                "&cc=" + buildEmailAddressesString(item.cc) +
                "&bcc=" + buildEmailAddressesString(item.bcc);
            var url = "https://apps.powerapps.com/play/" + appId + "?source=iframe" + parameters;
            console.log(url);
            $('#canvas-iframe').attr("src", url);
        });
    };


  // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
      if (addresses && addresses.length > 0) {
          var joinedAdresses = addresses.map(item => item.emailAddress).join(";");
          return joinedAdresses;
      }
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();