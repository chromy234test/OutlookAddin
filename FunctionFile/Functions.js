// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.




Office.initialize = function () {
}

// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

// Adds text into the body of the item, then reports the results
// to the info bar.
function addTextToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        statusUpdate(icon, "\"" + text + "\"  jia jia inserted successfully.");
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
          type: "errorMessage",
          message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
        });
      }
      event.completed();
    });
}

function addDefaultMsgToBody(event) {
  addTextToBody("Inserted by the Add-in Command Demo add-in.", "blue-icon-16", event);
}

function signDefaultMsgToBody(event) {
  addTextToBody("Inserted by 111111 sign Command Demo add-in.", "blue-icon-16", event);
  sign(event);
  addTextToBody("Inserted by 222222 hao sign Command Demo add-in.", "blue-icon-16", event);
}


function addMsg1ToBody(event) {
  addTextToBody("Hello World!", "red-icon-16", event);
}

function addMsg2ToBody(event) {
  addTextToBody("Add-in commands are cool!", "red-icon-16", event);
}

function addMsg3ToBody(event) {
  addTextToBody("Visit https://developer.microsoft.com/en-us/outlook/ today for all of your add-in development needs.", "red-icon-16", event);
}

// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
  var subject = Office.context.mailbox.item.subject;
  
  Office.context.mailbox.item.notificationMessages.addAsync("subject", {
    type: "informationalMessage",
    icon: "blue-icon-16",
    message: "Subject: " + subject,
    persistent: false
  });
  
  event.completed();
}

// Gets the item class of the item and displays it in the info bar.
function getItemClass(event) {
  var itemClass = Office.context.mailbox.item.itemClass;
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemClass", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item Class: " + itemClass,
    persistent: false
  });
  
  event.completed();
}

// Gets the date and time when the item was created and displays it in the info bar.
function getDateTimeCreated(event) {
  var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
  
  Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Created: " + dateTimeCreated.toLocaleString(),
    persistent: false
  });
  
  event.completed();
}

// Gets the ID of the item and displays it in the info bar.
function getItemID(event) {
  // Limited to 150 characters max in the info bar, so 
  // only grab the first 50 characters of the ID
  var itemID = Office.context.mailbox.item.itemId.substring(0, 50);
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemID", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item ID: " + itemID,
    persistent: false
  });
  
  event.completed();
}

   var hashes = {};
    hashes['SHA-1'] = 'c33df55b8aee82b5018130f61a9199f6a9d5d385';
    hashes['SHA-224'] = '614eadb55ecd6c4938fe23a450edd51292519f7ffb51e91dc8aa5fbe';
    hashes['SHA-256'] = '413140d54372f9baf481d4c54e2d5c7bcf28fd6087000280e07976121dd54af2';
    hashes['SHA-384'] = '71839e04e1f8c6e3a6697e27e9a7b8aff24c95358ea7dc7f98476c1e4d88c67d65803d175209689af02aa3cbf69f2fd3';
    hashes['SHA-512'] = 'c793dc32d969cd4982a1d6e630de5aa0ebcd37e3b8bd0095f383a839582b080b9fe2d00098844bd303b8736ca1000344c5128ea38584bbed2d77a3968c7cdd71';
    hashes['SHA-192'] = 'ad41e82bcff23839dc0d9683d46fbae0be3dfcbbb1b49c70';

    function log_text(s) {
        var d = document.createElement("div");
        d.innerHTML = s;
        document.getElementById('log').appendChild(d);
    }

    function debug() {
        window.hwcrypto.debug().then(function(response) {log_text("Debug: " + response);});
    }

	
	function sign(event) {
		addTextToBody("sign event ", "blue-icon-16", event);
        // Clear log
        //document.getElementById('log').innerHTML = '';
        // Timestamp
        //log_text("sign() clicked on " + new Date().toUTCString());
		addTextToBody("sign() clicked on " + new Date().toUTCString(),"blue-icon-16", event);
        // Select hash
        var hashtype = $("input[name=hash]:checked").val();
        // Set backend if asked
        var backend =  $("input[name=backend]:checked").val()
        // get language
        var lang = $("input[name=lang]:checked").val();
        if (!window.hwcrypto.use(backend)) {
          //log_text("Selecting backend failed.");
		  addTextToBody("Selecting backend failed.","blue-icon-16", event);
        }

        var hash = $("#hashvalue").val();
        //log_text("Signing " + hashtype + ": " + hash);
		addTextToBody("Signing " + hashtype + ": " + hash,"blue-icon-16", event);
        // debug
        window.hwcrypto.debug().then(function(response) {
          //log_text("Debug: " + response);
		  addTextToBody("Debug: " + response, event);
        }, function(err) {
            //log_text("debug() failed: " + err);
			addTextToBody("debug() failed: " + err,"blue-icon-16", event);
            return;
        });

        // Sign
        window.hwcrypto.getCertificate({lang: lang}).then(function(response) {
            var cert = response;
            //log_text("Using certificate:\n" + hexToPem(response.hex));
			addTextToBody("Using certificate:\n" + hexToPem(response.hex),"blue-icon-16", event);
            window.hwcrypto.sign(cert, {type: hashtype, hex: hash}, {lang: lang}).then(function(response) {
                //log_text("Generated signature:\n" + response.hex.match(/.{1,64}/g).join("\n"));
				addTextToBody("Generated signature:\n" + response.hex.match(/.{1,64}/g).join("\n"), "blue-icon-16",event);
            }, function(err) {
                //log_text("sign() failed: " + err);
				addTextToBody("sign() failed: " + err,"blue-icon-16", event);
            });
        }, function(err) {
            //log_text("getCertificate() failed lo: " + err);
			addTextToBody("getCertificate() failed lo: " + err,"blue-icon-16", event);
        });
    }

