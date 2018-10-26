// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

<reference path="../bower_components/jquery/dist/jquery.min.js" />
<reference path="../bower_components/hwcrypto/hwcrypto-legacy.js" />
<reference path="../bower_components/hwcrypto/hwcrypto.js" />
<reference path="../hex2base.js" />


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
  addTextToBody("Inserted by Jia hao sign Command Demo add-in.", "blue-icon-16", event);
  //sign(event);
  
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

var NO_NATIVE_URL = "https://open-eid.github.io/chrome-token-signing/missing.html";
var HELLO_URL = "https://open-eid.github.io/chrome-token-signing/hello.html";
var DEVELOPER_URL = "https://github.com/open-eid/chrome-token-signing/wiki/DeveloperTips";

var NATIVE_HOST = "ee.ria.esteid";

var K_SRC = "src";
var K_ORIGIN = "origin";
var K_NONCE = "nonce";
var K_RESULT = "result";
var K_TAB = "tab";
var K_EXTENSION = "extension";

// Stores the longrunning ports per tab
// Used to route all request from a tab to the same host instance
var ports = {};

// Probed to false if host component is OK.
var missing = true;

console.log("Background page activated");

// XXX: probe test, because connectNative() does not allow to check the presence
// of native component for some reason
typeof chrome.runtime.onStartup !== 'undefined' && chrome.runtime.onStartup.addListener(function() {
	// Also probed for in onInstalled()
	_testNativeComponent().then(function(result) {
		if (result === "ok") {
			missing = false;
		}
	});
});

// Force kill of native process
// Becasue Port.disconnect() does not work
function _killPort(tab) {
	if (tab in ports) {
		console.log("KILL " + tab);
		// Force killing with an empty message
		ports[tab].postMessage({});
	}
}

// Check if native implementation is OK resolves with "ok", "missing" or "forbidden"
function _testNativeComponent() {
	return new Promise(function(resolve, reject) {
		chrome.runtime.sendNativeMessage(NATIVE_HOST, {}, function(response) {
			if (!response) {
				console.log("TEST: ERROR " + JSON.stringify(chrome.runtime.lastError));
				// Try to be smart and do some string matching
				var permissions = "Access to the specified native messaging host is forbidden.";
				var missing = "Specified native messaging host not found.";
				if (chrome.runtime.lastError.message === permissions) {
					resolve("forbidden")
				} else if (chrome.runtime.lastError.message === missing) {
					resolve("missing");
				} else {
					resolve("missing");
				}
			} else {
				console.log("TEST: " + JSON.stringify(response));
				if (response["result"] === "invalid_argument") {
					resolve("ok");
				} else {
					resolve("missing"); // TODO: something better here
				}
			}
		});
	});
}


// When extension is installed, check for native component or direct to helping page
typeof chrome.runtime.onInstalled !== 'undefined' && chrome.runtime.onInstalled.addListener(function(details) {
	if (details.reason === "install" || details.reason === "update") {
		_testNativeComponent().then(function(result) {
				var url = null;
				if (result === "ok" && details.reason === "install") {
					// Also set the flag, onStartup() shall be called only
					// on next startup
					missing = false;
					// TODO: Add back HELLO page on install
					// once there is a nice tutorial
					// url = HELLO_URL;
				} else if (result === "forbidden") {
					url = DEVELOPER_URL;
				} else if (result === "missing"){
					url = NO_NATIVE_URL;
				}
				if (url) {
					chrome.tabs.create({'url': url + "?" + details.reason});
				}
		});
	}
});

// When message is received from page send it to native
chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
	if(sender.id !== chrome.runtime.id && sender.extensionId !== chrome.runtime.id) {
		console.log('WARNING: Ignoring message not from our extension');
		// Not our extension, do nothing
		return;
	}
	if (sender.tab) {
		// Check if page is DONE and close the native component without doing anything else
		if (request["type"] === "DONE") {
			console.log("DONE " + sender.tab.id);
			if (sender.tab.id in ports) {
				// FIXME: would want to use Port.disconnect() here
				_killPort(sender.tab.id);
			} 
		} else {
			request[K_TAB] = sender.tab.id;
			if (missing) {
				_testNativeComponent().then(function(result) {
					if (result === "ok") {
						missing = false;
						_forward(request);
					} else {
						return _fail_with(request, "no_implementation");
					}
				});
			} else {
				// TODO: Check if the URL is in allowed list or not
				// Either way forward to native currently
				_forward(request);
			}
		}
	}
});

// Send the message back to the originating tab
function _reply(tab, msg) {
	msg[K_SRC] = "background.js";
	msg[K_EXTENSION] = chrome.runtime.getManifest().version;
	chrome.tabs.sendMessage(tab, msg);
}

// Fail an incoming message if the underlying implementation is not
// present
function _fail_with(msg, result) {
	var resp = {};
	resp[K_NONCE] = msg[K_NONCE];
	resp[K_RESULT] = result;
	_reply(msg[K_TAB], resp);
}

// Forward a message to the native component
function _forward(message) {
	var tabid = message[K_TAB];
	console.log("SEND " + tabid + ": " + JSON.stringify(message));
	// Open a port if necessary
	if(!ports[tabid]) {
		// For some reason there does not seem to be a way to detect missing components from longrunning ports
		// So we probe before opening a new port.
		console.log("OPEN " + tabid + ": " + NATIVE_HOST);
		// create a new port
		var port = chrome.runtime.connectNative(NATIVE_HOST);
		// XXX: does not indicate anything for some reason.
		if (!port) {
			console.log("OPEN ERROR: " + JSON.stringify(chrome.runtime.lastError));
		}
		port.onMessage.addListener(function(response) {
			if (response) {
				console.log("RECV "+tabid+": " + JSON.stringify(response));
				_reply(tabid, response);
			} else {
				console.log("ERROR "+tabid+": " + JSON.stringify(chrome.runtime.lastError));
				_fail_with(message, "technical_error");
			}
		});
		port.onDisconnect.addListener(function() {
			console.log("QUIT " + tabid);
			delete ports[tabid];
			// TODO: reject all pending promises for tab, if any
		});
		ports[tabid] = port;
		ports[tabid].postMessage(message);
	} else {
		// Port already open
		ports[tabid].postMessage(message);
	}
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
        // Clear log
        document.getElementById('log').innerHTML = '';
        // Timestamp
        //log_text("sign() clicked on " + new Date().toUTCString());
		addTextToBody("sign() clicked on " + new Date().toUTCString(), event);
        // Select hash
        var hashtype = $("input[name=hash]:checked").val();
        // Set backend if asked
        var backend =  $("input[name=backend]:checked").val()
        // get language
        var lang = $("input[name=lang]:checked").val();
        if (!window.hwcrypto.use(backend)) {
          //log_text("Selecting backend failed.");
		  addTextToBody("Selecting backend failed.", event);
        }

        var hash = $("#hashvalue").val();
        //log_text("Signing " + hashtype + ": " + hash);
		addTextToBody("Signing " + hashtype + ": " + hash, event);
        // debug
        window.hwcrypto.debug().then(function(response) {
          //log_text("Debug: " + response);
		  addTextToBody("Debug: " + response, event);
        }, function(err) {
            //log_text("debug() failed: " + err);
			addTextToBody("debug() failed: " + err, event);
            return;
        });

        // Sign
        window.hwcrypto.getCertificate({lang: lang}).then(function(response) {
            var cert = response;
            //log_text("Using certificate:\n" + hexToPem(response.hex));
			addTextToBody("Using certificate:\n" + hexToPem(response.hex), event);
            window.hwcrypto.sign(cert, {type: hashtype, hex: hash}, {lang: lang}).then(function(response) {
                //log_text("Generated signature:\n" + response.hex.match(/.{1,64}/g).join("\n"));
				addTextToBody("Generated signature:\n" + response.hex.match(/.{1,64}/g).join("\n"), event);
            }, function(err) {
                //log_text("sign() failed: " + err);
				addTextToBody("sign() failed: " + err, event);
            });
        }, function(err) {
            //log_text("getCertificate() failed lo: " + err);
			addTextToBody("getCertificate() failed lo: " + err, event);
        });
    }

