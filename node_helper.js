var NodeHelper = require("node_helper");
var https = require("https");
var http = require("http");
var DOMParser = require("@xmldom/xmldom").DOMParser;

module.exports = NodeHelper.create({

  fetchTimer: null,
  config: null,

  start: function () {
    console.log("[MMM-Exchange] node_helper started.");
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === "FETCH_EVENTS") {
      this.config = payload;

      if (!this.config.username || !this.config.password || !this.config.host) {
        this.sendSocketNotification("EVENTS_ERROR", {
          message: "Missing required config: username, password, or host"
        });
        return;
      }

      this.fetchEvents();
      this.scheduleNextFetch();
    }
  },

  buildSoapXml: function () {
    var startDate = new Date();
    var endDate = new Date();
    endDate.setDate(endDate.getDate() + (this.config.daysToFetch || 14));
    var maxEvents = this.config.maxEvents || 5;

    return '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
      'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
      'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
      '<soap:Body>' +
      '<m:FindItem Traversal="Shallow">' +
      '<m:ItemShape>' +
      '<t:BaseShape>IdOnly</t:BaseShape>' +
      '<t:AdditionalProperties>' +
      '<t:FieldURI FieldURI="item:Subject"/>' +
      '<t:FieldURI FieldURI="calendar:Start"/>' +
      '<t:FieldURI FieldURI="calendar:End"/>' +
      '<t:FieldURI FieldURI="calendar:Location"/>' +
      '<t:FieldURI FieldURI="calendar:Organizer"/>' +
      '</t:AdditionalProperties>' +
      '</m:ItemShape>' +
      '<m:CalendarView MaxEntriesReturned="' + maxEvents + '" ' +
      'StartDate="' + startDate.toISOString() + '" ' +
      'EndDate="' + endDate.toISOString() + '"/>' +
      '<m:ParentFolderIds>' +
      '<t:DistinguishedFolderId Id="calendar"/>' +
      '</m:ParentFolderIds>' +
      '</m:FindItem>' +
      '</soap:Body>' +
      '</soap:Envelope>';
  },

  fetchEvents: function () {
    var self = this;
    var soapXml = this.buildSoapXml();
    var ewsUrl = this.config.host.replace(/\/+$/, "") + "/EWS/Exchange.asmx";

    console.log("[MMM-Exchange] Fetching calendar events from " + ewsUrl);

    var authString = Buffer.from(this.config.username + ":" + this.config.password).toString("base64");

    this.httpPost(ewsUrl, soapXml, authString, 0, function (err, body) {
      if (err) {
        console.error("[MMM-Exchange] HTTP error:", err);
        self.sendSocketNotification("EVENTS_ERROR", {
          message: String(err)
        });
        return;
      }

      try {
        var events = self.parseXmlResponse(body);
        console.log("[MMM-Exchange] Fetched " + events.length + " events.");
        self.sendSocketNotification("EVENTS_DATA", events);
      } catch (parseErr) {
        console.error("[MMM-Exchange] Parse error:", parseErr.message);
        console.error("[MMM-Exchange] Response body (first 500 chars):", String(body).substring(0, 500));
        self.sendSocketNotification("EVENTS_ERROR", {
          message: "Failed to parse EWS response: " + parseErr.message
        });
      }
    });
  },

  httpPost: function (url, body, authBase64, redirectCount, callback) {
    var self = this;

    if (redirectCount > 5) {
      callback("Too many redirects");
      return;
    }

    var parsedUrl = new URL(url);
    var isHttps = parsedUrl.protocol === "https:";
    var transport = isHttps ? https : http;

    var options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || (isHttps ? 443 : 80),
      path: parsedUrl.pathname + parsedUrl.search,
      method: "POST",
      headers: {
        "Content-Type": "text/xml; charset=utf-8",
        "Content-Length": Buffer.byteLength(body, "utf8"),
        "Authorization": "Basic " + authBase64
      },
      rejectAuthorized: !(self.config.allowInsecureSSL)
    };

    if (self.config.allowInsecureSSL) {
      options.rejectUnauthorized = false;
    }

    var req = transport.request(options, function (res) {
      // Follow redirects (301, 302, 307, 308)
      if ([301, 302, 307, 308].indexOf(res.statusCode) > -1 && res.headers.location) {
        var redirectUrl = res.headers.location;
        // Handle relative redirects
        if (redirectUrl.startsWith("/")) {
          redirectUrl = parsedUrl.protocol + "//" + parsedUrl.host + redirectUrl;
        }
        console.log("[MMM-Exchange] Following redirect to: " + redirectUrl);

        // For 301/302, resend as POST with body
        self.httpPost(redirectUrl, body, authBase64, redirectCount + 1, callback);
        // Consume the response to free up the socket
        res.resume();
        return;
      }

      if (res.statusCode !== 200) {
        var errBody = "";
        res.on("data", function (chunk) { errBody += chunk; });
        res.on("end", function () {
          console.error("[MMM-Exchange] HTTP " + res.statusCode + " response:", errBody.substring(0, 300));
          callback("HTTP " + res.statusCode + " from Exchange server");
        });
        return;
      }

      var responseBody = "";
      res.on("data", function (chunk) {
        responseBody += chunk;
      });
      res.on("end", function () {
        callback(null, responseBody);
      });
    });

    req.on("error", function (err) {
      callback(err.message || err);
    });

    req.write(body);
    req.end();
  },

  parseXmlResponse: function (xmlString) {
    var events = [];
    var doc = new DOMParser().parseFromString(xmlString, "text/xml");

    // Check for SOAP fault
    var faults = doc.getElementsByTagName("Fault");
    if (faults.length > 0) {
      var faultString = faults[0].getElementsByTagName("faultstring");
      throw new Error("SOAP Fault: " + (faultString.length > 0 ? faultString[0].textContent : "Unknown"));
    }

    // Check ResponseClass
    var responseMsg = doc.getElementsByTagNameNS(
      "http://schemas.microsoft.com/exchange/services/2006/messages",
      "FindItemResponseMessage"
    );
    if (responseMsg.length > 0) {
      var responseClass = responseMsg[0].getAttribute("ResponseClass");
      if (responseClass !== "Success") {
        var responseCode = doc.getElementsByTagNameNS(
          "http://schemas.microsoft.com/exchange/services/2006/messages",
          "ResponseCode"
        );
        throw new Error("EWS error: " + (responseCode.length > 0 ? responseCode[0].textContent : responseClass));
      }
    }

    // Extract CalendarItems
    var calendarItems = doc.getElementsByTagNameNS(
      "http://schemas.microsoft.com/exchange/services/2006/types",
      "CalendarItem"
    );

    if (calendarItems.length === 0) {
      console.log("[MMM-Exchange] No calendar items found.");
      return events;
    }

    for (var i = 0; i < calendarItems.length; i++) {
      var item = calendarItems[i];
      events.push({
        subject: this.getTagText(item, "Subject") || "(No subject)",
        start: this.getTagText(item, "Start") || null,
        end: this.getTagText(item, "End") || null,
        location: this.getTagText(item, "Location") || "",
        organizer: this.getOrganizerName(item)
      });
    }

    // Sort by start date ascending
    events.sort(function (a, b) {
      return new Date(a.start) - new Date(b.start);
    });

    // Limit to maxEvents
    var maxEvents = this.config.maxEvents || 5;
    events = events.slice(0, maxEvents);

    return events;
  },

  getTagText: function (parent, tagName) {
    var elements = parent.getElementsByTagNameNS(
      "http://schemas.microsoft.com/exchange/services/2006/types",
      tagName
    );
    if (elements.length > 0 && elements[0].textContent) {
      return elements[0].textContent;
    }
    return null;
  },

  getOrganizerName: function (item) {
    var organizer = item.getElementsByTagNameNS(
      "http://schemas.microsoft.com/exchange/services/2006/types",
      "Organizer"
    );
    if (organizer.length > 0) {
      var name = organizer[0].getElementsByTagNameNS(
        "http://schemas.microsoft.com/exchange/services/2006/types",
        "Name"
      );
      if (name.length > 0) {
        return name[0].textContent || "";
      }
    }
    return "";
  },

  scheduleNextFetch: function () {
    var self = this;
    if (self.fetchTimer) {
      clearTimeout(self.fetchTimer);
    }
    var interval = self.config.updateInterval || 5 * 60 * 1000;
    self.fetchTimer = setTimeout(function () {
      self.fetchEvents();
      self.scheduleNextFetch();
    }, interval);
  }
});
