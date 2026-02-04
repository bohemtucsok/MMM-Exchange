var NodeHelper = require("node_helper");
var EWS = require("node-ews");

module.exports = NodeHelper.create({

  ewsClient: null,
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

      this.initEwsClient();
      this.fetchEvents();
      this.scheduleNextFetch();
    }
  },

  initEwsClient: function () {
    if (this.ewsClient) return;

    var ewsConfig = {
      username: this.config.username,
      password: this.config.password,
      host: this.config.host,
      auth: "ntlm"
    };

    if (this.config.allowInsecureSSL) {
      process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
    }

    this.ewsClient = new EWS(ewsConfig);
  },

  fetchEvents: function () {
    var self = this;

    var startDate = new Date();
    var endDate = new Date();
    endDate.setDate(endDate.getDate() + (self.config.daysToFetch || 14));

    var ewsFunction = "FindItem";
    var ewsArgs = {
      attributes: {
        Traversal: "Shallow"
      },
      ItemShape: {
        BaseShape: "IdOnly",
        AdditionalProperties: {
          FieldURI: [
            { attributes: { FieldURI: "item:Subject" } },
            { attributes: { FieldURI: "calendar:Start" } },
            { attributes: { FieldURI: "calendar:End" } },
            { attributes: { FieldURI: "calendar:Location" } },
            { attributes: { FieldURI: "calendar:Organizer" } }
          ]
        }
      },
      CalendarView: {
        attributes: {
          MaxEntriesReturned: self.config.maxEvents || 5,
          StartDate: startDate.toISOString(),
          EndDate: endDate.toISOString()
        }
      },
      ParentFolderIds: {
        DistinguishedFolderId: {
          attributes: { Id: "calendar" }
        }
      }
    };

    console.log("[MMM-Exchange] Fetching calendar events...");

    self.ewsClient.run(ewsFunction, ewsArgs)
      .then(function (result) {
        var events = self.parseEvents(result);
        console.log("[MMM-Exchange] Fetched " + events.length + " events.");
        self.sendSocketNotification("EVENTS_DATA", events);
      })
      .catch(function (err) {
        console.error("[MMM-Exchange] EWS error:", err.message || err);
        self.sendSocketNotification("EVENTS_ERROR", {
          message: err.message || "Unknown EWS error"
        });
      });
  },

  parseEvents: function (result) {
    var events = [];

    try {
      var responseMessage = result.ResponseMessages.FindItemResponseMessage;

      if (
        responseMessage.attributes &&
        responseMessage.attributes.ResponseClass !== "Success"
      ) {
        console.error(
          "[MMM-Exchange] EWS response error:",
          responseMessage.ResponseCode
        );
        return events;
      }

      var rootFolder = responseMessage.RootFolder;

      if (!rootFolder || !rootFolder.Items || !rootFolder.Items.CalendarItem) {
        console.log("[MMM-Exchange] No calendar items found.");
        return events;
      }

      var calendarItems = rootFolder.Items.CalendarItem;

      // EWS returns a single object when there is only one item, normalize to array
      if (!Array.isArray(calendarItems)) {
        calendarItems = [calendarItems];
      }

      calendarItems.forEach(function (item) {
        events.push({
          subject: item.Subject || "(No subject)",
          start: item.Start || null,
          end: item.End || null,
          location: item.Location || "",
          organizer: item.Organizer
            ? item.Organizer.Mailbox
              ? item.Organizer.Mailbox.Name
              : ""
            : ""
        });
      });

      // Sort by start date ascending
      events.sort(function (a, b) {
        return new Date(a.start) - new Date(b.start);
      });

      // Limit to maxEvents
      var maxEvents = this.config.maxEvents || 5;
      events = events.slice(0, maxEvents);

    } catch (parseErr) {
      console.error(
        "[MMM-Exchange] Error parsing EWS response:",
        parseErr.message
      );
    }

    return events;
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
