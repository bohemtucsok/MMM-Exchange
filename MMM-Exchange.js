Module.register("MMM-Exchange", {

  defaults: {
    username: "",
    password: "",
    host: "",
    updateInterval: 5 * 60 * 1000,
    maxEvents: 5,
    daysToFetch: 14,
    allowInsecureSSL: false,
    animationSpeed: 1000,
    showLocation: true,
    showEnd: true,
    header: "Exchange Calendar"
  },

  events: [],
  loading: true,
  errorMessage: null,

  start: function () {
    Log.info("[MMM-Exchange] Module started.");
    this.events = [];
    this.loading = true;
    this.errorMessage = null;
    this.sendSocketNotification("FETCH_EVENTS", this.config);
  },

  getStyles: function () {
    return ["MMM-Exchange.css"];
  },

  getHeader: function () {
    return this.config.header;
  },

  getDom: function () {
    var wrapper = document.createElement("div");
    wrapper.className = "mmm-exchange";

    // Loading state
    if (this.loading) {
      wrapper.innerHTML = "Loading calendar...";
      wrapper.className += " dimmed light small";
      return wrapper;
    }

    // Error state
    if (this.errorMessage) {
      wrapper.innerHTML = "Calendar error: " + this.errorMessage;
      wrapper.className += " dimmed light small";
      return wrapper;
    }

    // No events
    if (!this.events || this.events.length === 0) {
      wrapper.innerHTML = "No upcoming events.";
      wrapper.className += " dimmed light small";
      return wrapper;
    }

    // Render event list
    var table = document.createElement("table");
    table.className = "small";

    var self = this;
    this.events.forEach(function (event) {
      var startDate = new Date(event.start);
      var endDate = new Date(event.end);
      var now = new Date();

      var row = document.createElement("tr");
      row.className = "event-row";

      // Highlight currently active event
      if (now >= startDate && now <= endDate) {
        row.className += " bright current-event";
      }

      // Cell 1: Date/Time
      var timeCell = document.createElement("td");
      timeCell.className = "event-time light";

      var dateStr = self.formatDate(startDate);
      var timeStr = self.formatTime(startDate);
      timeCell.innerHTML = dateStr + "<br/>" + timeStr;

      if (self.config.showEnd) {
        timeCell.innerHTML += " - " + self.formatTime(endDate);
      }

      row.appendChild(timeCell);

      // Cell 2: Subject + Location
      var detailCell = document.createElement("td");
      detailCell.className = "event-detail";

      var subjectSpan = document.createElement("span");
      subjectSpan.className = "event-subject bright";
      subjectSpan.textContent = event.subject;
      detailCell.appendChild(subjectSpan);

      if (self.config.showLocation && event.location) {
        var locationSpan = document.createElement("span");
        locationSpan.className = "event-location dimmed xsmall";
        locationSpan.textContent = event.location;
        detailCell.appendChild(locationSpan);
      }

      row.appendChild(detailCell);
      table.appendChild(row);
    });

    wrapper.appendChild(table);
    return wrapper;
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === "EVENTS_DATA") {
      this.loading = false;
      this.errorMessage = null;
      this.events = payload;
      this.updateDom(this.config.animationSpeed);
    }
    if (notification === "EVENTS_ERROR") {
      this.loading = false;
      this.errorMessage = payload.message;
      this.updateDom(this.config.animationSpeed);
    }
  },

  formatDate: function (date) {
    var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var months = [
      "Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];
    return days[date.getDay()] + ", " +
           months[date.getMonth()] + " " +
           date.getDate();
  },

  formatTime: function (date) {
    var hours = date.getHours().toString().padStart(2, "0");
    var minutes = date.getMinutes().toString().padStart(2, "0");
    return hours + ":" + minutes;
  }
});
