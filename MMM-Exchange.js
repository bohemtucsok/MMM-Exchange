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
    showTasks: false,
    maxTasks: 10,
    header: "Exchange Calendar"
  },

  events: [],
  tasks: [],
  loading: true,
  tasksLoading: false,
  errorMessage: null,
  tasksError: null,

  start: function () {
    Log.info("[MMM-Exchange] Module started.");
    this.events = [];
    this.tasks = [];
    this.loading = true;
    this.errorMessage = null;
    this.tasksError = null;
    this.sendSocketNotification("FETCH_EVENTS", this.config);
    if (this.config.showTasks) {
      this.tasksLoading = true;
      this.sendSocketNotification("FETCH_TASKS", this.config);
    }
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

    // Tasks section
    if (self.config.showTasks) {
      var taskSection = document.createElement("div");
      taskSection.className = "task-section";

      var taskHeader = document.createElement("div");
      taskHeader.className = "task-header dimmed small";
      taskHeader.textContent = "Tasks";
      taskSection.appendChild(taskHeader);

      if (self.tasksLoading) {
        var loadingDiv = document.createElement("div");
        loadingDiv.className = "dimmed light xsmall";
        loadingDiv.textContent = "Loading tasks...";
        taskSection.appendChild(loadingDiv);
      } else if (self.tasksError) {
        var errorDiv = document.createElement("div");
        errorDiv.className = "dimmed light xsmall";
        errorDiv.textContent = "Tasks error: " + self.tasksError;
        taskSection.appendChild(errorDiv);
      } else if (!self.tasks || self.tasks.length === 0) {
        var emptyDiv = document.createElement("div");
        emptyDiv.className = "dimmed light xsmall";
        emptyDiv.textContent = "No tasks.";
        taskSection.appendChild(emptyDiv);
      } else {
        var taskTable = document.createElement("table");
        taskTable.className = "small";

        self.tasks.forEach(function (task) {
          var row = document.createElement("tr");
          row.className = "task-row";

          // Check if overdue
          var now = new Date();
          if (task.dueDate && new Date(task.dueDate) < now) {
            row.className += " task-overdue";
          }

          // Cell 1: Importance + Subject
          var subjectCell = document.createElement("td");
          subjectCell.className = "task-detail";

          var subjectSpan = document.createElement("span");
          subjectSpan.className = "task-subject";
          if (task.importance === "High") {
            subjectSpan.className += " task-importance-high";
          }
          subjectSpan.textContent = (task.importance === "High" ? "\u26A1 " : "") + task.subject;
          subjectCell.appendChild(subjectSpan);

          // Status line
          var statusSpan = document.createElement("span");
          statusSpan.className = "task-status dimmed xsmall";
          statusSpan.textContent = self.formatTaskStatus(task.status);
          if (task.percentComplete > 0 && task.percentComplete < 100) {
            statusSpan.textContent += " (" + task.percentComplete + "%)";
          }
          subjectCell.appendChild(statusSpan);

          row.appendChild(subjectCell);

          // Cell 2: Due date
          var dueCell = document.createElement("td");
          dueCell.className = "task-due light";
          if (task.dueDate) {
            dueCell.textContent = self.formatDate(new Date(task.dueDate));
          } else {
            dueCell.innerHTML = "&mdash;";
          }
          row.appendChild(dueCell);

          taskTable.appendChild(row);
        });

        taskSection.appendChild(taskTable);
      }

      wrapper.appendChild(taskSection);
    }

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
    if (notification === "TASKS_DATA") {
      this.tasksLoading = false;
      this.tasksError = null;
      this.tasks = payload;
      this.updateDom(this.config.animationSpeed);
    }
    if (notification === "TASKS_ERROR") {
      this.tasksLoading = false;
      this.tasksError = payload.message;
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
  },

  formatTaskStatus: function (status) {
    var statusMap = {
      "NotStarted": "Not started",
      "InProgress": "In progress",
      "Completed": "Completed",
      "WaitingOnOthers": "Waiting",
      "Deferred": "Deferred"
    };
    return statusMap[status] || status;
  }
});
