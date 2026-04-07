/**
 * tsifl Calendar Add-on — Google Calendar integration
 * Manages calendar events via natural language through tsifl.
 *
 * Deploy: clasp push from this directory
 * Requires: Google Calendar API scope
 */

var BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
var SB_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
var SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

function onCalendarHomepage(e) {
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("tsifl").setSubtitle("AI Calendar Assistant"))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText("Open tsifl Sidebar")
          .setOnClickAction(CardService.newAction().setFunctionName("openSidebar")))
        .addWidget(CardService.newTextParagraph()
          .setText("Create, view, and manage events with AI."))
    );
  return card.build();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("tsifl")
    .setWidth(350);
  CalendarApp.getDefaultCalendar(); // Ensure permission
  SpreadsheetApp ? SpreadsheetApp.getUi().showSidebar(html) : null;
}

function onOpen(e) {
  try {
    // This runs when a calendar event is opened
  } catch (err) {}
}

// Auth functions (shared with workspace addon)
function signIn(email, password) {
  var resp = UrlFetchApp.fetch(SB_URL + "/auth/v1/token?grant_type=password", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SB_KEY },
    payload: JSON.stringify({ email: email, password: password }),
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function signUp(email, password) {
  var resp = UrlFetchApp.fetch(SB_URL + "/auth/v1/signup", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SB_KEY },
    payload: JSON.stringify({ email: email, password: password }),
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function refreshToken(token) {
  var resp = UrlFetchApp.fetch(SB_URL + "/auth/v1/token?grant_type=refresh_token", {
    method: "post",
    contentType: "application/json",
    headers: { "apikey": SB_KEY },
    payload: JSON.stringify({ refresh_token: token }),
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

// Calendar context
function getCalendarContext() {
  var cal = CalendarApp.getDefaultCalendar();
  var now = new Date();
  var weekLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  var events = cal.getEvents(now, weekLater);

  var eventList = events.slice(0, 20).map(function(ev) {
    return {
      id: ev.getId(),
      title: ev.getTitle(),
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString(),
      description: ev.getDescription() || "",
      location: ev.getLocation() || "",
      guests: ev.getGuestList().map(function(g) { return g.getEmail(); })
    };
  });

  return {
    app: "calendar",
    calendar_name: cal.getName(),
    timezone: Session.getScriptTimeZone(),
    upcoming_events: eventList,
    current_time: now.toISOString()
  };
}

// Send chat to backend
function sendChat(userId, message) {
  var context = getCalendarContext();
  var resp = UrlFetchApp.fetch(BACKEND_URL + "/chat/", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      user_id: userId,
      message: message,
      context: context,
      session_id: "calendar-" + userId
    }),
    muteHttpExceptions: true
  });

  var result = JSON.parse(resp.getContentText());
  var actions = result.actions || [];
  if (result.action && result.action.type) actions.push(result.action);

  var actionResults = [];
  for (var i = 0; i < actions.length; i++) {
    try {
      var r = executeCalendarAction(actions[i]);
      actionResults.push({ type: actions[i].type, success: true, message: r });
    } catch (e) {
      actionResults.push({ type: actions[i].type, success: false, message: e.message });
    }
  }

  return {
    reply: result.reply || "",
    tasks_remaining: result.tasks_remaining || -1,
    action_results: actionResults
  };
}

// Execute calendar actions
function executeCalendarAction(action) {
  var type = action.type;
  var p = action.payload || {};
  var cal = CalendarApp.getDefaultCalendar();

  switch (type) {
    case "create_event": {
      var start = new Date(p.start_time);
      var end = p.end_time ? new Date(p.end_time) : new Date(start.getTime() + 60 * 60 * 1000);
      var event = cal.createEvent(p.title || "New Event", start, end);
      if (p.description) event.setDescription(p.description);
      if (p.location) event.setLocation(p.location);
      if (p.attendees) {
        p.attendees.forEach(function(email) {
          event.addGuest(email);
        });
      }
      return "Created: " + (p.title || "New Event");
    }

    case "list_events": {
      var from = p.date ? new Date(p.date) : new Date();
      var days = p.days_ahead || 7;
      var to = new Date(from.getTime() + days * 24 * 60 * 60 * 1000);
      var events = cal.getEvents(from, to);
      return events.length + " events found in next " + days + " days";
    }

    case "update_event": {
      if (!p.event_id) return "No event ID provided";
      var event = cal.getEventById(p.event_id);
      if (!event) return "Event not found";
      if (p.title) event.setTitle(p.title);
      if (p.description) event.setDescription(p.description);
      if (p.start_time && p.end_time) event.setTime(new Date(p.start_time), new Date(p.end_time));
      return "Updated: " + event.getTitle();
    }

    case "delete_event": {
      if (!p.event_id) return "No event ID provided";
      var event = cal.getEventById(p.event_id);
      if (!event) return "Event not found";
      var title = event.getTitle();
      event.deleteEvent();
      return "Deleted: " + title;
    }

    case "open_notes":
      return "Open Notes in browser";

    default:
      return "Unknown action: " + type;
  }
}
