import React, { useMemo, useState } from "react";
import { Title } from "../common/Title";


function safeDate(value) {
  if (!value) return null;

  const d = new Date(value);

  return Number.isNaN(d.getTime()) ? null : d;
}


function formatTime(value) {
  const d = safeDate(value);

  if (!d) return "--";

  return d.toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit"
  });
}


function formatDate(value) {
  const d = safeDate(value);

  if (!d) return "--";

  return d.toLocaleDateString([], {
    weekday: "short",
    day: "2-digit",
    month: "short"
  });
}


function formatFullDate(value) {
  const d = safeDate(value);

  if (!d) return "--";

  return d.toLocaleString([], {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}


function cleanText(value) {
  if (!value) return "";

  return String(value)
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/\s+/g, " ")
    .trim();
}


function getOrganizer(event) {
  return (
    event.organizer?.emailAddress?.name ||
    event.organizer?.emailAddress?.address ||
    "Organizer unavailable"
  );
}


function getOrganizerEmail(event) {
  return event.organizer?.emailAddress?.address || "";
}


function getParticipants(event) {
  return Array.isArray(event.attendees)
    ? event.attendees
    : [];
}


function getLocation(event) {
  if (event.location?.displayName) {
    return event.location.displayName;
  }

  if (Array.isArray(event.locations) && event.locations.length) {
    return event.locations
      .map(x => x.displayName)
      .filter(Boolean)
      .join(", ");
  }

  return "No location";
}


function getMeetingLink(event) {
  return (
    event.onlineMeeting?.joinUrl ||
    event.onlineMeeting?.joinWebUrl ||
    event.webLink ||
    ""
  );
}


function getStatus(event) {

  if (event.isCancelled) {
    return "cancelled";
  }

  const start = safeDate(event.start?.dateTime);
  const end = safeDate(event.end?.dateTime);

  if (!start || !end) {
    return "upcoming";
  }

  const now = new Date();

  if (now >= start && now <= end) {
    return "ongoing";
  }

  if (end < now) {
    return "missed";
  }

  return "upcoming";
}


function getImportance(event) {
  return event.importance === "high";
}


function isToday(event) {
  const start = safeDate(event.start?.dateTime);

  if (!start) return false;

  const now = new Date();

  return (
    start.getFullYear() === now.getFullYear() &&
    start.getMonth() === now.getMonth() &&
    start.getDate() === now.getDate()
  );
}


function isSoon(event) {

  const start = safeDate(event.start?.dateTime);

  if (!start) return false;

  const now = new Date();

  const diff = start.getTime() - now.getTime();

  return diff >= 0 && diff <= 24 * 60 * 60 * 1000;
}


function isUrgent(event) {

  if (getStatus(event) === "ongoing") {
    return true;
  }

  if (getImportance(event)) {
    return true;
  }

  if (isSoon(event)) {
    return true;
  }

  return false;
}


function StatCard({
  icon,
  value,
  label,
  tone,
  onClick,
  active
}) {

  return (
    <button
      className={`calendarStat ${tone || ""} ${active ? "active" : ""}`}
      onClick={onClick}
    >

      <span className="calendarStatIcon">
        {icon}
      </span>

      <span className="calendarStatBody">
        <strong>{value}</strong>
        <small>{label}</small>
      </span>

      <span className="calendarStatGlow" />
    </button>
  );
}


function MeetingCard({
  event,
  important,
  onImportant,
  onMail,
  onJoin
}) {

  const [expanded, setExpanded] = useState(false);

  const status = getStatus(event);

  const participants = getParticipants(event);

  const description =
    cleanText(event.bodyPreview) ||
    cleanText(event.body?.content) ||
    "No description available.";

  const organizer = getOrganizer(event);

  const organizerEmail = getOrganizerEmail(event);

  const location = getLocation(event);

  const joinLink = getMeetingLink(event);

  return (

    <article
      className={`calendarMeeting ${status} ${
        important ? "important" : ""
      }`}
    >

      <div className="calendarMeetingRail" />

      <div className="calendarMeetingDate">

        <strong>
          {safeDate(event.start?.dateTime)
            ?.toLocaleDateString([], {
              day: "2-digit"
            }) || "--"}
        </strong>

        <small>
          {safeDate(event.start?.dateTime)
            ?.toLocaleDateString([], {
              month: "short"
            }) || ""}
        </small>

      </div>


      <div className="calendarMeetingMain">

        <div className="calendarMeetingTop">

          <div>

            <div className="calendarMeetingTime">

              {formatTime(event.start?.dateTime)}

              <span>→</span>

              {formatTime(event.end?.dateTime)}

              {status === "ongoing" && (
                <em className="meetingLive">
                  ● LIVE
                </em>
              )}

              {status === "missed" && (
                <em className="meetingMissed">
                  MISSED
                </em>
              )}

            </div>


            <h3>
              {event.subject || "Untitled meeting"}
            </h3>

          </div>


          <button
            className={`meetingStar ${
              important ? "selected" : ""
            }`}
            onClick={() => onImportant(event)}
            title={
              important
                ? "Remove important mark"
                : "Mark as important"
            }
          >
            {important ? "★" : "☆"}
          </button>

        </div>


        <div className="calendarMeetingMeta">

          <span>
            <b>ORG</b>
            {organizer}
          </span>

          <span>
            <b>PEOPLE</b>
            {participants.length}
          </span>

          <span>
            <b>LOCATION</b>
            {location}
          </span>

        </div>


        <div className="calendarMeetingDescription">

          {description.length > 220 && !expanded
            ? `${description.slice(0, 220)}...`
            : description}

          {description.length > 220 && (
            <button
              className="descriptionToggle"
              onClick={() => setExpanded(x => !x)}
            >
              {expanded ? "Show less" : "Read more"}
            </button>
          )}

        </div>


        {expanded && (

          <div className="calendarMeetingDetails">

            <div>

              <small>DATE & TIME</small>

              <strong>
                {formatFullDate(event.start?.dateTime)}
              </strong>

            </div>


            <div>

              <small>END</small>

              <strong>
                {formatFullDate(event.end?.dateTime)}
              </strong>

            </div>


            <div>

              <small>ORGANIZER</small>

              <strong>
                {organizer}
              </strong>

              {organizerEmail && (
                <span>{organizerEmail}</span>
              )}

            </div>


            <div>

              <small>MEETING TYPE</small>

              <strong>
                {event.isOnlineMeeting
                  ? event.onlineMeetingProvider ===
                    "teamsForBusiness"
                    ? "Microsoft Teams"
                    : "Online meeting"
                  : "Calendar meeting"}
              </strong>

            </div>


            <div className="meetingParticipants">

              <small>PARTICIPANTS</small>

              {participants.length === 0 ? (

                <span>
                  No participant information available.
                </span>

              ) : (

                <div className="participantList">

                  {participants.map((person, index) => (

                    <div
                      className="participant"
                      key={`${person.emailAddress?.address || "p"}-${index}`}
                    >

                      <span>
                        {(person.emailAddress?.name ||
                          person.emailAddress?.address ||
                          "?")
                          .charAt(0)
                          .toUpperCase()}
                      </span>

                      <div>

                        <strong>
                          {person.emailAddress?.name ||
                            person.emailAddress?.address ||
                            "Unknown"}
                        </strong>

                        <small>
                          {person.emailAddress?.address || ""}
                        </small>

                      </div>

                    </div>

                  ))}

                </div>

              )}

            </div>

          </div>

        )}


        <div className="calendarMeetingActions">

          {joinLink && (
            <button
              className="meetingAction primary"
              onClick={() => onJoin(event)}
            >
              ↗ Join meeting
            </button>
          )}


          {organizerEmail && (
            <button
              className="meetingAction"
              onClick={() => onMail(event)}
            >
              ✉ Mail organizer
            </button>
          )}


          <button
            className="meetingAction"
            onClick={() => setExpanded(x => !x)}
          >
            {expanded
              ? "⌃ Hide details"
              : "⌄ View details"}
          </button>

        </div>

      </div>

    </article>
  );
}


function InsightPanel({
  urgent,
  missed,
  important
}) {

  const nextUrgent = urgent[0];

  return (

    <section className="calendarInsightPanel">

      <div className="calendarInsightHeader">

        <div>

          <span>AI CALENDAR INTELLIGENCE</span>

          <h3>
            Your schedule, understood.
          </h3>

        </div>

        <div className="calendarInsightOrb">
          ✦
        </div>

      </div>


      <div className="calendarInsights">

        {nextUrgent ? (

          <div className="calendarInsight">

            <span className="insightIcon urgent">
              ⚡
            </span>

            <div>

              <strong>
                Upcoming meeting may need your attention
              </strong>

              <p>
                <b>
                  {nextUrgent.subject ||
                    "Untitled meeting"}
                </b>

                {" "}starts at{" "}

                {formatTime(
                  nextUrgent.start?.dateTime
                )}
                .
              </p>

            </div>

          </div>

        ) : (

          <div className="calendarInsight">

            <span className="insightIcon">
              ✓
            </span>

            <div>

              <strong>
                No immediate meeting pressure detected
              </strong>

              <p>
                Your calendar currently has no urgent
                upcoming event.
              </p>

            </div>

          </div>

        )}


        {missed.length > 0 && (

          <div className="calendarInsight">

            <span className="insightIcon missed">
              ◷
            </span>

            <div>

              <strong>
                {missed.length} meeting
                {missed.length === 1 ? "" : "s"} already ended
              </strong>

              <p>
                Review the missed meeting details and
                follow up with the organizer if required.
              </p>

            </div>

          </div>

        )}


        {important.length > 0 && (

          <div className="calendarInsight">

            <span className="insightIcon important">
              ★
            </span>

            <div>

              <strong>
                {important.length} important meeting
                {important.length === 1 ? "" : "s"} marked
              </strong>

              <p>
                These meetings are highlighted throughout
                your calendar.
              </p>

            </div>

          </div>

        )}

      </div>

    </section>
  );
}


export function Calendar({
  c = [],
  ask,
  onMailOrganizer,
  onJoinMeeting,
  onImportant
}) {

  const [filter, setFilter] = useState("all");

  const [search, setSearch] = useState("");

  const [importantIds, setImportantIds] = useState(
    () => {
      try {
        return JSON.parse(
          localStorage.getItem("wdImportantMeetings") || "[]"
        );
      } catch {
        return [];
      }
    }
  );


  const events = useMemo(() => {

    return [...c]
      .filter(Boolean)
      .sort((a, b) => {

        const da =
          safeDate(a.start?.dateTime)?.getTime() ||
          0;

        const db =
          safeDate(b.start?.dateTime)?.getTime() ||
          0;

        return da - db;
      });

  }, [c]);


  const categorized = useMemo(() => {

    const upcoming = events.filter(
      x => getStatus(x) === "upcoming"
    );

    const ongoing = events.filter(
      x => getStatus(x) === "ongoing"
    );

    const missed = events.filter(
      x => getStatus(x) === "missed"
    );

    const today = events.filter(
      x => isToday(x)
    );

    const urgent = events.filter(
      x => isUrgent(x) &&
      getStatus(x) !== "missed"
    );

    const important = events.filter(
      x =>
        getImportance(x) ||
        importantIds.includes(x.id)
    );

    return {
      upcoming,
      ongoing,
      missed,
      today,
      urgent,
      important
    };

  }, [events, importantIds]);


  const visibleEvents = useMemo(() => {

    let list = events;


    if (filter === "today") {
      list = categorized.today;
    }

    if (filter === "urgent") {
      list = categorized.urgent;
    }

    if (filter === "ongoing") {
      list = categorized.ongoing;
    }

    if (filter === "missed") {
      list = categorized.missed;
    }

    if (filter === "important") {
      list = categorized.important;
    }


    const q = search.trim().toLowerCase();

    if (q) {

      list = list.filter(event => {

        const subject =
          event.subject || "";

        const organizer =
          getOrganizer(event);

        const location =
          getLocation(event);

        const participants =
          getParticipants(event)
            .map(x =>
              x.emailAddress?.name ||
              x.emailAddress?.address ||
              ""
            )
            .join(" ");

        return `${subject} ${organizer} ${location} ${participants}`
          .toLowerCase()
          .includes(q);
      });
    }


    return list;

  }, [
    events,
    filter,
    search,
    categorized
  ]);


  function handleImportant(event) {

    const already =
      importantIds.includes(event.id) ||
      event.importance === "high";

    const next = already
      ? importantIds.filter(id => id !== event.id)
      : [...importantIds, event.id];


    setImportantIds(next);

    localStorage.setItem(
      "wdImportantMeetings",
      JSON.stringify(next)
    );


    if (onImportant) {
      onImportant(event, !already);
    }
  }


  function handleJoin(event) {

    const url =
      event.onlineMeeting?.joinUrl ||
      event.onlineMeeting?.joinWebUrl ||
      event.webLink;

    if (!url) {
      return;
    }

    if (onJoinMeeting) {
      onJoinMeeting(event);
      return;
    }

    window.open(
      url,
      "_blank",
      "noopener,noreferrer"
    );
  }


  if (!events.length) {

    return (

      <div className="module calendarModule">

        <Title
          k="TIME & FOCUS"
          t="Calendar intelligence"
        />

        <p className="sub">
          Your Microsoft calendar is currently empty
          or no calendar events were returned.
        </p>


        <div className="calendarEmpty">

          <div className="calendarEmptyIcon">
            ◫
          </div>

          <h3>
            No calendar events available
          </h3>

          <p>
            Workday Copilot could not find any
            accessible events in your Microsoft calendar.
          </p>

        </div>

      </div>
    );
  }


  return (

    <div className="module calendarModule">

      {/* HEADER */}

      <div className="calendarPageHeader">

        <div>

          <Title
            k="TIME & FOCUS"
            t="Calendar intelligence"
          />

          <p className="sub">
            Your Microsoft calendar, connected live.
          </p>

        </div>


        <div className="calendarLiveBadge">

          <i />

          LIVE MICROSOFT GRAPH

        </div>

      </div>


      {/* DASHBOARD */}

      <div className="calendarStats">

        <StatCard
          icon="⚡"
          value={categorized.urgent.length}
          label="Urgent"
          tone="urgent"
          active={filter === "urgent"}
          onClick={() =>
            setFilter(
              filter === "urgent"
                ? "all"
                : "urgent"
            )
          }
        />


        <StatCard
          icon="◷"
          value={categorized.today.length}
          label="Today's meetings"
          tone="today"
          active={filter === "today"}
          onClick={() =>
            setFilter(
              filter === "today"
                ? "all"
                : "today"
            )
          }
        />


        <StatCard
          icon="●"
          value={categorized.ongoing.length}
          label="Ongoing"
          tone="ongoing"
          active={filter === "ongoing"}
          onClick={() =>
            setFilter(
              filter === "ongoing"
                ? "all"
                : "ongoing"
            )
          }
        />


        <StatCard
          icon="↶"
          value={categorized.missed.length}
          label="Missed / ended"
          tone="missed"
          active={filter === "missed"}
          onClick={() =>
            setFilter(
              filter === "missed"
                ? "all"
                : "missed"
            )
          }
        />

      </div>


      {/* AI INSIGHTS */}

      <InsightPanel
        urgent={categorized.urgent}
        missed={categorized.missed}
        important={categorized.important}
      />


      {/* TOOLBAR */}

      <div className="calendarToolbar">

        <div>

          <span>
            {filter === "all"
              ? "ALL MEETINGS"
              : filter.toUpperCase()}
          </span>

          <small>
            {visibleEvents.length} meeting
            {visibleEvents.length === 1 ? "" : "s"}
          </small>

        </div>


        <div className="calendarToolbarControls">

          <button
            className={
              filter === "all"
                ? "selected"
                : ""
            }
            onClick={() => setFilter("all")}
          >
            All
          </button>

          <button
            className={
              filter === "important"
                ? "selected"
                : ""
            }
            onClick={() => setFilter("important")}
          >
            ★ Important
          </button>


          <div className="calendarSearch">

            <span>⌕</span>

            <input
              value={search}
              onChange={e =>
                setSearch(e.target.value)
              }
              placeholder="Search meetings..."
            />

          </div>

        </div>

      </div>


      {/* MEETINGS */}

      <div className="calendarList">

        {visibleEvents.length === 0 ? (

          <div className="calendarNoResults">

            <span>⌕</span>

            <h3>
              No meetings match this view
            </h3>

            <p>
              Try another filter or search term.
            </p>

          </div>

        ) : (

          visibleEvents.map(event => (

            <MeetingCard
              key={event.id}
              event={event}
              important={
                importantIds.includes(event.id) ||
                event.importance === "high"
              }
              onImportant={handleImportant}
              onMail={event =>
                onMailOrganizer
                  ? onMailOrganizer(event)
                  : ask?.(
                      `Help me prepare a message to the organizer of "${event.subject}".`
                    )
              }
              onJoin={handleJoin}
            />

          ))

        )}

      </div>

    </div>
  );
}


export default Calendar;
