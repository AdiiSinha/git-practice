
import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { MsalProvider, useMsal } from "@azure/msal-react";
import { msalInstance, loginRequest } from "./auth";
import {
  getProfile,
  getMail,
  getCalendar,
  sendMail,
  createDraft,
  markImportant,
  flagMail
} from "./graph";
import {
  teams,
  tasks,
  rooms,
  buses,
  news,
  projects
} from "./data";
import "./styles.css";

const API =
  import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";

const BASE_SUGGESTIONS = [
  "Catch me up",
  "What's urgent?",
  "What should I do next?",
  "What am I waiting for?",
  "What have I promised?",
  "Prepare me for my next meeting"
];

/* =========================================================
   ICONS
========================================================= */

function MicIcon({ muted = false, size = 18 }) {
  if (muted) {
    return (
      <svg
        width={size}
        height={size}
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
      >
        <path d="M12 2a3.5 3.5 0 0 0-3.5 3.5v5" />
        <path d="M8.5 10.5a3.5 3.5 0 0 0 7 0v-1" />
        <path d="M5.5 10.5v1.2a6.5 6.5 0 0 0 11.8 3.7" />
        <path d="M18.5 10.5v1.1" />
        <path d="M12 18.5V22" />
        <path d="M8.5 22h7" />
        <path d="M3 3l18 18" />
      </svg>
    );
  }

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <rect x="8.5" y="2.5" width="7" height="12" rx="3.5" />
      <path d="M5 11.5a7 7 0 0 0 14 0" />
      <path d="M12 18.5V22" />
      <path d="M8.5 22h7" />
    </svg>
  );
}

function SpeakerOffIcon({ size = 18 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M11 5 6.8 9H3v6h3.8L11 19z" />
      <path d="m16 9 5 6" />
      <path d="m21 9-5 6" />
    </svg>
  );
}

function SpeakerIcon({ size = 18 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M11 5 6.8 9H3v6h3.8L11 19z" />
      <path d="M15.5 8.5a5 5 0 0 1 0 7" />
      <path d="M18.5 6a8.5 8.5 0 0 1 0 12" />
    </svg>
  );
}

function LogoutIcon({ size = 18 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M10 4H5.5A1.5 1.5 0 0 0 4 5.5v13A1.5 1.5 0 0 0 5.5 20H10" />
      <path d="M14 8l4 4-4 4" />
      <path d="M18 12H9" />
    </svg>
  );
}

function SendIcon({ size = 17 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="m22 2-7 20-4-9-9-4Z" />
      <path d="M22 2 11 13" />
    </svg>
  );
}

function StopIcon({ size = 14 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="currentColor"
      aria-hidden="true"
    >
      <rect x="5" y="5" width="14" height="14" rx="2" />
    </svg>
  );
}

/* =========================================================
   AI API
========================================================= */

async function ai(payload) {
  const r = await fetch(`${API}/api/copilot`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  const d = await r.json();

  if (!r.ok) {
    throw Error(d.detail || "AI request failed");
  }

  return d;
}

/* =========================================================
   SPEECH CLEANUP
========================================================= */

function cleanSpeech(text) {
  return String(text || "")
    .replace(/[#*_`]/g, "")
    .replace(/\[(.*?)\]\(.*?\)/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

/* =========================================================
   MAIN APP
========================================================= */

function App() {
  const { instance, accounts } = useMsal();

  const account = accounts[0];

  const [mod, setMod] = useState("home");

  const [d, setD] = useState({
    p: null,
    m: [],
    c: []
  });

  const [msgs, setMsgs] = useState([]);

  const [q, setQ] = useState("");

  const [state, setState] = useState("idle");

  const [toast, setToast] = useState("");

  const [modal, setModal] = useState(null);

  const [profileOpen, setProfileOpen] = useState(false);

  const [splash, setSplash] = useState(false);

  const [suggestions, setSuggestions] =
    useState(BASE_SUGGESTIONS);

  const [autoListen, setAutoListen] = useState(true);

  /* NEW: AI speech mute state */
  const [muted, setMuted] = useState(false);

  const [drafts, setDrafts] = useState(() =>
    JSON.parse(localStorage.getItem("wdDrafts") || "[]")
  );

  const recognitionRef = useRef(null);

  const autoTimerRef = useRef(null);

  const mountedRef = useRef(true);

  /* =========================================================
     CLEANUP
  ========================================================= */

  useEffect(() => {
    return () => {
      mountedRef.current = false;

      clearTimeout(autoTimerRef.current);

      recognitionRef.current?.stop?.();

      window.speechSynthesis?.cancel();
    };
  }, []);

  /* =========================================================
     ACCOUNT
  ========================================================= */

  useEffect(() => {
    if (
      account &&
      !instance.getActiveAccount()
    ) {
      instance.setActiveAccount(account);
    }
  }, [account, instance]);

  /* =========================================================
     INITIAL LOAD
  ========================================================= */

  useEffect(() => {
    if (!account) return;

    setSplash(true);

    const timer = setTimeout(
      () => setSplash(false),
      1500
    );

    load();

    return () => clearTimeout(timer);
  }, [account]);

  /* =========================================================
     TOKEN
  ========================================================= */

  async function token() {
    try {
      return (
        await instance.acquireTokenSilent({
          scopes: loginRequest.scopes,
          account:
            instance.getActiveAccount() ||
            account
        })
      ).accessToken;
    } catch {
      return (
        await instance.acquireTokenPopup({
          scopes: loginRequest.scopes
        })
      ).accessToken;
    }
  }

  /* =========================================================
     LOAD GRAPH DATA
  ========================================================= */

  async function load() {
    try {
      const t = await token();

      const [p, m, c] =
        await Promise.all([
          getProfile(t),
          getMail(t),
          getCalendar(t)
        ]);

      if (mountedRef.current) {
        setD({
          p,
          m: m.value || [],
          c: c.value || []
        });
      }
    } catch (e) {
      setToast(
        `Graph data load failed: ${e.message}`
      );
    }
  }

  /* =========================================================
     AI CONTEXT
  ========================================================= */

  const ctx = useMemo(
    () => ({
      profile: d.p,

      emails: d.m,

      calendar: d.c,

      teams: teams.map((x) => ({
        person: x[0],
        project: x[1],
        text: x[2],
        priority: x[3]
      })),

      tasks: tasks.map((x) => ({
        title: x[0],
        project: x[1],
        due: x[2],
        priority: x[3],
        why: x[4]
      })),

      projects,

      workplace: {
        rooms,
        buses,
        news
      }
    }),
    [d]
  );

  /* =========================================================
     LOCAL MEMORY
  ========================================================= */

  function saveMemory(question, answer) {
    const old = JSON.parse(
      localStorage.getItem("wdmem") || "[]"
    );

    localStorage.setItem(
      "wdmem",
      JSON.stringify(
        [
          ...old,
          {
            q: question,
            a: answer,
            at: new Date().toISOString()
          }
        ].slice(-40)
      )
    );
  }

  /* =========================================================
     DYNAMIC SUGGESTIONS
  ========================================================= */

  function setDynamicSuggestions(items) {
    const clean = Array.from(
      new Set(
        (items || [])
          .filter(Boolean)
          .map(String)
      )
    ).slice(0, 6);

    setSuggestions(
      clean.length
        ? clean
        : BASE_SUGGESTIONS
    );
  }

  /* =========================================================
     ASK AI
  ========================================================= */

  async function ask(text) {
    if (!text.trim()) return;

    stopRecognition();

    clearTimeout(autoTimerRef.current);

    setQ("");

    setMod("copilot");

    setMsgs((x) => [
      ...x,
      {
        r: "u",
        t: text
      }
    ]);

    setState("thinking");

    try {
      const memory = JSON.parse(
        localStorage.getItem("wdmem") || "[]"
      );

      const z = await ai({
        query: text,
        context: ctx,
        memory
      });

      const answer =
        String(z.answer || "").trim();

      setMsgs((x) => [
        ...x,
        {
          r: "a",
          t: answer,
          actions: z.actions || []
        }
      ]);

      saveMemory(text, answer);

      setDynamicSuggestions(z.suggestions);

      speak(answer, true);
    } catch (e) {
      const answer = fallback(text);

      setMsgs((x) => [
        ...x,
        {
          r: "a",
          t: answer,
          actions: []
        }
      ]);

      saveMemory(text, answer);

      setDynamicSuggestions(
        dynamicFallback(text)
      );

      setToast(e.message);

      speak(answer, true);
    }
  }

  function dynamicFallback(text) {
    const x = text.toLowerCase();

    if (x.includes("meeting")) {
      return [
        "Prepare for the next meeting",
        "What's due before this meeting?",
        "What changed today?"
      ];
    }

    if (x.includes("mail")) {
      return [
        "Summarize important mail",
        "Draft replies",
        "What needs a response?"
      ];
    }

    return BASE_SUGGESTIONS;
  }

  function fallback(text) {
    const x = text.toLowerCase();

    if (
      x.includes("urgent") ||
      x.includes("next")
    ) {
      return (
        "1. Review the Phoenix deployment approval — due today at 5 PM. It blocks deployment.\n" +
        "2. Verify the Atlas login defect before the 4 PM client demo."
      );
    }

    if (x.includes("waiting")) {
      return (
        "1. Database credentials — waiting for DevOps.\n" +
        "This is blocking Atlas work."
      );
    }

    if (x.includes("promised")) {
      return "1. Nova test report — committed for EOD.";
    }

    if (x.includes("meeting")) {
      return (
        "Next meeting: client demo. Review the Atlas login defect and Phoenix deployment status before joining."
      );
    }

    if (x.includes("catch")) {
      return (
        `1. ${
          d.m[0]?.subject ||
          "No new important email"
        }\n${
          d.m[0]?.bodyPreview ||
          "Your inbox is quiet right now."
        }`
      );
    }

    return (
      `I can use your connected mail (${d.m.length}), calendar (${d.c.length}) and synthetic Teams context (${teams.length}) for this prototype.`
    );
  }

  /* =========================================================
     VOICE
  ========================================================= */

  function chooseVoice() {
    const voices =
      window.speechSynthesis?.getVoices?.() ||
      [];

    return (
      voices.find((v) =>
        /Microsoft.*(Aria|Jenny|Guy|Ryan|Sonia|Libby)/i.test(
          v.name
        )
      ) ||
      voices.find((v) =>
        /Google.*English/i.test(v.name)
      ) ||
      voices.find((v) =>
        /^en[-_]/i.test(v.lang)
      ) ||
      voices[0]
    );
  }

  function speak(
    text,
    continueListening = false
  ) {
    if (!window.speechSynthesis) return;

    /*
      NEW:
      If muted, do not speak at all.
    */
    if (muted) {
      setState("idle");
      return;
    }

    clearTimeout(autoTimerRef.current);

    window.speechSynthesis.cancel();

    const u =
      new SpeechSynthesisUtterance(
        cleanSpeech(text)
      );

    const voiceChoice = chooseVoice();

    if (voiceChoice) {
      u.voice = voiceChoice;
    }

    u.lang = "en-IN";

    u.rate = 0.94;

    u.pitch = 0.92;

    u.volume = 1;

    u.onstart = () => {
      if (mountedRef.current) {
        setState("speaking");
      }
    };

    u.onend = () => {
      if (!mountedRef.current) return;

      setState("idle");

      if (
        continueListening &&
        autoListen &&
        !muted
      ) {
        autoTimerRef.current =
          setTimeout(
            () => startListening(true),
            2300
          );
      }
    };

    u.onerror = () => {
      if (mountedRef.current) {
        setState("idle");
      }
    };

    window.speechSynthesis.speak(u);
  }

  /* =========================================================
     MUTE
  ========================================================= */

  function toggleMute() {
    setMuted((current) => {
      const next = !current;

      if (next) {
        window.speechSynthesis?.cancel();

        if (
          mountedRef.current &&
          state === "speaking"
        ) {
          setState("idle");
        }

        setToast("AI voice muted.");
      } else {
        setToast("AI voice enabled.");
      }

      return next;
    });
  }

  /* =========================================================
     STOP SPEAKING
  ========================================================= */

  function stopSpeaking() {
    window.speechSynthesis?.cancel();

    setState("idle");
  }

  /* =========================================================
     STOP RECOGNITION
  ========================================================= */

  function stopRecognition() {
    clearTimeout(autoTimerRef.current);

    try {
      recognitionRef.current?.stop?.();
    } catch {}

    recognitionRef.current = null;

    if (state === "listening") {
      setState("idle");
    }
  }

  /* =========================================================
     START LISTENING
  ========================================================= */

  function startListening(isAuto = false) {
    const R =
      window.SpeechRecognition ||
      window.webkitSpeechRecognition;

    if (!R) {
      setToast(
        "Speech recognition is not supported in this browser. Use Chrome or Edge."
      );

      return;
    }

    if (state === "speaking") {
      stopSpeaking();
    }

    stopRecognition();

    const r = new R();

    r.lang = "en-IN";

    r.interimResults = true;

    r.continuous = false;

    r.onstart = () => {
      if (mountedRef.current) {
        setState("listening");
      }
    };

    r.onresult = (e) => {
      const finalText = Array.from(
        e.results
      )
        .map(
          (v) => v[0].transcript
        )
        .join(" ")
        .trim();

      if (!finalText) return;

      if (
        /^(stop|stop listening|cancel|quiet)$/i.test(
          finalText
        )
      ) {
        stopRecognition();

        setState("idle");

        return;
      }

      if (
        e.results[
          e.results.length - 1
        ].isFinal
      ) {
        ask(finalText);
      }
    };

    r.onerror = (e) => {
      if (
        mountedRef.current &&
        e.error !== "aborted" &&
        e.error !== "no-speech"
      ) {
        setToast(
          `Voice input: ${e.error}`
        );
      }

      if (mountedRef.current) {
        setState("idle");
      }
    };

    r.onend = () => {
      recognitionRef.current = null;

      if (
        mountedRef.current &&
        state === "listening"
      ) {
        setState("idle");
      }
    };

    recognitionRef.current = r;

    try {
      r.start();
    } catch {
      setState("idle");
    }

    if (!isAuto) {
      setAutoListen(true);
    }
  }

  /* =========================================================
     TOGGLE VOICE
  ========================================================= */

  function toggleVoice() {
    if (state === "speaking") {
      return stopSpeaking();
    }

    if (state === "listening") {
      return stopRecognition();
    }

    startListening(false);
  }

  /* =========================================================
     LOGIN / LOGOUT
  ========================================================= */

  async function login() {
    await instance.loginPopup(loginRequest);
  }

  async function logout() {
    stopSpeaking();

    stopRecognition();

    setProfileOpen(false);

    await instance.logoutPopup();
  }

  /* =========================================================
     SEND EMAIL
  ========================================================= */

  async function doSend(x) {
    try {
      const t = await token();

      await sendMail(t, x);

      setModal(null);

      setToast(
        `Email sent successfully to ${x.to}`
      );

      load();
    } catch (e) {
      setToast(
        `Send failed: ${e.message}`
      );
    }
  }

  /* =========================================================
     DRAFT
  ========================================================= */

  async function doDraft(x) {
    try {
      const t = await token();

      const saved =
        await createDraft(t, x);

      setDrafts((prev) => {
        const next = [
          {
            ...x,
            id: saved.id,
            savedAt:
              new Date().toISOString()
          },
          ...prev
        ].slice(0, 20);

        localStorage.setItem(
          "wdDrafts",
          JSON.stringify(next)
        );

        return next;
      });

      setModal(null);

      setToast(
        "Draft saved to Outlook Drafts."
      );
    } catch (e) {
      setToast(
        `Draft save failed: ${e.message}`
      );
    }
  }

  /* =========================================================
     FLAG
  ========================================================= */

  async function doFlag(mail) {
    try {
      await flagMail(
        await token(),
        mail.id
      );

      setToast(
        "Mail flagged for follow-up."
      );

      load();
    } catch (e) {
      setToast(
        `Flag failed: ${e.message}`
      );
    }
  }

  /* =========================================================
     IMPORTANT
  ========================================================= */

  async function doImportant(mail) {
    try {
      await markImportant(
        await token(),
        mail.id
      );

      setToast(
        "Mail marked important."
      );

      load();
    } catch (e) {
      setToast(
        `Important action failed: ${e.message}`
      );
    }
  }

  /* =========================================================
     REPLY
  ========================================================= */

  function openReply(mail) {
    const sender =
      mail.from?.emailAddress?.address ||
      "";

    setModal({
      mode: "reply",

      to: sender,

      cc: "",

      subject: `Re: ${
        mail.subject || ""
      }`,

      body:
        `Hi ${
          mail.from?.emailAddress?.name ||
          "there"
        },\n\n` +
        `Thanks for the update. I’ll review this and get back to you shortly.\n\n` +
        `Regards,\n${
          d.p?.displayName ||
          "Employee"
        }`,

      sourceMail: mail
    });
  }

  /* =========================================================
     LOGIN
  ========================================================= */

  if (!account) {
    return <Login onLogin={login} />;
  }

  /* =========================================================
     SPLASH
  ========================================================= */

  if (splash) {
    return (
      <div className="splash">
        <Orb
          mode="thinking"
          size="large"
        />

        <small>
          CONTEXT ENGINE INITIALIZING
        </small>

        <h2>
          Good afternoon,{" "}
          {d.p?.givenName ||
            account.name ||
            "Employee"}
          .
        </h2>

        <p>
          Connecting your workday
          context…
        </p>

        <div className="load" />

        <div className="tags">
          MAIL · CALENDAR · TEAMS
          CONTEXT · WORKPLACE
        </div>
      </div>
    );
  }

  const voiceLabel =
    state === "listening"
      ? "Listening"
      : state === "speaking"
      ? "AI speaking"
      : state === "thinking"
      ? "Thinking"
      : "Ready";

  /* =========================================================
     APP UI
  ========================================================= */

  return (
    <div className="app">

      {/* =====================================================
          HEADER
      ===================================================== */}

      <header>

        <div className="brand">
          <b>✦</b>

          <span>
            WORKDAY{" "}
            <em>COPILOT</em>

            <small>
              EMPLOYEE EXPERIENCE AI
            </small>
          </span>
        </div>

        <div className="live">
          <span className="liveDot">
            ●
          </span>

          LIVE CONTEXT ·{" "}
          {d.m.length} MAILS ·{" "}
          {d.c.length} EVENTS ·
          TEAMS SYNTHETIC
        </div>

        <div className="top">

          {/* MIC */}

          <button
            className={`voiceTop ${
              state !== "idle"
                ? "active"
                : ""
            }`}
            onClick={toggleVoice}
            title={voiceLabel}
          >
            <span className="voiceGlyph">
              <MicIcon
                size={17}
                muted={
                  state !== "listening"
                }
              />
            </span>

            <small>
              {voiceLabel}
            </small>
          </button>

          {/* STOP */}

          <button
            className="stopTop"
            onClick={() => {
              stopSpeaking();
              stopRecognition();
            }}
            title="Stop voice"
          >
            <StopIcon />
          </button>

          {/* AUTO */}

          <button
            className="autoTop"
            onClick={() =>
              setAutoListen(
                (v) => !v
              )
            }
            title="Automatic listening after responses"
          >
            AUTO{" "}
            {autoListen
              ? "ON"
              : "OFF"}
          </button>

          {/* PROFILE */}

          <button
            className="avatar"
            onClick={() =>
              setProfileOpen(true)
            }
            title="Open employee profile"
          >
            {(
              d.p?.displayName ||
              account.name ||
              "A"
            )
              .slice(0, 1)
              .toUpperCase()}
          </button>

        </div>
      </header>

      {/* =====================================================
          LAYOUT
      ===================================================== */}

      <div className="layout">

        {/* ===================================================
            LEFT SIDEBAR
        =================================================== */}

        <aside>

          <small>
            WORKSPACE
          </small>

          <Nav
            a={mod}
            s={setMod}
            id="home"
            t="Today"
            i="⌂"
          />

          <Nav
            a={mod}
            s={setMod}
            id="copilot"
            t="AI Copilot"
            i="✦"
          />

          <Nav
            a={mod}
            s={setMod}
            id="mail"
            t="Inbox"
            i="✉"
            n={d.m.length}
          />

          <Nav
            a={mod}
            s={setMod}
            id="calendar"
            t="Calendar"
            i="◷"
            n={d.c.length}
          />

          <Nav
            a={mod}
            s={setMod}
            id="commit"
            t="Commitments"
            i="◎"
          />

          <Nav
            a={mod}
            s={setMod}
            id="projects"
            t="Projects"
            i="◇"
          />

          <Nav
            a={mod}
            s={setMod}
            id="waiting"
            t="Waiting For"
            i="⌛"
          />

          <small className="lower">
            WORKPLACE
          </small>

          <Nav
            a={mod}
            s={setMod}
            id="workplace"
            t="Campus"
            i="▦"
          />

          <Nav
            a={mod}
            s={setMod}
            id="rooms"
            t="Meeting Rooms"
            i="▣"
          />

          <footer>
            <span className="shield">
              ✓
            </span>

            Identity protected

            <br />

            <button
              onClick={logout}
            >
              <LogoutIcon size={13} />
              Sign out
            </button>
          </footer>

        </aside>

        {/* ===================================================
            MAIN
        =================================================== */}

        <main>

          {toast && (
            <div className="toast">

              <span>✦</span>

              {toast}

              <button
                onClick={() =>
                  setToast("")
                }
              >
                ×
              </button>

            </div>
          )}

          {mod === "home" && (
            <Home
              p={d.p}
              m={d.m}
              c={d.c}
              ask={ask}
              set={setMod}
              suggestions={
                suggestions
              }
            />
          )}

          {mod === "copilot" && (
            <Copilot
              msgs={msgs}
              q={q}
              setQ={setQ}
              ask={ask}
              state={state}
              voice={toggleVoice}
              stop={() => {
                stopSpeaking();
                stopRecognition();
              }}
              mute={toggleMute}
              muted={muted}
              modal={setModal}
              suggestions={
                suggestions
              }
              autoListen={
                autoListen
              }
            />
          )}

          {mod === "mail" && (
            <Mail
              m={d.m}
              ask={ask}
              reply={openReply}
              flag={doFlag}
              important={
                doImportant
              }
            />
          )}

          {mod === "calendar" && (
            <Calendar c={d.c} />
          )}

          {mod === "commit" && (
            <Commit ask={ask} />
          )}

          {mod === "projects" && (
            <Projects />
          )}

          {mod === "waiting" && (
            <Waiting ask={ask} />
          )}

          {mod === "workplace" && (
            <Workplace />
          )}

          {mod === "rooms" && (
            <Rooms />
          )}

        </main>

        {/* ===================================================
            RIGHT PANEL
        =================================================== */}

        <Right set={setMod} />

      </div>

      {/* =====================================================
          MODALS
      ===================================================== */}

      {modal && (
        <Modal
          x={modal}
          close={() =>
            setModal(null)
          }
          send={doSend}
          draft={doDraft}
        />
      )}

      {/* =====================================================
          PROFILE
      ===================================================== */}

      {profileOpen && (
        <ProfileCard
          p={d.p}
          close={() =>
            setProfileOpen(false)
          }
          logout={logout}
        />
      )}

    </div>
  );
}

/* =========================================================
   LOGIN
========================================================= */

function Login({ onLogin }) {
  return (
    <div className="login">

      <div className="scan" />

      <div className="loginOrb">
        <div />
      </div>

      <div className="loginText">

        <small>
          AI EMPLOYEE EXPERIENCE
        </small>

        <h1>
          WORKDAY
          <br />
          <span>COPILOT</span>
        </h1>

        <p>
          Understand your work.
          Know what matters.
          Take the next best action.
        </p>

        <button onClick={onLogin}>
          Sign in with Microsoft{" "}
          <b>→</b>
        </button>

        <i>
          Secure identity ·
          Context-aware ·
          Human controlled
        </i>

      </div>

    </div>
  );
}

/* =========================================================
   NAV
========================================================= */

function Nav({
  a,
  s,
  id,
  t,
  i,
  n
}) {
  return (
    <button
      className={`nav ${
        a === id
          ? "on"
          : ""
      }`}
      onClick={() => s(id)}
    >
      <b>{i}</b>

      {t}

      {n !== undefined && (
        <small>{n}</small>
      )}
    </button>
  );
}

/* =========================================================
   ORB
========================================================= */

function Orb({
  mode = "idle",
  size = "normal"
}) {
  return (
    <div
      className={`orb ${mode} ${size}`}
    >
      <div className="orbAura" />

      <div className="orbRing ring1" />

      <div className="orbRing ring2" />

      <div className="orbRing ring3" />

      <div className="orbGrid" />

      <div className="orbCore" />

      <i className="p1" />
      <i className="p2" />
      <i className="p3" />
      <i className="p4" />
      <i className="p5" />
      <i className="p6" />
    </div>
  );
}

/* =========================================================
   HOME
========================================================= */

function Home({
  p,
  m,
  c,
  ask,
  set,
  suggestions
}) {
  const u =
    tasks.filter(
      (x) =>
        x[3] === "critical" ||
        x[3] === "high"
    ).length;

  return (
    <div>

      <div className="hero">

        <div>
          <small>
            YOUR WORKDAY, AT A GLANCE
          </small>

          <h2>
            Good afternoon,{" "}
            {p?.givenName ||
              "there"}
            .
          </h2>

          <p>
            I've connected the latest
            context. Here's what matters
            right now.
          </p>
        </div>

        <span>
          ● Context synced
        </span>

      </div>

      <div className="kpis">

        <K
          n={u}
          l="Urgent"
          s="needs attention"
          cl="red"
        />

        <K
          n={tasks.length}
          l="Actions"
          s="across your work"
          cl="amber"
        />

        <K
          n={c.length}
          l="Meetings"
          s="next 7 days"
          cl="cyan"
        />

        <K
          n={1}
          l="Waiting"
          s="dependency detected"
          cl="violet"
        />

      </div>

      <div className="two">

        <section className="panel">

          <Title
            k="AI PRIORITIZATION"
            t="Today's Action Plan"
          />

          {tasks.map((x) => (
            <div
              className="action"
              key={x[0]}
            >

              <span
                className={x[3]}
              />

              <div>

                <b>{x[0]}</b>

                <em>{x[1]}</em>

                <p>{x[4]}</p>

                <small>
                  ◷ {x[2]}
                </small>

              </div>

              <button
                onClick={() =>
                  ask(
                    `What should I do about ${x[0]}?`
                  )
                }
              >
                →
              </button>

            </div>
          ))}

        </section>

        <section className="panel center">

          <Orb />

          <small>
            READY
          </small>

          <h3>
            Ask your Workday Copilot
          </h3>

          <p>
            Voice or text. Try a command.
          </p>

          <div className="chips">

            {suggestions
              .slice(0, 4)
              .map((x) => (
                <button
                  onClick={() =>
                    ask(x)
                  }
                  key={x}
                >
                  {x} ↗
                </button>
              ))}

          </div>

          <button
            className="open"
            onClick={() =>
              set("copilot")
            }
          >
            Open full Copilot →
          </button>

        </section>

      </div>

      <div className="two lowergrid">

        <section className="panel">

          <Title
            k="INCOMING"
            t="Important mail"
          />

          {m
            .slice(0, 3)
            .map((x) => (
              <div
                className="line"
                key={x.id}
              >

                <span />

                <div>
                  <b>
                    {x.subject ||
                      "(No subject)"}
                  </b>

                  <small>
                    {x.bodyPreview}
                  </small>
                </div>

              </div>
            ))}

        </section>

        <section className="panel">

          <Title
            k="UP NEXT"
            t="Calendar"
          />

          {c
            .slice(0, 3)
            .map((x) => (
              <div
                className="line"
                key={x.id}
              >

                <strong>
                  {new Date(
                    x.start.dateTime
                  ).toLocaleTimeString(
                    [],
                    {
                      hour: "2-digit",
                      minute: "2-digit"
                    }
                  )}
                </strong>

                <div>
                  <b>
                    {x.subject}
                  </b>

                  <small>
                    {x.location
                      ?.displayName ||
                      "Online"}
                  </small>
                </div>

              </div>
            ))}

        </section>

      </div>

    </div>
  );
}

function K({
  n,
  l,
  s,
  cl
}) {
  return (
    <div className="kpi">

      <i className={cl} />

      <div>
        <b>{n}</b>

        <strong>{l}</strong>

        <small>{s}</small>
      </div>

    </div>
  );
}

function Title({ k, t }) {
  return (
    <div className="title">
      <div>
        <small>{k}</small>
        <h3>{t}</h3>
      </div>
    </div>
  );
}

/* =========================================================
   COPILOT
========================================================= */

function Copilot({
  msgs,
  q,
  setQ,
  ask,
  state,
  voice,
  stop,
  mute,
  muted,
  modal,
  suggestions,
  autoListen
}) {
  return (
    <div className="cp">

      <div className="cphead">

        <div>

          <small>
            CONVERSATIONAL WORK INTERFACE
          </small>

          <h2>
            Workday Copilot
          </h2>

          <p>
            Ask anything about your
            current work context.
          </p>

        </div>

        <span
          className={`status ${state}`}
        >

          <i />

          {state === "listening"
            ? "Listening"
            : state === "thinking"
            ? "Thinking"
            : state === "speaking"
            ? "AI speaking"
            : "Ready"}

        </span>

      </div>

      {/* =====================================================
          ORB STAGE
      ===================================================== */}

      <div className="stage">

        <Orb mode={state} />

        <div className="stageText">

          <strong>
            {state === "listening"
              ? "LISTENING"
              : state === "thinking"
              ? "ANALYZING YOUR WORK CONTEXT"
              : state === "speaking"
              ? "AI IS SPEAKING"
              : "READY FOR YOUR COMMAND"}
          </strong>

          <p>
            {state === "listening"
              ? 'Speak naturally. Say "stop" to cancel.'
              : state === "thinking"
              ? "Connecting mail, calendar and workplace context…"
              : state === "speaking"
              ? "You can stop the response at any time."
              : "Your workday context is ready."}
          </p>

        </div>

        <div className="voiceControl">

          <button
            className={
              state === "listening"
                ? "listeningBtn"
                : ""
            }
            onClick={voice}
          >

            <MicIcon
              size={15}
              muted={
                state !== "listening"
              }
            />

            {state === "listening"
              ? "Stop listening"
              : "Voice command"}

          </button>

          {state === "speaking" && (
            <button
              onClick={stop}
            >
              <StopIcon size={12} />
              Stop response
            </button>
          )}

          <button
            className={`stageMute ${
              muted ? "muted" : ""
            }`}
            onClick={mute}
            title={
              muted
                ? "Unmute AI voice"
                : "Mute AI voice"
            }
          >

            {muted ? (
              <SpeakerOffIcon
                size={15}
              />
            ) : (
              <SpeakerIcon
                size={15}
              />
            )}

            {muted
              ? "AI muted"
              : "AI voice"}

          </button>

          <label>

            <span
              className={
                autoListen
                  ? "toggle on"
                  : "toggle"
              }
            />

            <input
              type="checkbox"
              checked={autoListen}
              onChange={() =>
                {}
              }
              readOnly
            />

            Auto-listen after response

          </label>

        </div>

      </div>

      {/* =====================================================
          CHAT
      ===================================================== */}

      <div className="chat">

        {msgs.length === 0 && (
          <div className="empty">

            <h3>
              What can I help you with?
            </h3>

            <p>
              Try a command below. I’ll use
              your authorized work context
              and remember this conversation
              locally.
            </p>

          </div>
        )}

        {msgs.map((m, i) => (
          <div
            className={`msg ${m.r}`}
            key={i}
          >

            <span>
              {m.r === "u"
                ? "A"
                : "✦"}
            </span>

            <div>

              <div className="msgLabel">
                {m.r === "u"
                  ? "YOU"
                  : "WORKDAY COPILOT"}
              </div>

              {m.t
                .split("\n")
                .map((x, j) => (
                  <p key={j}>
                    {x || " "}
                  </p>
                ))}

              {m.actions?.length >
                0 && (
                <div className="actionButtons">

                  {m.actions.map(
                    (x, j) => (
                      <button
                        key={j}
                        onClick={() =>
                          modal(x)
                        }
                      >
                        {x.label ||
                          "Review action"}{" "}
                        →
                      </button>
                    )
                  )}

                </div>
              )}

            </div>

          </div>
        ))}

      </div>

      {/* =====================================================
          QUICK PROMPTS
      ===================================================== */}

      <div className="quick">

        {suggestions.map((x) => (
          <button
            key={x}
            onClick={() =>
              ask(x)
            }
          >
            {x}
          </button>
        ))}

      </div>

      {/* =====================================================
          COPILOT INPUT
          MUTE + MIC + INPUT + SEND
      ===================================================== */}

      <div className="input">

        {/* NEW MUTE BUTTON */}

        <button
          className={`inputIconButton muteButton ${
            muted ? "muted" : ""
          }`}
          onClick={mute}
          title={
            muted
              ? "Unmute AI voice"
              : "Mute AI voice"
          }
          aria-label={
            muted
              ? "Unmute AI voice"
              : "Mute AI voice"
          }
        >

          {muted ? (
            <SpeakerOffIcon size={18} />
          ) : (
            <SpeakerIcon size={18} />
          )}

        </button>

        {/* NEW PROFESSIONAL MIC BUTTON */}

        <button
          className={`inputIconButton micButton ${
            state === "listening"
              ? "micLive"
              : ""
          }`}
          onClick={voice}
          title={
            state === "listening"
              ? "Stop listening"
              : "Voice command"
          }
          aria-label={
            state === "listening"
              ? "Stop listening"
              : "Voice command"
          }
        >

          {state === "listening" ? (
            <StopIcon size={13} />
          ) : (
            <MicIcon size={19} />
          )}

        </button>

        <input
          value={q}
          onChange={(e) =>
            setQ(e.target.value)
          }
          onKeyDown={(e) =>
            e.key === "Enter" &&
            ask(q)
          }
          placeholder="Ask your workday copilot…"
        />

        <button
          className="sendInput"
          onClick={() =>
            ask(q)
          }
          title="Send"
          aria-label="Send"
        >
          <SendIcon size={17} />
        </button>

      </div>

    </div>
  );
}

/* =========================================================
   MAIL
========================================================= */

function Mail({
  m,
  ask,
  reply,
  flag,
  important
}) {
  return (
    <div className="module">

      <Title
        k="COMMUNICATION"
        t="Inbox intelligence"
      />

      <p className="sub">
        Real Microsoft Graph mail,
        with concise AI actions.
      </p>

      {m.map((x) => (
        <article
          className="mail"
          key={x.id}
        >

          <span>
            {(
              x.from
                ?.emailAddress
                ?.name || "?"
            )
              .slice(0, 1)
              .toUpperCase()}
          </span>

          <div>

            <div className="mailMeta">

              <b>
                {x.from
                  ?.emailAddress
                  ?.name ||
                  x.from
                    ?.emailAddress
                    ?.address ||
                  "Unknown sender"}
              </b>

              <small>
                {new Date(
                  x.receivedDateTime
                ).toLocaleString()}
              </small>

            </div>

            <h3>
              {x.subject ||
                "(No subject)"}
            </h3>

            <p>
              {x.bodyPreview ||
                "No preview available."}
            </p>

            <div className="mailActions">

              <button
                onClick={() =>
                  ask(
                    `Give me a short, crisp professional response for this email. Subject: ${
                      x.subject || ""
                    }. Body: ${
                      x.bodyPreview || ""
                    }`
                  )
                }
              >
                AI response
              </button>

              <button
                className="primaryAction"
                onClick={() =>
                  reply(x)
                }
              >
                Draft reply
              </button>

              <button
                onClick={() =>
                  flag(x)
                }
              >
                ☆ Flag
              </button>

              <button
                onClick={() =>
                  important(x)
                }
              >
                ! Important
              </button>

            </div>

          </div>

        </article>
      ))}

    </div>
  );
}

/* =========================================================
   CALENDAR
========================================================= */

function Calendar({ c }) {
  return (
    <div className="module">

      <Title
        k="TIME & FOCUS"
        t="Calendar intelligence"
      />

      <p className="sub">
        Your next 7 days from Microsoft
        Graph.
      </p>

      {c.map((x) => (
        <article
          className="cal"
          key={x.id}
        >

          <strong>
            {new Date(
              x.start.dateTime
            ).toLocaleDateString(
              [],
              {
                day: "2-digit"
              }
            )}
          </strong>

          <div>

            <small>
              {new Date(
                x.start.dateTime
              ).toLocaleTimeString(
                [],
                {
                  hour: "2-digit",
                  minute: "2-digit"
                }
              )}
            </small>

            <h3>
              {x.subject}
            </h3>

            <p>
              {x.location
                ?.displayName ||
                "No location"}
            </p>

          </div>

        </article>
      ))}

    </div>
  );
}

/* =========================================================
   COMMITMENTS
========================================================= */

function Commit({ ask }) {
  return (
    <div className="module">

      <Title
        k="MEMORY & FOLLOW-UP"
        t="Your commitments"
      />

      <p className="sub">
        Extracted from email +
        synthetic Teams context.
      </p>

      {tasks
        .slice(0, 3)
        .map((x) => (
          <article
            className="commit"
            key={x[0]}
          >

            <b>◎</b>

            <div>

              <h3>{x[0]}</h3>

              <small>
                {x[2]} · {x[1]}
              </small>

              <p>
                Detected as an
                action/commitment that
                needs follow-up.
              </p>

            </div>

            <button
              onClick={() =>
                ask(
                  `How should I follow up on ${x[0]}?`
                )
              }
            >
              Ask AI →
            </button>

          </article>
        ))}

    </div>
  );
}

/* =========================================================
   PROJECTS
========================================================= */

function Projects() {
  return (
    <div className="module">

      <Title
        k="PROJECT INTELLIGENCE"
        t="Project pulse"
      />

      <div className="cards">

        {projects.map((x) => (
          <article
            className="project"
            key={x[0]}
          >

            <i className={x[4]} />

            <h3>{x[0]}</h3>

            <span>{x[1]}</span>

            <div>

              <b>
                {x[2]}
                <small>
                  {" "}
                  Actions
                </small>
              </b>

              <b>
                {x[3]}
                <small>
                  {" "}
                  Blockers
                </small>
              </b>

            </div>

            <hr />

          </article>
        ))}

      </div>

    </div>
  );
}

/* =========================================================
   WAITING
========================================================= */

function Waiting({ ask }) {
  return (
    <div className="module">

      <Title
        k="DEPENDENCIES"
        t="What am I waiting for?"
      />

      <article className="waiting">

        <b>⌛</b>

        <div>

          <em>BLOCKING</em>

          <h3>
            Database credentials
          </h3>

          <p>
            Waiting for DevOps. This is
            blocking Atlas work.
          </p>

          <small>
            Last mentioned · Today
            11:08 AM
          </small>

        </div>

        <button
          onClick={() =>
            ask(
              "What should I do about the database credentials dependency?"
            )
          }
        >
          Prepare follow-up →
        </button>

      </article>

    </div>
  );
}

/* =========================================================
   WORKPLACE
========================================================= */

function Workplace() {
  return (
    <div className="module">

      <Title
        k="EMPLOYEE EXPERIENCE"
        t="Campus pulse"
      />

      <div className="work">

        <div className="mapbig">

          <div>
            CAPGEMINI
            <br />
            <small>HQ</small>
          </div>

          <span>●</span>
          <span>●</span>

        </div>

        <div>

          {news.map((x) => (
            <article
              className="news"
              key={x[1]}
            >

              <b>
                {x[0]} {x[1]}
              </b>

              <p>
                {x[2]}
              </p>

            </article>
          ))}

          {buses.map((x) => (
            <article
              className="bus"
              key={x[0]}
            >

              <b>{x[0]}</b>

              <span>{x[3]}</span>

              <small>
                {x[1]} · {x[2]}
              </small>

            </article>
          ))}

        </div>

      </div>

    </div>
  );
}

/* =========================================================
   ROOMS
========================================================= */

function Rooms() {
  return (
    <div className="module">

      <Title
        k="WORKPLACE"
        t="Meeting room availability"
      />

      <div className="cards">

        {rooms.map((x) => (
          <article
            className="room"
            key={x[0]}
          >

            <i
              className={
                x[3]
                  ? "green"
                  : "red"
              }
            />

            <h3>{x[0]}</h3>

            <p>
              {x[1]} · {x[2]} seats
            </p>

            <b
              className={
                x[3]
                  ? "green"
                  : "red"
              }
            >
              {x[3]
                ? "Available"
                : "Busy"}
            </b>

            <small>
              {x[3]
                ? `Available until ${x[4]}`
                : `Booked until ${x[4]}`}
            </small>

          </article>
        ))}

      </div>

    </div>
  );
}

/* =========================================================
   RIGHT PANEL
========================================================= */

function Right({ set }) {
  return (
    <div className="right">

      <small>
        WORKPLACE PULSE{" "}
        <em>● LIVE</em>
      </small>

      <section>

        <b>
          MEETING ROOMS

          <button
            onClick={() =>
              set("rooms")
            }
          >
            View all
          </button>

        </b>

        {rooms.map((x) => (
          <div
            className="rrow"
            key={x[0]}
          >

            <span>

              <strong>
                {x[0]}
              </strong>

              <small>
                {x[1]} · {x[2]} seats
              </small>

            </span>

            <em
              className={
                x[3]
                  ? "green"
                  : "red"
              }
            >
              {x[3]
                ? "Available"
                : "Busy"}
            </em>

          </div>
        ))}

      </section>

      <section>

        <b>
          CAMPUS UPDATES
        </b>

        {news.map((x) => (
          <div
            className="rnews"
            key={x[1]}
          >

            <span>
              {x[0]}
            </span>

            <p>

              <strong>
                {x[1]}
              </strong>

              <small>
                {x[2]}
              </small>

            </p>

          </div>
        ))}

      </section>

      <section>

        <b>
          CAMPUS MOBILITY{" "}
          <em>LIVE MAP</em>
        </b>

        <div className="minimap">

          <i>HQ</i>

          <i>GATE 2</i>

          <span>●</span>

        </div>

        {buses.map((x) => (
          <div
            className="rrow"
            key={x[0]}
          >

            <span>

              <strong>
                ▰ {x[0]}
              </strong>

              <small>
                {x[1]}
              </small>

            </span>

            <em>
              {x[3]}
            </em>

          </div>
        ))}

      </section>

    </div>
  );
}

/* =========================================================
   EMAIL MODAL
========================================================= */

function Modal({
  x,
  close,
  send,
  draft
}) {
  const [to, setTo] =
    useState(x.to || "");

  const [cc, setCc] =
    useState(x.cc || "");

  const [s, setS] =
    useState(x.subject || "");

  const [b, setB] =
    useState(x.body || "");

  const [saving, setSaving] =
    useState(false);

  const isReply =
    x.mode === "reply" ||
    x.label
      ?.toLowerCase()
      .includes("reply");

  async function save(mode) {
    setSaving(true);

    try {
      if (mode === "draft") {
        await draft({
          to,
          cc,
          subject: s,
          body: b
        });
      } else {
        await send({
          to,
          cc,
          subject: s,
          body: b
        });
      }
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="back">

      <div className="modal">

        <button
          className="modalClose"
          onClick={close}
        >
          ×
        </button>

        <small>
          HUMAN APPROVAL
        </small>

        <h2>
          {isReply
            ? "Reply to email"
            : "Confirm email action"}
        </h2>

        <p>
          ✦ AI-assisted action ·
          Review the recipient, CC,
          subject and message before
          saving or sending.
        </p>

        <div className="recipientRow">

          <label>
            Recipient

            <input
              value={to}
              onChange={(e) =>
                setTo(e.target.value)
              }
              placeholder="person@company.com"
            />
          </label>

          <label>
            CC

            <input
              value={cc}
              onChange={(e) =>
                setCc(e.target.value)
              }
              placeholder="Optional"
            />
          </label>

        </div>

        <label>
          Subject

          <input
            value={s}
            onChange={(e) =>
              setS(e.target.value)
            }
          />
        </label>

        <label>
          Message

          <textarea
            rows="10"
            value={b}
            onChange={(e) =>
              setB(e.target.value)
            }
          />
        </label>

        <div className="draftInfo">
          {draft &&
            "Saving creates an Outlook Draft you can edit later."}
        </div>

        <footer>

          <button onClick={close}>
            Cancel
          </button>

          <button
            disabled={saving}
            onClick={() =>
              save("draft")
            }
          >
            ▣ Save draft
          </button>

          <button
            className="sendBtn"
            disabled={saving}
            onClick={() =>
              save("send")
            }
          >
            ✦ Approve & Send
          </button>

        </footer>

      </div>

    </div>
  );
}

/* =========================================================
   PROFILE CARD
========================================================= */

function ProfileCard({
  p,
  close,
  logout
}) {
  const email =
    p?.mail ||
    p?.userPrincipalName ||
    "Not available";

  return (
    <div
      className="profileBack"
      onClick={close}
    >

      <div
        className="profileCard"
        onClick={(e) =>
          e.stopPropagation()
        }
      >

        <button
          className="modalClose"
          onClick={close}
        >
          ×
        </button>

        <div className="profileHero">

          <div className="profileAvatar">
            {(
              p?.displayName ||
              "A"
            )
              .slice(0, 1)
              .toUpperCase()}
          </div>

          <div>

            <small>
              AUTHORIZED EMPLOYEE PROFILE
            </small>

            <h2>
              {p?.displayName ||
                "Employee"}
            </h2>

            <p>
              {p?.jobTitle ||
                "Employee"}
            </p>

          </div>

        </div>

        <div className="profileGrid">

          <Info
            k="Work email"
            v={email}
          />

          <Info
            k="Department"
            v={p?.department}
          />

          <Info
            k="Office"
            v={p?.officeLocation}
          />

          <Info
            k="Preferred language"
            v={p?.preferredLanguage}
          />

          <Info
            k="Mobile"
            v={p?.mobilePhone}
          />

          <Info
            k="User ID"
            v={p?.id}
          />

        </div>

        <div className="profileNote">
          ✦ Profile information is
          fetched from Microsoft Graph{" "}
          <b>/me</b> for the signed-in
          employee.
        </div>

        {/* =================================================
            NEW PROFESSIONAL SIGN OUT
        ================================================= */}

        <div className="profileLogout">

          <div className="profileLogoutInfo">

            <div className="logoutIcon">
              <LogoutIcon size={18} />
            </div>

            <div>
              <strong>
                Sign out
              </strong>

              <small>
                End this Workday Copilot
                session
              </small>
            </div>

          </div>

          <button
            className="profileLogoutButton"
            onClick={logout}
            title="Sign out"
          >
            <span>
              Sign out
            </span>

            <span className="logoutArrow">
              →
            </span>
          </button>

        </div>

      </div>

    </div>
  );
}

function Info({ k, v }) {
  return (
    <div className="info">

      <small>{k}</small>

      <b>
        {v || "Not available"}
      </b>

    </div>
  );
}

/* =========================================================
   MSAL INITIALIZATION
========================================================= */

msalInstance
  .initialize()
  .then(() =>
    createRoot(
      document.getElementById("root")
    ).render(
      <MsalProvider
        instance={msalInstance}
      >
        <App />
      </MsalProvider>
    )
  );
