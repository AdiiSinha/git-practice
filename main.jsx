import React, { useEffect, useMemo, useRef, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "./auth";
import {
  getProfile,
  getMail,
  getCalendar,
  sendMail,
  createDraft,
  markImportant,
  flagMail,
  updateCalendarImportance
} from "./graph";
import { teams, tasks, rooms, buses, news, projects } from "./data";
import { ai, BASE_SUGGESTIONS, dynamicFallback, fallback } from "./services/ai";
import { cleanSpeech, chooseVoice } from "./services/speech";

import { Header } from "./components/layout/Header";
import { Sidebar } from "./components/layout/Sidebar";
import { Right } from "./components/layout/Right";

import { Home } from "./components/modules/Home";
import { Copilot } from "./components/modules/Copilot";
import { Mail } from "./components/modules/Mail";
import { Calendar } from "./components/modules/Calendar";
import { Commit } from "./components/modules/Commit";
import { Projects } from "./components/modules/Projects";
import { Waiting } from "./components/modules/Waiting";
import { Workplace } from "./components/modules/Workplace";
import { Rooms } from "./components/modules/Rooms";

import { Login } from "./components/modals/Login";
import { Splash } from "./components/modals/Splash";
import { Modal } from "./components/modals/Modal";
import { ProfileCard } from "./components/modals/ProfileCard";

export function App() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const [mod, setMod] = useState("home");
  const [d, setD] = useState({ p: null, m: [], c: [] });
  const [msgs, setMsgs] = useState([]);
  const [q, setQ] = useState("");
  const [state, setState] = useState("idle");
  const [toast, setToast] = useState("");
  const [modal, setModal] = useState(null);
  const [profileOpen, setProfileOpen] = useState(false);
  const [splash, setSplash] = useState(false);
  const [suggestions, setSuggestions] = useState(BASE_SUGGESTIONS);
  const [autoListen, setAutoListen] = useState(true);
  const [muted, setMuted] = useState(false);
  const [drafts, setDrafts] = useState(() => JSON.parse(localStorage.getItem("wdDrafts") || "[]"));
  const recognitionRef = useRef(null);
  const autoTimerRef = useRef(null);
  const mountedRef = useRef(true);

  useEffect(() => () => {
    mountedRef.current = false;
    clearTimeout(autoTimerRef.current);
    recognitionRef.current?.stop?.();
    window.speechSynthesis?.cancel();
  }, []);

  useEffect(() => {
    if (account && !instance.getActiveAccount()) instance.setActiveAccount(account);
  }, [account, instance]);

  useEffect(() => {
    if (!account) return;
    setSplash(true);
    const timer = setTimeout(() => setSplash(false), 1500);
    load();
    return () => clearTimeout(timer);
  }, [account]);

  async function token() {
    try {
      return (await instance.acquireTokenSilent({
        scopes: loginRequest.scopes,
        account: instance.getActiveAccount() || account
      })).accessToken;
    } catch {
      return (await instance.acquireTokenPopup({ scopes: loginRequest.scopes })).accessToken;
    }
  }

  async function load() {
    try {
      const t = await token();
      const [p, m, c] = await Promise.all([getProfile(t), getMail(t), getCalendar(t)]);
      if (mountedRef.current) setD({ p, m: m.value || [], c: c.value || [] });
    } catch (e) {
      setToast(`Graph data load failed: ${e.message}`);
    }
  }

  const ctx = useMemo(() => ({
    profile: d.p,
    emails: d.m,
    calendar: d.c,
    teams: teams.map(x => ({ person: x[0], project: x[1], text: x[2], priority: x[3] })),
    tasks: tasks.map(x => ({ title: x[0], project: x[1], due: x[2], priority: x[3], why: x[4] })),
    projects,
    workplace: { rooms, buses, news }
  }), [d]);

  function saveMemory(question, answer) {
    const old = JSON.parse(localStorage.getItem("wdmem") || "[]");
    localStorage.setItem("wdmem", JSON.stringify([
      ...old,
      { q: question, a: answer, at: new Date().toISOString() }
    ].slice(-40)));
  }

  function setDynamicSuggestions(items) {
    const clean = Array.from(new Set((items || []).filter(Boolean).map(String))).slice(0, 6);
    setSuggestions(clean.length ? clean : BASE_SUGGESTIONS);
  }

  async function ask(text) {
    if (!text.trim()) return;
    stopRecognition();
    clearTimeout(autoTimerRef.current);
    setQ("");
    setMod("copilot");
    setMsgs(x => [...x, { r: "u", t: text }]);
    setState("thinking");
    try {
      const memory = JSON.parse(localStorage.getItem("wdmem") || "[]");
      const z = await ai({ query: text, context: ctx, memory });
      const answer = String(z.answer || "").trim();
      setMsgs(x => [...x, { r: "a", t: answer, actions: z.actions || [] }]);
      saveMemory(text, answer);
      setDynamicSuggestions(z.suggestions);
      speak(answer, true);
    } catch (e) {
      const answer = fallback(text, d, teams.length);
      setMsgs(x => [...x, { r: "a", t: answer, actions: [] }]);
      saveMemory(text, answer);
      setDynamicSuggestions(dynamicFallback(text));
      setToast(e.message);
      speak(answer, true);
    }
  }

  function speak(text, continueListening = false) {
    if (muted) {
      if (continueListening && autoListen) {
        clearTimeout(autoTimerRef.current);
        autoTimerRef.current = setTimeout(() => startListening(true), 2300);
      }
      return;
    }
    if (!window.speechSynthesis) return;
    clearTimeout(autoTimerRef.current);
    window.speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(cleanSpeech(text));
    const voiceChoice = chooseVoice();
    if (voiceChoice) u.voice = voiceChoice;
    u.lang = "en-IN";
    u.rate = 0.94;
    u.pitch = 0.92;
    u.volume = 1;
    u.onstart = () => mountedRef.current && setState("speaking");
    u.onend = () => {
      if (!mountedRef.current) return;
      setState("idle");
      if (continueListening && autoListen) {
        autoTimerRef.current = setTimeout(() => startListening(true), 2300);
      }
    };
    u.onerror = () => {
      if (mountedRef.current) setState("idle");
    };
    window.speechSynthesis.speak(u);
  }

  function stopSpeaking() {
    window.speechSynthesis?.cancel();
    setState("idle");
  }

  function stopRecognition() {
    clearTimeout(autoTimerRef.current);
    try { recognitionRef.current?.stop?.(); } catch {}
    recognitionRef.current = null;
    if (state === "listening") setState("idle");
  }

  function startListening(isAuto = false) {
    const R = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!R) {
      setToast("Speech recognition is not supported in this browser. Use Chrome or Edge.");
      return;
    }
    if (state === "speaking") stopSpeaking();
    stopRecognition();
    const r = new R();
    r.lang = "en-IN";
    r.interimResults = true;
    r.continuous = false;
    r.onstart = () => mountedRef.current && setState("listening");
    r.onresult = e => {
      const finalText = Array.from(e.results).map(v => v[0].transcript).join(" ").trim();
      if (!finalText) return;
      if (/^(stop|stop listening|cancel|quiet)$/i.test(finalText)) {
        stopRecognition();
        setState("idle");
        return;
      }
      if (e.results[e.results.length - 1].isFinal) ask(finalText);
    };
    r.onerror = e => {
      if (mountedRef.current && e.error !== "aborted" && e.error !== "no-speech") setToast(`Voice input: ${e.error}`);
      if (mountedRef.current) setState("idle");
    };
    r.onend = () => {
      recognitionRef.current = null;
      if (mountedRef.current && state === "listening") setState("idle");
    };
    recognitionRef.current = r;
    try { r.start(); } catch { setState("idle"); }
    if (!isAuto) setAutoListen(true);
  }

  function toggleVoice() {
    if (state === "speaking") return stopSpeaking();
    if (state === "listening") return stopRecognition();
    startListening(false);
  }

  async function login() { await instance.loginPopup(loginRequest); }
  async function logout() {
    stopSpeaking();
    stopRecognition();
    await instance.logoutPopup();
  }

  async function doSend(x) {
    try {
      const t = await token();
      await sendMail(t, x);
      setModal(null);
      setToast(`Email sent successfully to ${x.to}`);
      load();
    } catch (e) { setToast(`Send failed: ${e.message}`); }
  }

  async function doDraft(x) {
    try {
      const t = await token();
      const saved = await createDraft(t, x);
      setDrafts(prev => {
        const next = [{ ...x, id: saved.id, savedAt: new Date().toISOString() }, ...prev].slice(0, 20);
        localStorage.setItem("wdDrafts", JSON.stringify(next));
        return next;
      });
      setModal(null);
      setToast("Draft saved to Outlook Drafts.");
    } catch (e) { setToast(`Draft save failed: ${e.message}`); }
  }

  async function doCalendarImportant(event, important) {
  try {
    const t = await token();
 
    await updateCalendarImportance(
      t,
      event.id,
      important
    );
 
    setToast(
      important
        ? `"${event.subject || "Meeting"}" marked important.`
        : `"${event.subject || "Meeting"}" removed from important.`
    );
 
    load();
 
  } catch (e) {
 
    setToast(
      `Calendar update failed: ${e.message}`
    );
  }
}

function openMeetingMail(event) {
 
  const organizer =
    event.organizer?.emailAddress?.address || "";
 
  const organizerName =
    event.organizer?.emailAddress?.name ||
    "there";
 
  setModal({
    mode: "meeting",
    to: organizer,
    cc: "",
    subject: `Regarding: ${event.subject || "Meeting"}`,
    body:
      `Hi ${organizerName},\n\n` +
      `I wanted to follow up regarding "${event.subject || "the meeting"}".\n\n` +
      `Regards,\n${d.p?.displayName || "Employee"}`,
    sourceMeeting: event
  });
}
 

  async function doFlag(mail) {
    try { await flagMail(await token(), mail.id); setToast("Mail flagged for follow-up."); load(); }
    catch (e) { setToast(`Flag failed: ${e.message}`); }
  }

  async function doImportant(mail) {
    try { await markImportant(await token(), mail.id); setToast("Mail marked important."); load(); }
    catch (e) { setToast(`Important action failed: ${e.message}`); }
  }

  function openReply(mail) {
    const sender = mail.from?.emailAddress?.address || "";
    setModal({
      mode: "reply",
      to: sender,
      cc: "",
      subject: `Re: ${mail.subject || ""}`,
      body: `Hi ${mail.from?.emailAddress?.name || "there"},\n\nThanks for the update. I’ll review this and get back to you shortly.\n\nRegards,\n${d.p?.displayName || "Employee"}`,
      sourceMail: mail
    });
  }

  if (!account) return <Login onLogin={login} />;
  if (splash) return <Splash p={d.p} account={account} />;

  const voiceLabel =
    state === "listening"
      ? "Listening"
      : state === "speaking"
      ? "AI speaking"
      : state === "thinking"
      ? "Thinking"
      : "Ready";

  return (
    <div className="app">
      <Header
        d={d}
        account={account}
        state={state}
        voiceLabel={voiceLabel}
        toggleVoice={toggleVoice}
        stopSpeaking={stopSpeaking}
        stopRecognition={stopRecognition}
        muted={muted}
        setMuted={setMuted}
        autoListen={autoListen}
        setAutoListen={setAutoListen}
        setProfileOpen={setProfileOpen}
      />

      <div className="layout">
        <Sidebar
          mod={mod}
          setMod={setMod}
          mailCount={d.m.length}
          calendarCount={d.c.length}
          logout={logout}
        />

        <main>
          {toast && (
            <div className="toast">
              <span>✦</span>
              {toast}
              <button onClick={() => setToast("")}>×</button>
            </div>
          )}
          {mod === "home" && (
            <Home
              p={d.p}
              m={d.m}
              c={d.c}
              ask={ask}
              set={setMod}
              suggestions={suggestions}
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
              modal={setModal}
              suggestions={suggestions}
              autoListen={autoListen}
              muted={muted}
              setMuted={setMuted}
            />
          )}
          {mod === "mail" && (
            <Mail
              m={d.m}
              ask={ask}
              reply={openReply}
              flag={doFlag}
              important={doImportant}
            />
          )}
          {mod === "calendar" && (
  <Calendar
    c={d.c}
    ask={ask}
    onMailOrganizer={openMeetingMail}
    onJoinMeeting={(event) => {
      const url =
        event.onlineMeeting?.joinUrl ||
        event.onlineMeeting?.joinWebUrl ||
        event.webLink;
 
      if (url) {
        window.open(
          url,
          "_blank",
          "noopener,noreferrer"
        );
      } else {
        setToast(
          "No online meeting link is available for this event."
        );
      }
    }}
    onImportant={doCalendarImportant}
  />
)}
 
          {mod === "commit" && <Commit ask={ask} />}
          {mod === "projects" && <Projects />}
          {mod === "waiting" && <Waiting ask={ask} />}
          {mod === "workplace" && <Workplace />}
          {mod === "rooms" && <Rooms />}
        </main>

        <Right set={setMod} />
      </div>

      {modal && (
        <Modal
          x={modal}
          close={() => setModal(null)}
          send={doSend}
          draft={doDraft}
        />
      )}
      {profileOpen && (
        <ProfileCard p={d.p} close={() => setProfileOpen(false)} />
      )}
    </div>
  );
}

export default App;
