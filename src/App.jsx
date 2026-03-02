import { useCallback, useEffect, useRef, useState } from "react";

const API_BASE = (import.meta.env.VITE_API_BASE_URL || "/avatar-service").replace(/\/+$/, "");
const TEACHER_NAME = "SP Sir";
const AI_HEALTH_CACHE_MS = 180000;
const RESPONSE_DELAY_MS = 5000;
const SPEECH_RECOGNITION_LANG = "hi-IN";

function normalizeVoiceCommand(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isStopCommand(text) {
  const cmd = normalizeVoiceCommand(text);
  if (!cmd) return false;
  const words = cmd.split(" ").filter(Boolean);
  if (cmd === "stop" || cmd === "stop now" || cmd === "pause" || cmd === "mute" || cmd === "silence") {
    return true;
  }
  if (cmd.startsWith("stop ") && words.length <= 5) return true;
  if (cmd === "stop avatar" || cmd === "stop talking" || cmd === "be quiet") return true;
  if (words.length <= 7 && (words.includes("stop") || words.includes("pause") || words.includes("quiet"))) {
    return true;
  }
  return false;
}

function inferMoodFromReply(text) {
  const t = String(text || "").toLowerCase();
  if (t.includes("i hear you") || t.includes("together") || t.includes("feel")) return "empathetic";
  if (t.includes("?") || t.includes("question")) return "curious";
  if (t.includes("let's") || t.includes("step") || t.includes("plan")) return "confident";
  return "neutral";
}

function pickMaleBrowserVoice() {
  const synth = window.speechSynthesis;
  if (!synth) return null;
  const voices = synth.getVoices() || [];
  if (!voices.length) return null;
  const maleHints = [/alex/i, /daniel/i, /david/i, /thomas/i, /james/i, /male/i];
  const english = voices.filter((v) => String(v.lang || "").toLowerCase().startsWith("en"));
  const pool = english.length ? english : voices;
  return pool.find((v) => maleHints.some((hint) => hint.test(v.name))) || pool[0] || null;
}

function parseInlineFormattedText(text, keyPrefix) {
  const source = String(text || "");
  const tokens = source.split(/(\*\*[^*]+\*\*|`[^`]+`)/g).filter(Boolean);
  return tokens.map((token, index) => {
    if (token.startsWith("**") && token.endsWith("**") && token.length > 4) {
      return <strong key={`${keyPrefix}-b-${index}`}>{token.slice(2, -2)}</strong>;
    }
    if (token.startsWith("`") && token.endsWith("`") && token.length > 2) {
      return <code key={`${keyPrefix}-c-${index}`}>{token.slice(1, -1)}</code>;
    }
    return token;
  });
}

function normalizeSirMessage(text) {
  let normalized = String(text || "").replace(/\r\n/g, "\n").replace(/[ \t]+\n/g, "\n").trim();
  if (!normalized) return "";
  normalized = normalized.replace(
    /\b(Concept|Formula|Step-by-step|Final Answer|Summary|Explanation|Answer|Example)\s*:/gi,
    "\n$1:"
  );
  normalized = normalized.replace(/\s+(Step\s*\d+\s*:)/gi, "\n$1");
  return normalized.trim();
}

function renderFormattedSirResponse(text) {
  const normalized = normalizeSirMessage(text);
  if (!normalized) return null;

  const blocks = normalized
    .split(/\n{2,}/)
    .map((chunk) => chunk.trim())
    .filter(Boolean);

  return (
    <div className="sir-formatted">
      {blocks.map((block, blockIndex) => {
        const lines = block.split("\n").map((line) => line.trim()).filter(Boolean);
        const isBullet = lines.length > 1 && lines.every((line) => /^[-*•]\s+/.test(line));
        const isOrdered = lines.length > 1 && lines.every((line) => /^\d+\.\s+/.test(line));

        if (isBullet) {
          return (
            <ul key={`sir-ul-${blockIndex}`}>
              {lines.map((line, lineIndex) => (
                <li key={`sir-ul-${blockIndex}-${lineIndex}`}>
                  {parseInlineFormattedText(line.replace(/^[-*•]\s+/, ""), `sir-ul-${blockIndex}-${lineIndex}`)}
                </li>
              ))}
            </ul>
          );
        }

        if (isOrdered) {
          return (
            <ol key={`sir-ol-${blockIndex}`}>
              {lines.map((line, lineIndex) => (
                <li key={`sir-ol-${blockIndex}-${lineIndex}`}>
                  {parseInlineFormattedText(line.replace(/^\d+\.\s+/, ""), `sir-ol-${blockIndex}-${lineIndex}`)}
                </li>
              ))}
            </ol>
          );
        }

        if (lines.length > 1 && lines.every((line) => /^[A-Za-z][A-Za-z \-]{1,34}:\s+/.test(line))) {
          return (
            <div key={`sir-labeled-${blockIndex}`}>
              {lines.map((line, lineIndex) => {
                const parts = line.split(":");
                const label = parts.shift() || "";
                const body = parts.join(":").trim();
                return (
                  <p key={`sir-labeled-${blockIndex}-${lineIndex}`}>
                    <strong>{label}:</strong> {parseInlineFormattedText(body, `sir-label-${blockIndex}-${lineIndex}`)}
                  </p>
                );
              })}
            </div>
          );
        }

        const labeled = block.match(/^([A-Za-z][A-Za-z \-]{1,34}):\s*(.+)$/s);
        if (labeled) {
          const label = labeled[1].trim();
          const body = labeled[2].trim();
          return (
            <p key={`sir-p-${blockIndex}`}>
              <strong>{label}:</strong> {parseInlineFormattedText(body, `sir-p-${blockIndex}`)}
            </p>
          );
        }

        const paragraph = lines.join(" ");
        return <p key={`sir-p-${blockIndex}`}>{parseInlineFormattedText(paragraph, `sir-p-${blockIndex}`)}</p>;
      })}
    </div>
  );
}

export default function App() {
  const audioRef = useRef(null);
  const fileInputRef = useRef(null);
  const studentImageInputRef = useRef(null);
  const mediaStreamRef = useRef(null);
  const speechRecognitionRef = useRef(null);
  const speechBufferRef = useRef("");
  const studentTranscriptFinalRef = useRef("");
  const studentTranscriptInterimRef = useRef("");
  const speechResponseTimerRef = useRef(null);
  const currentTurnAbortControllerRef = useRef(null);
  const pendingUserTextRef = useRef("");
  const userStopRequestedRef = useRef(false);
  const lastStopHandledAtRef = useRef(0);
  const aiLastCheckedAtRef = useRef(0);
  const aiCheckInFlightRef = useRef(null);
  const lastAssistantTextRef = useRef("");
  const toastTimerRef = useRef(null);

  const [userId, setUserId] = useState("");
  const [conversationId, setConversationId] = useState("");
  const [avatars, setAvatars] = useState([]);
  const [selectedAvatarId, setSelectedAvatarId] = useState("");
  const [selectedVoiceId, setSelectedVoiceId] = useState("alloy");
  const [callStatus, setCallStatus] = useState("idle");
  const [callActive, setCallActive] = useState(false);
  const [callStarting, setCallStarting] = useState(false);
  const [listeningActive, setListeningActive] = useState(false);
  const [muted, setMuted] = useState(false);
  const [processingTurn, setProcessingTurn] = useState(false);
  const [avatarSpeaking, setAvatarSpeaking] = useState(false);
  const [mood, setMood] = useState("neutral");
  const [liveCaption, setLiveCaption] = useState("Voice-only mode: avatar responses are spoken.");
  const [talkHint, setTalkHint] = useState("Continuous listening is off");
  const [studentTranscript, setStudentTranscript] = useState("Listening transcript will appear here...");
  const [conversationTranscript, setConversationTranscript] = useState([]);
  const [toast, setToast] = useState("");
  const [authToken, setAuthToken] = useState(() => window.localStorage.getItem("sp_sir_auth_token") || "");
  const [authUser, setAuthUser] = useState(null);
  const [authConfig, setAuthConfig] = useState({ googleClientIdConfigured: false, allowDevGoogleLogin: true });
  const [loginEmail, setLoginEmail] = useState("");
  const [loginName, setLoginName] = useState("");
  const [loggingIn, setLoggingIn] = useState(false);
  const [managedUsers, setManagedUsers] = useState([]);
  const [newUserEmail, setNewUserEmail] = useState("");
  const [newUserName, setNewUserName] = useState("");
  const [newUserRole, setNewUserRole] = useState("teacher");
  const [creatingUser, setCreatingUser] = useState(false);
  const [deletingUserId, setDeletingUserId] = useState("");
  const [chatInput, setChatInput] = useState("");
  const [trainingDocs, setTrainingDocs] = useState([]);
  const [trainingCounts, setTrainingCounts] = useState({ documents: 0, chunks: 0, vectors: 0, totalChars: 0 });
  const [uploadingTraining, setUploadingTraining] = useState(false);
  const [uploadingStudentImage, setUploadingStudentImage] = useState(false);
  const [studentImageHint, setStudentImageHint] = useState("");
  const [preferBrowserTts] = useState(false);
  const [browserTtsRate] = useState(0.84);
  const [audioPlaybackRate] = useState(0.92);
  const [useBrowserStt] = useState(Boolean(window.SpeechRecognition || window.webkitSpeechRecognition));

  const userIdRef = useRef(userId);
  const authTokenRef = useRef(authToken);
  const conversationIdRef = useRef(conversationId);
  const selectedAvatarIdRef = useRef(selectedAvatarId);
  const selectedVoiceIdRef = useRef(selectedVoiceId);
  const callActiveRef = useRef(callActive);
  const mutedRef = useRef(muted);
  const processingTurnRef = useRef(processingTurn);
  const listeningActiveRef = useRef(listeningActive);

  useEffect(() => {
    userIdRef.current = userId;
  }, [userId]);
  useEffect(() => {
    authTokenRef.current = authToken;
  }, [authToken]);
  useEffect(() => {
    conversationIdRef.current = conversationId;
  }, [conversationId]);
  useEffect(() => {
    selectedAvatarIdRef.current = selectedAvatarId;
  }, [selectedAvatarId]);
  useEffect(() => {
    selectedVoiceIdRef.current = selectedVoiceId;
  }, [selectedVoiceId]);
  useEffect(() => {
    callActiveRef.current = callActive;
  }, [callActive]);
  useEffect(() => {
    mutedRef.current = muted;
  }, [muted]);
  useEffect(() => {
    processingTurnRef.current = processingTurn;
  }, [processingTurn]);
  useEffect(() => {
    listeningActiveRef.current = listeningActive;
  }, [listeningActive]);

  const showToast = useCallback((message) => {
    setToast(message);
    if (toastTimerRef.current) {
      window.clearTimeout(toastTimerRef.current);
    }
    toastTimerRef.current = window.setTimeout(() => {
      setToast("");
      toastTimerRef.current = null;
    }, 2200);
  }, []);

  const api = useCallback(async (path, options = {}) => {
    const headers = { ...(options.headers || {}) };
    if (options.body != null && !(options.body instanceof FormData)) {
      headers["Content-Type"] = headers["Content-Type"] || "application/json";
    }
    if (!options.skipAuth) {
      const token = String(options.authToken || authTokenRef.current || "").trim();
      if (token) headers.Authorization = `Bearer ${token}`;
      if (userIdRef.current) headers["X-User-Id"] = userIdRef.current;
    }

    const response = await fetch(path, { ...options, headers });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(body.error || `Request failed (${response.status})`);
    }
    return body;
  }, []);

  const isAvatarBusy = useCallback(() => {
    return (
      processingTurnRef.current ||
      Boolean(audioRef.current && !audioRef.current.paused) ||
      Boolean(window.speechSynthesis?.speaking)
    );
  }, []);

  const shouldTriggerStop = useCallback(
    (text) => {
      const cmd = normalizeVoiceCommand(text);
      if (!cmd) return false;
      if (isStopCommand(cmd)) return true;
      if (isAvatarBusy() && /\b(stop|pause|quiet|silence)\b/.test(cmd)) return true;
      return false;
    },
    [isAvatarBusy]
  );

  const stopAvatarPlayback = useCallback(() => {
    try {
      window.speechSynthesis?.cancel();
    } catch (_err) {
      // no-op
    }
    const audio = audioRef.current;
    if (audio && !audio.paused) {
      audio.pause();
      audio.currentTime = 0;
    }
    setAvatarSpeaking(false);
    setMood("curious");
  }, []);

  const clearSpeechResponseTimer = useCallback(() => {
    if (speechResponseTimerRef.current) {
      window.clearTimeout(speechResponseTimerRef.current);
      speechResponseTimerRef.current = null;
    }
  }, []);

  const clearSpeechBuffer = useCallback(() => {
    speechBufferRef.current = "";
    studentTranscriptInterimRef.current = "";
    const finalText = studentTranscriptFinalRef.current.trim();
    setStudentTranscript(finalText || "Listening transcript will appear here...");
  }, []);

  const renderStudentTranscript = useCallback(() => {
    const finalText = studentTranscriptFinalRef.current.trim();
    const interimText = studentTranscriptInterimRef.current.trim();
    const combined = `${finalText} ${interimText}`.trim();
    setStudentTranscript(combined || "Listening transcript will appear here...");
  }, []);

  const appendConversationTranscript = useCallback((speaker, text) => {
    const clean = String(text || "").trim();
    if (!clean) return;
    setConversationTranscript((prev) => {
      const last = prev[prev.length - 1];
      if (last && last.speaker === speaker && last.text === clean) {
        return prev;
      }
      const next = [...prev, { id: `${Date.now()}-${Math.random()}`, speaker, text: clean }];
      return next.length > 80 ? next.slice(next.length - 80) : next;
    });
  }, []);

  const speakWithBrowserVoice = useCallback(
    (text, { signal } = {}) => {
      return new Promise((resolve, reject) => {
        const synth = window.speechSynthesis;
        if (!synth) {
          reject(new Error("browser_tts_unavailable"));
          return;
        }
        if (signal?.aborted) {
          resolve();
          return;
        }

        let finished = false;
        const finish = (fn) => {
          if (finished) return;
          finished = true;
          if (signal) signal.removeEventListener("abort", onAbort);
          fn();
        };

        const onAbort = () => {
          try {
            synth.cancel();
          } catch (_err) {
            // no-op
          }
          setAvatarSpeaking(false);
          finish(() => resolve());
        };

        if (signal) signal.addEventListener("abort", onAbort, { once: true });

        synth.cancel();
        const utter = new SpeechSynthesisUtterance(text);
        const maleVoice = pickMaleBrowserVoice();
        if (maleVoice) utter.voice = maleVoice;
        utter.rate = browserTtsRate;
        utter.pitch = 0.94;
        utter.volume = 1;

        utter.onstart = () => {
          if (signal?.aborted) {
            onAbort();
            return;
          }
          setAvatarSpeaking(true);
        };
        utter.onend = () => {
          setAvatarSpeaking(false);
          finish(() => resolve());
        };
        utter.onerror = () => {
          setAvatarSpeaking(false);
          if (signal?.aborted) {
            finish(() => resolve());
            return;
          }
          finish(() => reject(new Error("browser_tts_failed")));
        };
        synth.speak(utter);
      });
    },
    [browserTtsRate]
  );

  const synthesizeAndPlay = useCallback(
    async (text, { signal } = {}) => {
      if (mutedRef.current || !String(text || "").trim() || signal?.aborted || userStopRequestedRef.current) {
        setAvatarSpeaking(false);
        return;
      }

      lastAssistantTextRef.current = text;
      if (preferBrowserTts) {
        try {
          await speakWithBrowserVoice(text, { signal });
          return;
        } catch (_err) {
          // continue with API TTS
        }
      }

      try {
        const tts = await api(`${API_BASE}/voice/tts`, {
          method: "POST",
          signal,
          body: JSON.stringify({ text, voiceId: selectedVoiceIdRef.current || "alloy" }),
        });
        if (signal?.aborted || userStopRequestedRef.current) return;
        const audio = audioRef.current;
        if (!audio) return;
        audio.src = `data:${tts.mimeType};base64,${tts.audioBase64}`;
        audio.playbackRate = audioPlaybackRate;
        if ("preservesPitch" in audio) audio.preservesPitch = true;
        if ("webkitPreservesPitch" in audio) audio.webkitPreservesPitch = true;
        setAvatarSpeaking(true);
        await audio.play();
      } catch (err) {
        if (err?.name === "AbortError" || signal?.aborted || userStopRequestedRef.current) {
          return;
        }
        await speakWithBrowserVoice(text, { signal });
      }
    },
    [api, audioPlaybackRate, preferBrowserTts, speakWithBrowserVoice]
  );

  const isLikelyEcho = useCallback((text) => {
    const userText = String(text || "").toLowerCase().trim();
    const assistant = String(lastAssistantTextRef.current || "").toLowerCase().trim();
    if (!assistant || userText.length < 8) return false;
    return assistant.includes(userText) || userText.includes(assistant.slice(0, 30));
  }, []);

  const requestImmediateStop = useCallback(
    (source = "voice") => {
      const now = Date.now();
      if (now - lastStopHandledAtRef.current < 350) return;
      lastStopHandledAtRef.current = now;
      userStopRequestedRef.current = true;
      pendingUserTextRef.current = "";
      clearSpeechResponseTimer();
      clearSpeechBuffer();
      studentTranscriptInterimRef.current = "";
      renderStudentTranscript();
      if (currentTurnAbortControllerRef.current) {
        currentTurnAbortControllerRef.current.abort();
        currentTurnAbortControllerRef.current = null;
      }
      stopAvatarPlayback();
      setLiveCaption("Stopped. You can continue speaking.");
      setTalkHint("stopped immediately");
      if (source !== "system") {
        showToast("Stopped");
      }
    },
    [clearSpeechBuffer, clearSpeechResponseTimer, renderStudentTranscript, showToast, stopAvatarPlayback]
  );

  const restoreAuthSession = useCallback(async () => {
    const saved = window.localStorage.getItem("sp_sir_auth_token") || "";
    if (!saved) {
      setAuthToken("");
      setAuthUser(null);
      setUserId("");
      return null;
    }
    const me = await api(`${API_BASE}/auth/me`, { method: "GET", authToken: saved });
    setAuthToken(saved);
    setAuthUser(me.user);
    setUserId(me.user.id || "");
    return me.user;
  }, [api]);

  const logout = useCallback(() => {
    window.localStorage.removeItem("sp_sir_auth_token");
    setAuthToken("");
    setAuthUser(null);
    setUserId("");
    setConversationId("");
    setCallActive(false);
    callActiveRef.current = false;
    setCallStatus("idle");
    setListeningActive(false);
    const rec = speechRecognitionRef.current;
    if (rec) {
      rec.onend = null;
      rec.onerror = null;
      rec.onresult = null;
      rec.onspeechstart = null;
      try {
        rec.stop();
      } catch (_err) {
        // no-op
      }
      speechRecognitionRef.current = null;
    }
    clearSpeechResponseTimer();
    clearSpeechBuffer();
    setConversationTranscript([]);
    setStudentTranscript("Listening transcript will appear here...");
    setTalkHint("Continuous listening is off");
    setAvatars([]);
    setSelectedAvatarId("");
    setSelectedVoiceId("alloy");
    setStudentImageHint("");
  }, [clearSpeechBuffer, clearSpeechResponseTimer]);

  const loginWithGoogle = useCallback(async () => {
    const email = loginEmail.trim().toLowerCase();
    if (!email) {
      showToast("Enter email");
      return;
    }
    setLoggingIn(true);
    try {
      const res = await api(`${API_BASE}/auth/google`, {
        method: "POST",
        skipAuth: true,
        body: JSON.stringify({ email, name: loginName.trim() }),
      });
      const token = String(res.token || "").trim();
      if (!token) throw new Error("missing_auth_token");
      window.localStorage.setItem("sp_sir_auth_token", token);
      setAuthToken(token);
      setAuthUser(res.user || null);
      setUserId(res.user?.id || "");
      if (res.bootstrapAdmin) {
        showToast("Admin account bootstrapped");
      } else {
        showToast("Login successful");
      }
    } catch (err) {
      showToast(String(err?.message || err));
    } finally {
      setLoggingIn(false);
    }
  }, [api, loginEmail, loginName, showToast]);

  const fetchAuthConfig = useCallback(async () => {
    try {
      const config = await api(`${API_BASE}/auth/config`, { method: "GET", skipAuth: true });
      setAuthConfig({
        googleClientIdConfigured: Boolean(config.googleClientIdConfigured),
        allowDevGoogleLogin: Boolean(config.allowDevGoogleLogin),
      });
    } catch (_err) {
      setAuthConfig({ googleClientIdConfigured: false, allowDevGoogleLogin: true });
    }
  }, [api]);

  const loadManagedUsers = useCallback(async () => {
    const users = await api(`${API_BASE}/admin/users`, { method: "GET" });
    setManagedUsers(Array.isArray(users) ? users : []);
  }, [api]);

  const createManagedUser = useCallback(async () => {
    const email = newUserEmail.trim().toLowerCase();
    const name = newUserName.trim();
    if (!email) {
      showToast("User email required");
      return;
    }
    setCreatingUser(true);
    try {
      await api(`${API_BASE}/admin/users`, {
        method: "POST",
        body: JSON.stringify({ email, name, role: newUserRole }),
      });
      setNewUserEmail("");
      setNewUserName("");
      await loadManagedUsers();
      showToast("User created");
    } catch (err) {
      showToast(String(err?.message || err));
    } finally {
      setCreatingUser(false);
    }
  }, [api, loadManagedUsers, newUserEmail, newUserName, newUserRole, showToast]);

  const deleteManagedUser = useCallback(
    async (targetUser) => {
      const user = targetUser || {};
      const id = String(user.id || "").trim();
      if (!id) return;
      const role = String(user.role || "").trim().toLowerCase();
      if (role !== "teacher" && role !== "student") {
        showToast("Only teacher/student users can be removed");
        return;
      }
      setDeletingUserId(id);
      try {
        await api(`${API_BASE}/admin/users/${encodeURIComponent(id)}`, {
          method: "DELETE",
        });
        await loadManagedUsers();
        showToast("User removed");
      } catch (err) {
        showToast(String(err?.message || err));
      } finally {
        setDeletingUserId("");
      }
    },
    [api, loadManagedUsers, showToast]
  );

  const loadAvatars = useCallback(async () => {
    const list = await api(`${API_BASE}/avatars`, { method: "GET", skipAuth: true });
    setAvatars(list);
    if (!selectedAvatarIdRef.current && list.length) {
      setSelectedAvatarId(list[0].id);
      setSelectedVoiceId(list[0].voiceId || "alloy");
    }
    return list;
  }, [api]);

  const createConversation = useCallback(
    async (avatarId) => {
      const id = avatarId || selectedAvatarIdRef.current;
      if (!id) return "";
      const convo = await api(`${API_BASE}/conversations`, {
        method: "POST",
        body: JSON.stringify({ avatarId: id }),
      });
      setConversationId(convo.id);
      return convo.id;
    },
    [api]
  );

  const refreshTrainingPanel = useCallback(
    async (avatarId) => {
      const id = avatarId || selectedAvatarIdRef.current;
      if (!id) return;
      const encoded = encodeURIComponent(id);
      const [status, docs] = await Promise.all([
        api(`${API_BASE}/training/status?avatarId=${encoded}`, { method: "GET", skipAuth: true }),
        api(`${API_BASE}/training/documents?avatarId=${encoded}`, { method: "GET", skipAuth: true }),
      ]);
      setTrainingCounts({
        documents: Number(status.documents || 0),
        chunks: Number(status.chunks || 0),
        vectors: Number(status.vectors || 0),
        totalChars: Number(status.totalChars || 0),
      });
      setTrainingDocs(docs);
    },
    [api]
  );

  const ensureCallContext = useCallback(async () => {
    if (!conversationIdRef.current) {
      await createConversation();
    }
  }, [createConversation]);

  const ensureAiReady = useCallback(
    async ({ force = false } = {}) => {
      const now = Date.now();
      if (!force && now - aiLastCheckedAtRef.current < AI_HEALTH_CACHE_MS) {
        setCallStatus("ready");
        return;
      }
      if (aiCheckInFlightRef.current) {
        await aiCheckInFlightRef.current;
        return;
      }
      aiCheckInFlightRef.current = (async () => {
        const health = await api(`${API_BASE}/ai/health`, { method: "GET", skipAuth: true });
        if (!health.ok) {
          throw new Error(`AI unavailable: ${health.error || "unknown"}`);
        }
        aiLastCheckedAtRef.current = Date.now();
        setCallStatus("ready");
      })();
      try {
        await aiCheckInFlightRef.current;
      } finally {
        aiCheckInFlightRef.current = null;
      }
    },
    [api]
  );

  const sendTurnFallback = useCallback(
    async (userText, signal) => {
      if (!conversationIdRef.current) {
        await ensureCallContext();
      }
      try {
        return await api(`${API_BASE}/conversations/${conversationIdRef.current}/messages`, {
          method: "POST",
          signal,
          body: JSON.stringify({ text: userText }),
        });
      } catch (err) {
        const message = String(err?.message || err || "");
        if (err?.name === "AbortError") throw err;
        if (!message.includes("conversation_not_found")) {
          throw err;
        }
        await createConversation();
        return api(`${API_BASE}/conversations/${conversationIdRef.current}/messages`, {
          method: "POST",
          signal,
          body: JSON.stringify({ text: userText }),
        });
      }
    },
    [api, createConversation, ensureCallContext]
  );

  const runTurn = useCallback(
    async (userText) => {
      if (authUser?.role !== "student") {
        showToast("Only students can chat with SP Sir");
        return;
      }
      const clean = String(userText || "").trim();
      if (!clean) return;
      if (shouldTriggerStop(clean)) {
        requestImmediateStop("voice");
        return;
      }

      if (processingTurnRef.current) {
        pendingUserTextRef.current = clean;
        return;
      }

      appendConversationTranscript("Student", clean);

      const requestController = new AbortController();
      currentTurnAbortControllerRef.current = requestController;
      userStopRequestedRef.current = false;
      setProcessingTurn(true);
      setMood("thinking");
      setLiveCaption("SP Sir is thinking...");

      try {
        const turn = await sendTurnFallback(clean, requestController.signal);
        if (requestController.signal.aborted || userStopRequestedRef.current || !callActiveRef.current) {
          return;
        }
        const finalText = String(turn.assistantMessage?.text || "");
        appendConversationTranscript(TEACHER_NAME, finalText);
        setMood(inferMoodFromReply(finalText));
        setLiveCaption("SP Sir is speaking...");
        await synthesizeAndPlay(finalText, { signal: requestController.signal });
        if (requestController.signal.aborted || userStopRequestedRef.current || !callActiveRef.current) {
          return;
        }
        setLiveCaption("Voice-only mode: SP Sir responses are spoken.");
      } catch (err) {
        if (err?.name !== "AbortError") {
          showToast(String(err?.message || err));
        }
      } finally {
        if (currentTurnAbortControllerRef.current === requestController) {
          currentTurnAbortControllerRef.current = null;
        }
        setProcessingTurn(false);
        if (userStopRequestedRef.current) {
          userStopRequestedRef.current = false;
          return;
        }
        if (pendingUserTextRef.current) {
          const queued = pendingUserTextRef.current;
          pendingUserTextRef.current = "";
          await runTurn(queued);
        }
      }
    },
    [appendConversationTranscript, authUser?.role, requestImmediateStop, sendTurnFallback, shouldTriggerStop, showToast, synthesizeAndPlay]
  );

  const scheduleBufferedSpeechResponse = useCallback(() => {
    clearSpeechResponseTimer();
    const text = speechBufferRef.current.trim();
    if (!text) return;
    setTalkHint("Waiting 5s silence, then SP Sir will answer...");
    speechResponseTimerRef.current = window.setTimeout(async () => {
      speechResponseTimerRef.current = null;
      const finalText = speechBufferRef.current.trim();
      clearSpeechBuffer();
      if (!finalText || !callActiveRef.current || mutedRef.current) return;
      await runTurn(finalText);
    }, RESPONSE_DELAY_MS);
  }, [clearSpeechBuffer, clearSpeechResponseTimer, runTurn]);

  const stopContinuousListening = useCallback(() => {
    setListeningActive(false);
    clearSpeechResponseTimer();
    clearSpeechBuffer();
    const rec = speechRecognitionRef.current;
    if (rec) {
      rec.onend = null;
      rec.onerror = null;
      rec.onresult = null;
      rec.onspeechstart = null;
      try {
        rec.stop();
      } catch (_err) {
        // no-op
      }
      speechRecognitionRef.current = null;
    }
  }, [clearSpeechBuffer, clearSpeechResponseTimer]);

  const startRecognitionSession = useCallback(() => {
    const Ctor = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!Ctor) {
      throw new Error("SpeechRecognition not supported");
    }

    const recognition = new Ctor();
    recognition.lang = SPEECH_RECOGNITION_LANG;
    recognition.interimResults = true;
    recognition.continuous = true;

    recognition.onresult = (event) => {
      let finalText = "";
      let interimText = "";
      for (let i = event.resultIndex; i < event.results.length; i += 1) {
        const result = event.results[i];
        if (result.isFinal) {
          finalText += result[0].transcript;
        } else {
          interimText += result[0].transcript;
        }
      }

      const combined = `${interimText} ${finalText}`.trim();
      const interim = interimText.trim();
      const interimEcho = interim ? isLikelyEcho(interim) : false;
      const combinedEcho = combined ? isLikelyEcho(combined) : false;
      if (
        (interim && shouldTriggerStop(interim) && !interimEcho) ||
        (combined && shouldTriggerStop(combined) && !combinedEcho)
      ) {
        requestImmediateStop("voice");
        return;
      }
      const avatarBusy = isAvatarBusy();
      if (avatarBusy) {
        if (interim || finalText.trim()) {
          setTalkHint("SP Sir is speaking. Say STOP to interrupt.");
        }
        return;
      }

      if (interim) {
        studentTranscriptInterimRef.current = interim;
        renderStudentTranscript();
        if (speechBufferRef.current.trim()) {
          scheduleBufferedSpeechResponse();
        }
        setTalkHint("listening...");
      }

      const text = finalText.trim();
      if (!text || mutedRef.current || !callActiveRef.current) return;
      const echoText = isLikelyEcho(text);
      if (shouldTriggerStop(text) && !echoText) {
        requestImmediateStop("voice");
        return;
      }
      if (echoText) {
        return;
      }

      const existingTranscript = studentTranscriptFinalRef.current.trim();
      studentTranscriptFinalRef.current = existingTranscript ? `${existingTranscript} ${text}` : text;
      studentTranscriptInterimRef.current = "";
      renderStudentTranscript();

      const existing = speechBufferRef.current.trim();
      speechBufferRef.current = existing ? `${existing} ${text}` : text;
      scheduleBufferedSpeechResponse();
    };

    recognition.onspeechstart = () => {
      if (!callActiveRef.current || !listeningActiveRef.current) return;
      const avatarBusy = isAvatarBusy();
      if (!avatarBusy && !speechBufferRef.current.trim()) {
        studentTranscriptFinalRef.current = "";
        studentTranscriptInterimRef.current = "";
        renderStudentTranscript();
      }
      setTalkHint(avatarBusy ? "SP Sir is speaking. Say STOP to interrupt." : "you are speaking...");
    };

    recognition.onerror = () => {
      if (!callActiveRef.current || !listeningActiveRef.current) return;
      setTalkHint("listening recovered");
    };

    recognition.onend = () => {
      if (!callActiveRef.current || !listeningActiveRef.current) return;
      startRecognitionSession();
    };

    speechRecognitionRef.current = recognition;
    recognition.start();
  }, [isLikelyEcho, renderStudentTranscript, requestImmediateStop, scheduleBufferedSpeechResponse, shouldTriggerStop]);

  const startContinuousListening = useCallback(async () => {
    if (!callActiveRef.current) {
      showToast("Start call first");
      return;
    }
    if (!useBrowserStt) {
      showToast("SpeechRecognition unavailable");
      return;
    }
    if (listeningActiveRef.current) return;
    setListeningActive(true);
    studentTranscriptFinalRef.current = "";
    studentTranscriptInterimRef.current = "";
    renderStudentTranscript();
    setTalkHint("continuous listening on");
    startRecognitionSession();
  }, [renderStudentTranscript, showToast, startRecognitionSession, useBrowserStt]);

  const startCall = useCallback(async () => {
    if (authUser?.role !== "student") {
      showToast("Only students can start call");
      return;
    }
    if (callActiveRef.current || callStarting) return;
    setCallStarting(true);
    setCallStatus("starting...");
    let stream = null;
    try {
      const mediaPromise = navigator.mediaDevices.getUserMedia({
        audio: {
          echoCancellation: true,
          noiseSuppression: true,
          autoGainControl: true,
          channelCount: 1,
        },
      });
      const prepPromise = Promise.all([ensureCallContext(), ensureAiReady()]);
      [stream] = await Promise.all([mediaPromise, prepPromise]);
      mediaStreamRef.current = stream;
      setCallActive(true);
      callActiveRef.current = true;
      setConversationTranscript([]);
      setMood("neutral");
      setCallStatus("in call");
      showToast("Call started");
    } catch (err) {
      if (stream) stream.getTracks().forEach((t) => t.stop());
      setCallStatus("ready");
      throw err;
    } finally {
      setCallStarting(false);
    }

    if (!useBrowserStt) {
      setTalkHint("SpeechRecognition not supported in this browser");
      return;
    }
    await startContinuousListening();
  }, [authUser?.role, callStarting, ensureAiReady, ensureCallContext, showToast, startContinuousListening, useBrowserStt]);

  const endCall = useCallback(() => {
    setCallActive(false);
    callActiveRef.current = false;
    stopContinuousListening();
    clearSpeechResponseTimer();
    clearSpeechBuffer();
    studentTranscriptFinalRef.current = "";
    studentTranscriptInterimRef.current = "";
    setStudentTranscript("Listening transcript will appear here...");
    requestImmediateStop("system");
    if (mediaStreamRef.current) {
      mediaStreamRef.current.getTracks().forEach((t) => t.stop());
      mediaStreamRef.current = null;
    }
    setCallStatus("ended");
    setTalkHint("Continuous listening is off");
    setStudentImageHint("");
  }, [clearSpeechBuffer, clearSpeechResponseTimer, requestImmediateStop, stopContinuousListening]);

  const handleUploadTraining = useCallback(
    async (event) => {
      event.preventDefault();
      if (authUser?.role !== "teacher") {
        showToast("Only teachers can upload books");
        return;
      }
      const file = fileInputRef.current?.files?.[0];
      if (!selectedAvatarIdRef.current) {
        showToast("Choose avatar first");
        return;
      }
      if (!file) {
        showToast("Choose a file");
        return;
      }
      setUploadingTraining(true);
      try {
        const formData = new FormData();
        formData.append("avatarId", selectedAvatarIdRef.current);
        formData.append("filename", file.name);
        formData.append("file", file, file.name);
        const result = await api(`${API_BASE}/training/upload`, {
          method: "POST",
          body: formData,
        });
        fileInputRef.current.value = "";
        await refreshTrainingPanel();
        showToast(`Uploaded ${result.filename}`);
      } catch (err) {
        showToast(String(err?.message || err));
      } finally {
        setUploadingTraining(false);
      }
    },
    [api, authUser?.role, refreshTrainingPanel, showToast]
  );

  const handleClearTraining = useCallback(async () => {
    try {
      if (authUser?.role !== "teacher") {
        showToast("Only teachers can clear books");
        return;
      }
      if (!selectedAvatarIdRef.current) {
        showToast("Choose avatar first");
        return;
      }
      const avatarId = encodeURIComponent(selectedAvatarIdRef.current);
      const result = await api(`${API_BASE}/training/documents?avatarId=${avatarId}`, { method: "DELETE" });
      await refreshTrainingPanel();
      showToast(`Cleared ${result.deleted} document(s)`);
    } catch (err) {
      showToast(String(err?.message || err));
    }
  }, [api, authUser?.role, refreshTrainingPanel, showToast]);

  const handleUploadStudentImage = useCallback(
    async (event) => {
      event.preventDefault();
      if (authUser?.role !== "student") {
        showToast("Only students can upload image questions");
        return;
      }
      if (!callActiveRef.current) {
        showToast("Start Call first");
        return;
      }
      const file = studentImageInputRef.current?.files?.[0];
      if (!file) {
        showToast("Choose an image");
        return;
      }
      setUploadingStudentImage(true);
      try {
        if (!conversationIdRef.current) {
          await ensureCallContext();
        }
        const formData = new FormData();
        formData.append("filename", file.name);
        formData.append("file", file, file.name);
        const result = await api(`${API_BASE}/conversations/${conversationIdRef.current}/image`, {
          method: "POST",
          body: formData,
        });
        if (studentImageInputRef.current) {
          studentImageInputRef.current.value = "";
        }
        const hint = String(result.preview || "").trim();
        setStudentImageHint(hint || "Image parsed. Ask your question now.");
        showToast("Image parsed. Ask your question.");
      } catch (err) {
        showToast(String(err?.message || err));
      } finally {
        setUploadingStudentImage(false);
      }
    },
    [api, authUser?.role, ensureCallContext, showToast]
  );

  const handleToggleMute = useCallback(() => {
    setMuted((prev) => !prev);
  }, []);

  const handleTestVoice = useCallback(async () => {
    try {
      await synthesizeAndPlay("Hello. SP Sir voice test is active.");
      showToast("Voice test played");
    } catch (err) {
      showToast(`Voice failed: ${String(err?.message || err)}`);
    }
  }, [showToast, synthesizeAndPlay]);

  const handleSendTextMessage = useCallback(async (event) => {
    event.preventDefault();
    if (!callActiveRef.current) {
      showToast("Start Call first");
      return;
    }
    const text = chatInput.trim();
    if (!text) return;
    setChatInput("");
    await runTurn(text);
  }, [chatInput, runTurn, showToast]);

  useEffect(() => {
    const audio = audioRef.current;
    if (!audio) return undefined;

    const onPlay = () => setAvatarSpeaking(true);
    const onPause = () => setAvatarSpeaking(false);
    const onEnded = () => setAvatarSpeaking(false);
    const onError = () => setAvatarSpeaking(false);

    audio.addEventListener("play", onPlay);
    audio.addEventListener("pause", onPause);
    audio.addEventListener("ended", onEnded);
    audio.addEventListener("error", onError);
    return () => {
      audio.removeEventListener("play", onPlay);
      audio.removeEventListener("pause", onPause);
      audio.removeEventListener("ended", onEnded);
      audio.removeEventListener("error", onError);
    };
  }, []);

  useEffect(() => {
    (async () => {
      await fetchAuthConfig();
      try {
        await restoreAuthSession();
      } catch (_err) {
        window.localStorage.removeItem("sp_sir_auth_token");
        setAuthToken("");
        setAuthUser(null);
        setUserId("");
      }
    })();
  }, [fetchAuthConfig, restoreAuthSession]);

  useEffect(() => {
    if (!authUser?.id) return;
    (async () => {
      try {
        if (authUser.role === "admin") {
          await loadManagedUsers();
          setCallStatus("ready");
          return;
        }
        const list = await loadAvatars();
        const avatarId = selectedAvatarIdRef.current || list[0]?.id;
        if (!avatarId) return;
        if (authUser.role === "teacher") {
          await refreshTrainingPanel(avatarId);
          setCallStatus("ready");
          return;
        }
        if (!conversationIdRef.current) {
          await createConversation(avatarId);
        }
        await ensureAiReady();
        setCallStatus("ready");
      } catch (err) {
        showToast(`Boot failed: ${String(err?.message || err)}`);
      }
    })();
  }, [authUser?.id, authUser?.role, createConversation, ensureAiReady, loadAvatars, loadManagedUsers, refreshTrainingPanel, showToast]);

  useEffect(() => {
    return () => {
      clearSpeechResponseTimer();
      if (currentTurnAbortControllerRef.current) {
        currentTurnAbortControllerRef.current.abort();
      }
      if (mediaStreamRef.current) {
        mediaStreamRef.current.getTracks().forEach((t) => t.stop());
      }
      stopContinuousListening();
      try {
        window.speechSynthesis?.cancel();
      } catch (_err) {
        // no-op
      }
      if (toastTimerRef.current) {
        window.clearTimeout(toastTimerRef.current);
      }
    };
  }, [clearSpeechResponseTimer, stopContinuousListening]);

  const avatarName = avatars.find((a) => a.id === selectedAvatarId)?.name || TEACHER_NAME;
  const userRole = String(authUser?.role || "").toLowerCase();
  const endCallDisabled = !callActive;

  if (!authUser) {
    return (
      <div className="app-shell">
        <div className="noise" />
        <main className="layout login-layout">
          <section className="card panel auth-panel">
            <div className="brand">
              <div className="logo-badge" aria-hidden="true">ABC</div>
              <div>
                <h1>SP Sir</h1>
                <p>Google Login</p>
              </div>
            </div>
            <p className="note">
              Sign in with your Google email. First login bootstraps Admin, then Admin provisions Teacher/Student.
            </p>
            <form
              className="stack"
              onSubmit={(event) => {
                event.preventDefault();
                loginWithGoogle().catch((err) => showToast(String(err?.message || err)));
              }}
            >
              <input
                type="email"
                placeholder="you@school.edu"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
                required
              />
              <input
                type="text"
                placeholder="Full name (optional)"
                value={loginName}
                onChange={(e) => setLoginName(e.target.value)}
              />
              <button type="submit" disabled={loggingIn}>
                {loggingIn ? "Signing in..." : "Continue with Google"}
              </button>
            </form>
            {!authConfig.allowDevGoogleLogin && (
              <p className="note">Dev email-only login is disabled. Use valid Google credential integration.</p>
            )}
          </section>
        </main>
        <div className={`toast ${toast ? "show" : ""}`}>{toast}</div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <div className="noise" />
      <main className="layout">
        <section className="stage card">
          <header className="stage-head">
            <div className="brand">
              <div className="logo-badge" aria-hidden="true">ABC</div>
              <div>
                <h1>SP Sir</h1>
                <p>
                  {authUser.name || authUser.email} | {userRole || "user"}
                </p>
              </div>
            </div>
            <div className="row-head">
              <code>{callStatus}</code>
              <button className="ghost" onClick={logout}>Logout</button>
            </div>
          </header>
          {userRole === "admin" && (
            <section className="panel admin-panel">
              <div className="row-head">
                <h2>User Management</h2>
                <button className="ghost" onClick={() => loadManagedUsers().catch((err) => showToast(String(err?.message || err)))}>
                  Refresh
                </button>
              </div>
              <form
                className="stack"
                onSubmit={(event) => {
                  event.preventDefault();
                  createManagedUser().catch((err) => showToast(String(err?.message || err)));
                }}
              >
                <input
                  type="email"
                  placeholder="user@school.edu"
                  value={newUserEmail}
                  onChange={(e) => setNewUserEmail(e.target.value)}
                  required
                />
                <input
                  type="text"
                  placeholder="Full name"
                  value={newUserName}
                  onChange={(e) => setNewUserName(e.target.value)}
                />
                <select value={newUserRole} onChange={(e) => setNewUserRole(e.target.value)}>
                  <option value="teacher">Teacher</option>
                  <option value="student">Student</option>
                </select>
                <button type="submit" disabled={creatingUser}>{creatingUser ? "Creating..." : "Add User"}</button>
              </form>
              <div className="doc-list">
                {managedUsers.length === 0 ? (
                  <div className="empty">No users yet.</div>
                ) : (
                  managedUsers.map((u) => (
                    <div key={u.id} className="doc-item user-item">
                      <div className="user-meta">
                        <span>{u.name || u.email}</span>
                        <span className={`role-tag ${String(u.role || "").toLowerCase()}`}>{String(u.role || "").toUpperCase()}</span>
                      </div>
                      <div className="user-actions">
                        <button
                          type="button"
                          className="ghost"
                          disabled={deletingUserId === u.id || !["teacher", "student"].includes(String(u.role || "").toLowerCase())}
                          onClick={() => deleteManagedUser(u)}
                        >
                          {deletingUserId === u.id ? "Removing..." : "Remove"}
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </section>
          )}

          {userRole === "teacher" && (
            <section className="panel">
              <div className="row-head">
                <h2>Teacher Training</h2>
                <button className="ghost" onClick={() => refreshTrainingPanel().catch((err) => showToast(String(err?.message || err)))}>
                  Refresh
                </button>
              </div>
              <form className="stack" onSubmit={handleUploadTraining}>
                <input ref={fileInputRef} type="file" accept=".pdf,.txt,.md,.rtf,.doc,.docx" required />
                <button type="submit" disabled={uploadingTraining}>{uploadingTraining ? "Uploading..." : "Upload Book"}</button>
              </form>
              <p className="note">Only Teacher can upload/update books for SP Sir.</p>
              <div className="line"><span>Documents</span><code>{trainingCounts.documents}</code></div>
              <div className="line"><span>Chunks</span><code>{trainingCounts.chunks}</code></div>
              <div className="line"><span>Vectors</span><code>{trainingCounts.vectors}</code></div>
              <div className="line"><span>Characters</span><code>{trainingCounts.totalChars}</code></div>
              <div className="row-head">
                <h3>Uploaded</h3>
                <button className="ghost" type="button" onClick={handleClearTraining}>Clear</button>
              </div>
              <div className="doc-list">
                {trainingDocs.length === 0 ? (
                  <div className="empty">No books uploaded yet.</div>
                ) : (
                  trainingDocs.map((doc) => (
                    <div key={doc.id} className="doc-item">
                      {doc.filename} ({doc.characters} chars)
                    </div>
                  ))
                )}
              </div>
            </section>
          )}

          {userRole === "student" && (
            <>
              <div className="avatar-frame">
                <div className={`avatar-rig mood-${mood} ${avatarSpeaking ? "speaking" : ""}`}>
                  <div className="head" />
                  <div className="body" />
                  <div className="pulse" />
                </div>
                <div className="caption">{liveCaption}</div>
              </div>

              <div className="controls">
                <button onClick={() => startCall().catch((err) => showToast(String(err?.message || err)))} disabled={callActive || callStarting}>
                  {callStarting ? "Starting..." : "Start Call"}
                </button>
                <button className="ghost" onClick={endCall} disabled={endCallDisabled}>
                  End Call
                </button>
                <button className="ghost" onClick={handleToggleMute}>{muted ? "Unmute Mic" : "Mute Mic"}</button>
                <button className="ghost" onClick={handleTestVoice}>Test Voice</button>
              </div>

              <div className="listen-row">
                <span>{talkHint}</span>
              </div>

              <div className="student-transcript">
                <div className="label">Image Question</div>
                <form className="image-upload-row" onSubmit={handleUploadStudentImage}>
                  <input type="file" ref={studentImageInputRef} accept="image/*" disabled={endCallDisabled || uploadingStudentImage} />
                  <button type="submit" disabled={endCallDisabled || uploadingStudentImage}>
                    {uploadingStudentImage ? "Parsing..." : "Upload Image"}
                  </button>
                </form>
                <div className="value muted">Upload an image, then ask question by voice or chat.</div>
                {studentImageHint && (
                  <div className="value live"><span className="speaker">Image:</span> {studentImageHint}</div>
                )}
                <div className="label">Live Speech-to-Text</div>
                <div className="value live"><span className="speaker">Student:</span> {studentTranscript}</div>
                <div className="conversation-list whatsapp">
                  {conversationTranscript.length === 0 ? (
                    <div className="value muted">No conversation yet.</div>
                  ) : (
                    conversationTranscript.map((line) => (
                      <div key={line.id} className={`value transcript-line bubble ${line.speaker === TEACHER_NAME ? "sir" : "student"}`}>
                        <span className="speaker">{line.speaker}:</span>{" "}
                        {line.speaker === TEACHER_NAME ? (
                          <div className="sir-response">{renderFormattedSirResponse(line.text)}</div>
                        ) : (
                          line.text
                        )}
                      </div>
                    ))
                  )}
                </div>
                <form className="chat-input-row" onSubmit={handleSendTextMessage}>
                  <input
                    type="text"
                    placeholder="Type your question..."
                    value={chatInput}
                    onChange={(e) => setChatInput(e.target.value)}
                    disabled={endCallDisabled}
                  />
                  <button type="submit" disabled={endCallDisabled}>Send</button>
                </form>
              </div>

              <audio ref={audioRef} controls />
            </>
          )}
        </section>

        <aside className="side">
          {userRole === "student" && (
            <section className="card panel">
              <h2>Session</h2>
              <div className="line"><span>User</span><code>{authUser.email}</code></div>
              <div className="line"><span>Role</span><code>{userRole}</code></div>
              <div className="line"><span>Avatar</span><code>{avatarName}</code></div>
              <div className="line"><span>Conversation</span><code>{conversationId || "not started"}</code></div>
            </section>
          )}
        </aside>
      </main>

      <div className={`toast ${toast ? "show" : ""}`}>{toast}</div>
    </div>
  );
}
