// IGCSE Study Hub — Google Sheets-backed app logic
const KEY_THEME = "igcse_theme_v1";
const KEY_TOKEN = "igcse_google_token_v1";
const KEY_USERNAME = "igcse_google_username_v1";
const SUBJECT_KEYS = ["bio", "chem", "phy", "ict"];

function createSubjectProgress(){
  return { idx: 0, correct: 0, attempted: 0, answers: {} };
}

function normalizeSubjectProgress(p){
  if (!p || typeof p !== "object") p = createSubjectProgress();
  if (typeof p.idx !== "number") p.idx = 0;
  if (typeof p.correct !== "number") p.correct = 0;
  if (typeof p.attempted !== "number") p.attempted = 0;
  if (!p.answers || typeof p.answers !== "object") p.answers = {};
  return p;
}

function getStoredToken(){ return localStorage.getItem(KEY_TOKEN) || ""; }
function getStoredUsername(){ return localStorage.getItem(KEY_USERNAME) || ""; }
function saveAuth(username, token){
  localStorage.setItem(KEY_USERNAME, username);
  localStorage.setItem(KEY_TOKEN, token);
}
function clearAuth(){
  localStorage.removeItem(KEY_USERNAME);
  localStorage.removeItem(KEY_TOKEN);
}

function apiRequest(action, params = {}){
  return new Promise((resolve, reject) => {
    if(!GOOGLE_SHEETS_API_URL || GOOGLE_SHEETS_API_URL.includes("PASTE_YOUR")){
      reject(new Error("Google Sheets API URL is not configured yet."));
      return;
    }

    const callbackName = "studyHubCallback_" + Date.now() + "_" + Math.floor(Math.random() * 100000);
    const query = new URLSearchParams({ action, callback: callbackName, ...params });
    const script = document.createElement("script");
    const cleanup = () => {
      delete window[callbackName];
      script.remove();
    };

    window[callbackName] = (data) => {
      cleanup();
      if(data && data.ok === false) reject(new Error(data.error || "Request failed."));
      else resolve(data);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Could not connect to the Google Sheets database."));
    };
    script.src = GOOGLE_SHEETS_API_URL + "?" + query.toString();
    document.body.appendChild(script);

    setTimeout(() => {
      if(window[callbackName]){
        cleanup();
        reject(new Error("The database request timed out."));
      }
    }, 15000);
  });
}

async function hashPassword(password){
  const bytes = new TextEncoder().encode(password);
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, "0")).join("");
}

async function signup(username, password){
  const passwordHash = await hashPassword(password);
  const data = await apiRequest("signup", { username, passwordHash });
  saveAuth(data.username, data.token);
  return data;
}

async function login(username, password){
  const passwordHash = await hashPassword(password);
  const data = await apiRequest("login", { username, passwordHash });
  saveAuth(data.username, data.token);
  return data;
}

async function getSession(){
  const token = getStoredToken();
  const username = getStoredUsername();
  if(!token || !username) return null;
  try{
    const data = await apiRequest("session", { token });
    return { token, user: { id: username, username: data.username } };
  }catch(error){
    clearAuth();
    return null;
  }
}

async function getCurrentUser(){
  const session = await getSession();
  return session ? session.user : null;
}

async function logout(){
  const token = getStoredToken();
  if(token){
    try { await apiRequest("logout", { token }); } catch(_) {}
  }
  clearAuth();
}

async function ensureUserProgress(){
  // Progress rows are created only when a user answers a question.
  return true;
}

async function allProgress(){
  const data = await apiRequest("allProgress", { token: getStoredToken() });
  return data.progress;
}

async function progressForSubject(_userId, subjectKey){
  const all = await allProgress();
  return normalizeSubjectProgress(all[subjectKey] || createSubjectProgress());
}

async function updateProgress(_userId, subjectKey, questionIndex, selectedIndex, isCorrect, skipped = false){
  await apiRequest("saveAnswer", {
    token: getStoredToken(),
    subject: subjectKey,
    questionIndex: String(questionIndex),
    selectedIndex: selectedIndex === null ? "" : String(selectedIndex),
    isCorrect: String(Boolean(isCorrect)),
    skipped: String(Boolean(skipped))
  });
  return progressForSubject(null, subjectKey);
}

async function getWrongAnswers(_userId, subjectKey){
  const p = await progressForSubject(null, subjectKey);
  return Object.entries(p.answers)
    .filter(([, result]) => result && result.correct === false)
    .map(([index, result]) => ({ index: Number(index), ...result }));
}

async function resetSubjectProgress(_userId, subjectKey){
  await apiRequest("resetSubject", { token: getStoredToken(), subject: subjectKey });
}

function percentCorrectFromProgress(p){
  if(!p || p.attempted === 0) return 0;
  return Math.round((p.correct / p.attempted) * 100);
}

// ---------- Theme (local preference only) ----------
function applyTheme(theme){
  const chosen = theme === "light" ? "light" : "dark";
  document.documentElement.dataset.theme = chosen;
  localStorage.setItem(KEY_THEME, chosen);
  const toggle = document.getElementById("themeToggle");
  if(toggle){
    toggle.setAttribute("aria-label", chosen === "dark" ? "Switch to light mode" : "Switch to dark mode");
    toggle.innerHTML = chosen === "dark" ? "☀️ <span>Light</span>" : "🌙 <span>Dark</span>";
  }
}
function initTheme(){
  const saved = localStorage.getItem(KEY_THEME);
  const systemLight = window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches;
  applyTheme(saved || (systemLight ? "light" : "dark"));
}
function toggleTheme(){
  const current = document.documentElement.dataset.theme || "dark";
  applyTheme(current === "dark" ? "light" : "dark");
}
document.addEventListener("DOMContentLoaded", initTheme);
