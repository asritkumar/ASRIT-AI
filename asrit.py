# ============================================================
#                         ASRIT AI
#              Personal AI Assistant by ASRIT
# ============================================================

import os
import io
import json
import base64
import sqlite3
import mimetypes
import tempfile
import asyncio
import websockets
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from google import genai
from google.genai import types

from PIL import Image
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

TEXT_MODEL = "gemini-3.5-flash-lite"
LIVE_MODEL = "gemini-3.8-live"

ASRIT_SYSTEM_PROMPT = """
You are ASRIT.
Your name is Asrit.
You are a personal AI assistant created by Asrit.

If the user asks:
"What is your name?"
"Who are you?"
"What's your name?"
respond naturally that your name is Asrit.

You are a highly capable multimodal AI assistant.
"""

# ============================================================
# DATABASE
# ============================================================

DATABASE_FILE = "asrit.db"
UPLOAD_FOLDER = Path("asrit_uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)


def database():
    connection = sqlite3.connect(DATABASE_FILE)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    db = database()
    db.execute("""
        CREATE TABLE IF NOT EXISTS conversations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    db.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conversation_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(conversation_id) REFERENCES conversations(id)
        )
    """)
    db.execute("""
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conversation_id INTEGER,
            filename TEXT,
            mime_type TEXT,
            path TEXT,
            created_at TEXT NOT NULL
        )
    """)
    db.commit()
    db.close()


initialize_database()

if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)
else:
    client = None

app = FastAPI(title="ASRIT AI")

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def now():
    return datetime.now().isoformat(timespec="seconds")

def create_conversation(title="New conversation"):
    db = database()
    timestamp = now()
    cursor = db.execute(
        "INSERT INTO conversations (title, created_at, updated_at) VALUES (?, ?, ?)",
        (title, timestamp, timestamp)
    )
    conversation_id = cursor.lastrowid
    db.commit()
    db.close()
    return conversation_id

def add_message(conversation_id, role, content):
    db = database()
    db.execute(
        "INSERT INTO messages (conversation_id, role, content, created_at) VALUES (?, ?, ?, ?)",
        (conversation_id, role, content, now())
    )
    db.execute(
        "UPDATE conversations SET updated_at = ? WHERE id = ?",
        (now(), conversation_id)
    )
    db.commit()
    db.close()

def get_messages(conversation_id):
    db = database()
    rows = db.execute(
        "SELECT role, content, created_at FROM messages WHERE conversation_id = ? ORDER BY id ASC",
        (conversation_id,)
    ).fetchall()
    db.close()
    return [{"role": r["role"], "content": r["content"], "created_at": r["created_at"]} for r in rows]

def conversation_title_from_text(text):
    clean = " ".join(text.strip().split())
    if not clean:
        return "New conversation"
    return clean[:45] + "..." if len(clean) > 45 else clean

# ============================================================
# FILE EXTRACTION
# ============================================================

def extract_pdf(path):
    reader = PdfReader(path)
    text = ""
    for page in reader.pages:
        try:
            text += page.extract_text() or ""
        except Exception:
            pass
    return text

def extract_docx(path):
    document = Document(path)
    return "\n".join([p.text for p in document.paragraphs if p.text.strip()])

def extract_xlsx(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    output = []
    for sheet in workbook.worksheets:
        output.append(f"\n--- SHEET: {sheet.title} ---\n")
        for row in sheet.iter_rows(values_only=True):
            output.append(" | ".join(["" if v is None else str(v) for v in row]))
    return "\n".join(output)

def extract_text_file(path):
    for encoding in ["utf-8", "utf-16", "latin-1"]:
        try:
            with open(path, "r", encoding=encoding, errors="ignore") as f:
                return f.read()
        except Exception:
            pass
    return ""

def extract_file_text(path, mime_type, filename):
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".docx":
        return extract_docx(path)
    if suffix in [".xlsx", ".xlsm"]:
        return extract_xlsx(path)
    if suffix in [".txt", ".csv", ".md", ".json", ".py", ".html", ".css", ".js"]:
        return extract_text_file(path)
    return ""

# ============================================================
# GEMINI API INTERACTION
# ============================================================

async def ask_gemini(user_message, conversation_id=None, extra_context=""):
    if client is None:
        return "ASRIT is not connected to Gemini yet. Configure your GEMINI_API_KEY environment variable."

    history = ""
    if conversation_id:
        messages = get_messages(conversation_id)
        for message in messages[-20:]:
            role = "USER" if message["role"] == "user" else "ASRIT"
            history += f"\n{role}: {message['content']}\n"

    prompt = f"{ASRIT_SYSTEM_PROMPT}\n\nPrevious conversation:\n{history}\n\nUploaded info:\n{extra_context}\n\nUser: {user_message}\nAnswer as ASRIT."

    try:
        response = await asyncio.to_thread(
            client.models.generate_content,
            model=TEXT_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.7)
        )
        return response.text or "I couldn't generate a response."
    except Exception as error:
        return f"ASRIT encountered an error while contacting Gemini:\n{error}"

async def analyze_image(image_bytes, user_message, conversation_id=None):
    if client is None:
        return "Please configure your Gemini API key first."
    try:
        image = Image.open(io.BytesIO(image_bytes))
        image.load()
        prompt = f"{ASRIT_SYSTEM_PROMPT}\nAnalyze the provided image.\nUser question: {user_message}"
        response = await asyncio.to_thread(
            client.models.generate_content,
            model=TEXT_MODEL,
            contents=[prompt, image]
        )
        return response.text or "I couldn't analyze the image."
    except Exception as error:
        return f"ASRIT could not analyze this image:\n{error}"

async def analyze_uploaded_file(path, filename, mime_type, question):
    if client is None:
        return "Please configure your Gemini API key first."

    suffix = Path(filename).suffix.lower()
    if suffix in [".jpg", ".jpeg", ".png", ".webp"]:
        with open(path, "rb") as file:
            return await analyze_image(file.read(), question)

    extracted = extract_file_text(path, mime_type, filename)
    if extracted:
        extracted = extracted[:120000]
        prompt = f"{ASRIT_SYSTEM_PROMPT}\nUploaded File: {filename}\nContent:\n{extracted}\nQuestion: {question}"
        try:
            response = await asyncio.to_thread(client.models.generate_content, model=TEXT_MODEL, contents=prompt)
            return response.text or "No answer generated."
        except Exception as error:
            return f"Error analyzing document: {error}"

    if suffix in [".mp3", ".wav", ".m4a", ".aac", ".mp4", ".mov", ".webm"]:
        try:
            uploaded = await asyncio.to_thread(client.files.upload, file=path)
            prompt = f"{ASRIT_SYSTEM_PROMPT}\nUploaded Media: {filename}\nQuestion: {question}"
            response = await asyncio.to_thread(client.models.generate_content, model=TEXT_MODEL, contents=[prompt, uploaded])
            return response.text or "No answer generated."
        except Exception as error:
            return f"ASRIT could not process media: {error}"

    return f"I received `{filename}` and saved it, but direct inspection is unsupported for this format."

# ============================================================
# API ENDPOINTS
# ============================================================

@app.post("/api/chat")
async def chat(message: str = Form(...), conversation_id: int = Form(0)):
    if conversation_id == 0:
        conversation_id = create_conversation(conversation_title_from_text(message))
    add_message(conversation_id, "user", message)
    answer = await ask_gemini(message, conversation_id)
    add_message(conversation_id, "assistant", answer)
    return {"conversation_id": conversation_id, "answer": answer}

@app.post("/api/image")
async def image_api(image: UploadFile = File(...), message: str = Form("What is in this image?"), conversation_id: int = Form(0)):
    if conversation_id == 0:
        conversation_id = create_conversation("Image conversation")
    image_bytes = await image.read()
    answer = await analyze_image(image_bytes, message, conversation_id)
    add_message(conversation_id, "user", "[Image] " + message)
    add_message(conversation_id, "assistant", answer)
    return {"conversation_id": conversation_id, "answer": answer}

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...), message: str = Form("Analyze this file."), conversation_id: int = Form(0)):
    if conversation_id == 0:
        conversation_id = create_conversation(file.filename)
    safe_name = datetime.now().strftime("%Y%m%d_%H%M%S_") + Path(file.filename).name
    save_path = UPLOAD_FOLDER / safe_name

    content = await file.read()
    with open(save_path, "wb") as output:
        output.write(content)

    mime_type = file.content_type or mimetypes.guess_type(file.filename)[0] or "application/octet-stream"

    db = database()
    db.execute(
        "INSERT INTO files (conversation_id, filename, mime_type, path, created_at) VALUES (?, ?, ?, ?, ?)",
        (conversation_id, file.filename, mime_type, str(save_path), now())
    )
    db.commit()
    db.close()

    answer = await analyze_uploaded_file(str(save_path), file.filename, mime_type, message)
    add_message(conversation_id, "user", f"[Uploaded file: {file.filename}] {message}")
    add_message(conversation_id, "assistant", answer)
    return {"conversation_id": conversation_id, "filename": file.filename, "answer": answer}

@app.get("/api/conversations")
async def conversations():
    db = database()
    rows = db.execute("SELECT id, title, created_at, updated_at FROM conversations ORDER BY updated_at DESC").fetchall()
    db.close()
    return [dict(row) for row in rows]

@app.get("/api/conversations/{conversation_id}")
async def conversation(conversation_id: int):
    return {"id": conversation_id, "messages": get_messages(conversation_id)}

@app.delete("/api/conversations/{conversation_id}")
async def delete_conversation(conversation_id: int):
    db = database()
    db.execute("DELETE FROM messages WHERE conversation_id = ?", (conversation_id,))
    db.execute("DELETE FROM files WHERE conversation_id = ?", (conversation_id,))
    db.execute("DELETE FROM conversations WHERE id = ?", (conversation_id,))
    db.commit()
    db.close()
    return {"success": True}

# ============================================================
# GEMINI LIVE WEBSOCKET PROXY
# ============================================================

LIVE_WS_URL = "wss://generativelanguage.googleapis.com/ws/google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent"

async def relay_browser_to_google(browser_ws, google_ws):
    try:
        while True:
            message = await browser_ws.receive_json()
            kind = message.get("type")
            if kind == "audio" and message.get("data"):
                await google_ws.send(json.dumps({
                    "realtimeInput": {"audio": {"data": message["data"], "mimeType": "audio/pcm;rate=16000"}}
                }))
            elif kind == "video" and message.get("data"):
                await google_ws.send(json.dumps({
                    "realtimeInput": {"video": {"data": message["data"], "mimeType": "image/jpeg"}}
                }))
            elif kind == "text" and message.get("text"):
                await google_ws.send(json.dumps({"realtimeInput": {"text": str(message["text"])}}))
            elif kind == "audio_end":
                await google_ws.send(json.dumps({"realtimeInput": {"audioStreamEnd": True}}))
            elif kind == "close":
                return
    except WebSocketDisconnect:
        pass

async def relay_google_to_browser(browser_ws, google_ws):
    async for raw in google_ws:
        try:
            data = json.loads(raw)
            await browser_ws.send_json(data)
        except Exception:
            continue

@app.websocket("/ws/live")
async def live_websocket(websocket: WebSocket):
    await websocket.accept()
    if not GEMINI_API_KEY:
        await websocket.send_json({"type": "error", "message": "GEMINI_API_KEY is missing on server."})
        await websocket.close(code=1011)
        return

    google_ws = None
    try:
        google_ws = await websockets.connect(
            f"{LIVE_WS_URL}?key={GEMINI_API_KEY}",
            ping_interval=20,
            ping_timeout=20,
            max_size=20 * 1024 * 1024
        )
        setup = {
            "setup": {
                "model": f"models/{LIVE_MODEL}",
                "responseModalities": ["AUDIO"],
                "systemInstruction": {"parts": [{"text": ASRIT_SYSTEM_PROMPT}]},
                "generationConfig": {
                    "speechConfig": {"voiceConfig": {"prebuiltVoiceConfig": {"voiceName": "Puck"}}}
                }
            }
        }
        await google_ws.send(json.dumps(setup))
        first = await google_ws.recv()
        first_data = json.loads(first)
        await websocket.send_json(first_data)

        a = asyncio.create_task(relay_browser_to_google(websocket, google_ws))
        b = asyncio.create_task(relay_google_to_browser(websocket, google_ws))
        done, pending = await asyncio.wait({a, b}, return_when=asyncio.FIRST_COMPLETED)
        for task in pending:
            task.cancel()
    except Exception as error:
        try:
            await websocket.send_json({"type": "error", "message": f"Live error: {error}"})
        except Exception:
            pass
    finally:
        if google_ws:
            await google_ws.close()

# ============================================================
# HTML INTERFACE
# ============================================================

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASRIT AI</title>
<style>
* { box-sizing: border-box; }
:root {
    --background: radial-gradient(circle at 20% 0%, rgba(91, 109, 255, .22), transparent 35%),
                  radial-gradient(circle at 100% 100%, rgba(0, 220, 255, .15), transparent 35%), #070911;
    --glass: rgba(255,255,255,.065);
    --glass-strong: rgba(255,255,255,.10);
    --border: rgba(255,255,255,.12);
    --text: #f8f9ff;
    --muted: #a7adbf;
}
html, body { width: 100%; height: 100%; margin: 0; overflow: hidden; font-family: system-ui, sans-serif; color: var(--text); background: var(--background); }
.app { width: 100%; height: 100%; display: flex; position: relative; z-index: 2; }
.sidebar { width: 285px; padding: 18px; border-right: 1px solid var(--border); background: rgba(5, 7, 15, .56); backdrop-filter: blur(30px); display: flex; flex-direction: column; gap: 15px; }
.logo { display: flex; align-items: center; gap: 12px; padding: 8px; }
.logo-orb { width: 43px; height: 43px; border-radius: 15px; background: linear-gradient(135deg, #a7b0ff, #57e7ff); display: grid; place-items: center; color: #07101b; font-weight: 900; }
.logo-title { font-size: 20px; font-weight: 800; letter-spacing: .08em; }
.glass-button { border: 1px solid var(--border); background: var(--glass); color: var(--text); padding: 13px 15px; border-radius: 15px; cursor: pointer; text-align: left; }
.glass-button:hover { background: var(--glass-strong); }
.search { width: 100%; border: 1px solid var(--border); background: rgba(255,255,255,.045); color: var(--text); padding: 12px 14px; outline: none; border-radius: 13px; }
.history { flex: 1; overflow-y: auto; }
.history-item { padding: 11px 12px; border-radius: 12px; cursor: pointer; color: #dce0ec; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.history-item.active { background: rgba(139,156,255,.13); border: 1px solid rgba(139,156,255,.16); }
.main { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.topbar { height: 75px; display: flex; align-items: center; justify-content: space-between; padding: 0 28px; border-bottom: 1px solid var(--border); }
.status { display: flex; align-items: center; gap: 9px; color: var(--muted); font-size: 13px; }
.status-dot { width: 8px; height: 8px; border-radius: 50%; background: #65ffbb; box-shadow: 0 0 15px #65ffbb; }
.chat { flex: 1; overflow-y: auto; padding: 30px max(5vw, 40px); }
.welcome { min-height: 100%; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; }
.big-orb { width: 100px; height: 100px; border-radius: 34px; background: linear-gradient(135deg, #9eaaff, #4ee6ff); display: grid; place-items: center; color: #07101b; font-size: 28px; font-weight: 900; }
.message { display: flex; margin-bottom: 22px; }
.message.user { justify-content: flex-end; }
.message-bubble { max-width: min(780px, 85%); padding: 14px 17px; border-radius: 19px; line-height: 1.65; white-space: pre-wrap; word-break: break-word; }
.message.assistant .message-bubble { background: rgba(255,255,255,.045); border: 1px solid var(--border); }
.message.user .message-bubble { background: linear-gradient(135deg, rgba(128,142,255,.22), rgba(60,220,255,.10)); border: 1px solid rgba(139,156,255,.16); }
.input-area { padding: 16px max(5vw, 40px) 24px; }
.input-glass { display: flex; align-items: flex-end; gap: 10px; padding: 10px; border: 1px solid var(--border); border-radius: 22px; background: rgba(255,255,255,.065); }
#messageInput { flex: 1; resize: none; min-height: 48px; border: 0; outline: none; background: transparent; color: var(--text); padding: 13px 10px; }
.icon-button { width: 47px; height: 47px; border-radius: 15px; border: 1px solid var(--border); background: rgba(255,255,255,.055); color: white; cursor: pointer; display: grid; place-items: center; }
.camera-panel { display: none; position: fixed; inset: 0; z-index: 100; background: rgba(0,0,0,.72); backdrop-filter: blur(30px); align-items: center; justify-content: center; padding: 30px; }
.camera-panel.open { display: flex; }
.camera-box { width: min(1000px, 95vw); height: min(700px, 90vh); border: 1px solid var(--border); border-radius: 30px; background: rgba(15,18,29,.7); padding: 16px; display: flex; flex-direction: column; gap: 12px; }
#camera { width: 100%; flex: 1; object-fit: cover; border-radius: 22px; background: #020306; }
.camera-controls { display: flex; justify-content: center; gap: 10px; }
.live-meta { display:flex; align-items:center; justify-content:space-between; color:#fff; font-size:13px; font-weight:800; }
.live-state { text-align:center; color:var(--muted); font-size:13px; }
</style>
</head>
<body>
<div class="app">
    <aside class="sidebar" id="sidebar">
        <div class="logo">
            <div class="logo-orb">A</div>
            <div><div class="logo-title">ASRIT</div></div>
        </div>
        <button class="glass-button" onclick="newChat()">＋ New conversation</button>
        <input id="search" class="search" placeholder="Search chats..." oninput="searchChats()">
        <div class="history"><div id="historyList"></div></div>
    </aside>
    <main class="main">
        <header class="topbar">
            <button class="icon-button" onclick="toggleSidebar()">☰</button>
            <div class="status"><span class="status-dot"></span><span id="statusText">ASRIT online</span></div>
            <button class="glass-button" onclick="openCamera()">◉ Live</button>
        </header>
        <section id="chat" class="chat">
            <div id="welcome" class="welcome">
                <div class="big-orb">A</div>
                <h1>ASRIT</h1>
                <p>Personal multimodal AI.</p>
            </div>
            <div id="messages"></div>
        </section>
        <div class="input-area">
            <div class="input-glass">
                <input id="fileInput" type="file" hidden multiple>
                <button class="icon-button" onclick="document.getElementById('fileInput').click()">＋</button>
                <textarea id="messageInput" placeholder="Ask ASRIT anything..." rows="1" onkeydown="handleKey(event)"></textarea>
                <button class="icon-button" onclick="sendMessage()">↑</button>
            </div>
        </div>
    </main>
</div>

<div id="cameraPanel" class="camera-panel">
    <div class="camera-box">
        <video id="camera" autoplay playsinline muted></video>
        <div class="live-meta"><div>ASRIT LIVE</div><span id="liveStatusPill">● Connecting…</span></div>
        <div id="liveState" class="live-state">Opening camera and microphone…</div>
        <div class="camera-controls">
            <button class="glass-button" onclick="closeCamera()">End call</button>
        </div>
    </div>
</div>

<script>
let conversationId = 0;
let allConversations = [];
let cameraStream = null;
let liveSocket = null;
let liveAudioContext = null;
let liveProcessor = null;

window.addEventListener("load", loadHistory);

function toggleSidebar() { document.getElementById("sidebar").classList.toggle("mobile-open"); }
function newChat() {
    conversationId = 0;
    document.getElementById("messages").innerHTML = "";
    document.getElementById("welcome").style.display = "flex";
}

async function loadHistory() {
    try {
        const response = await fetch("/api/conversations");
        allConversations = await response.json();
        renderHistory(allConversations);
    } catch(e) {}
}

function renderHistory(items) {
    const list = document.getElementById("historyList");
    list.innerHTML = "";
    for (const c of items) {
        const item = document.createElement("div");
        item.className = "history-item" + (c.id === conversationId ? " active" : "");
        item.textContent = c.title;
        item.onclick = () => loadConversation(c.id);
        list.appendChild(item);
    }
}

function searchChats() {
    const query = document.getElementById("search").value.toLowerCase().trim();
    renderHistory(allConversations.filter(i => i.title.toLowerCase().includes(query)));
}

async function loadConversation(id) {
    conversationId = id;
    const response = await fetch(`/api/conversations/${id}`);
    const data = await response.json();
    document.getElementById("messages").innerHTML = "";
    document.getElementById("welcome").style.display = "none";
    for (const m of data.messages) addMessage(m.role, m.content);
}

function addMessage(role, text) {
    document.getElementById("welcome").style.display = "none";
    const wrapper = document.createElement("div");
    wrapper.className = `message ${role}`;
    const bubble = document.createElement("div");
    bubble.className = "message-bubble";
    bubble.textContent = text;
    wrapper.appendChild(bubble);
    document.getElementById("messages").appendChild(wrapper);
    document.getElementById("chat").scrollTop = document.getElementById("chat").scrollHeight;
}

async function sendMessage() {
    const input = document.getElementById("messageInput");
    const text = input.value.trim();
    if (!text) return;
    input.value = "";
    addMessage("user", text);

    const form = new FormData();
    form.append("message", text);
    form.append("conversation_id", conversationId);

    const response = await fetch("/api/chat", { method: "POST", body: form });
    const data = await response.json();
    conversationId = data.conversation_id;
    addMessage("assistant", data.answer);
    loadHistory();
}

function handleKey(e) { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(); } }

async function openCamera() {
    if (liveSocket || cameraStream) return;
    const panel = document.getElementById("cameraPanel");
    const video = document.getElementById("camera");
    panel.classList.add("open");

    try {
        cameraStream = await navigator.mediaDevices.getUserMedia({ video: true, audio: true });
        video.srcObject = cameraStream;
        await connectLiveSocket();
    } catch (error) {
        alert("Camera and Microphone permissions are required over HTTPS.\n" + error);
        closeCamera();
    }
}

async function connectLiveSocket() {
    const protocol = location.protocol === "https:" ? "wss:" : "ws:";
    liveSocket = new WebSocket(`${protocol}//${location.host}/ws/live`);
    liveSocket.onopen = () => {
        document.getElementById("liveStatusPill").textContent = "● Live";
    };
    liveSocket.onerror = (e) => {
        console.error("Live socket error:", e);
    };
}

function closeCamera() {
    if (liveSocket) { liveSocket.close(); liveSocket = null; }
    if (cameraStream) { cameraStream.getTracks().forEach(t => t.stop()); cameraStream = null; }
    document.getElementById("cameraPanel").classList.remove("open");
}
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def home():
    return HTML

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "8000")),
        reload=False
    )