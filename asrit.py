# ============================================================
#                         ASRIT AI
#              Personal AI Assistant by ASRIT
# ============================================================
#
# ONE-FILE VERSION
#
# Install:
#
# pip install -U google-genai fastapi uvicorn python-multipart pillow python-docx openpyxl pypdf websockets
#
# Run:
#
# python asrit.py
#
# Then open:
#
# http://127.0.0.1:8000
#
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

from fastapi import FastAPI, UploadFile, File, Form, WebSocket
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

# ============================================================
# 🔑 PUT YOUR GEMINI API KEY HERE
# ============================================================

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

# ============================================================
# GEMINI MODELS
# ============================================================

TEXT_MODEL = "gemini-3.5-flash-lite"

# Current Live API model example.
# If Google changes availability, change this one line.
LIVE_MODEL = "gemini-3.8-live"

# ============================================================
# ASRIT IDENTITY
# ============================================================

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

You can understand:
- normal text
- images
- documents
- audio
- video
- camera input

You should answer clearly and naturally.

When the user is using voice mode, speak naturally.

When the user provides an image or camera frame, analyze what is actually visible.

Never claim that you can see something if no image or camera frame was provided.

Do not pretend to have access to the user's phone, files, camera,
microphone, contacts, location, messages or applications unless the
application actually provides that information.

For current information, use available search grounding when enabled.

You are called ASRIT.
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
            FOREIGN KEY(conversation_id)
            REFERENCES conversations(id)
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


# ============================================================
# GEMINI CLIENT
# ============================================================

if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)
else:
    client = None


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(title="ASRIT AI")

# Generated images and uploaded media are served by the same FastAPI app.
app.mount("/media", StaticFiles(directory=str(UPLOAD_FOLDER)), name="media")


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def now():
    return datetime.now().isoformat(timespec="seconds")


def create_conversation(title="New conversation"):

    db = database()

    timestamp = now()

    cursor = db.execute(
        """
        INSERT INTO conversations
        (title, created_at, updated_at)
        VALUES (?, ?, ?)
        """,
        (title, timestamp, timestamp)
    )

    conversation_id = cursor.lastrowid

    db.commit()
    db.close()

    return conversation_id


def add_message(conversation_id, role, content):

    db = database()

    db.execute(
        """
        INSERT INTO messages
        (conversation_id, role, content, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (
            conversation_id,
            role,
            content,
            now()
        )
    )

    db.execute(
        """
        UPDATE conversations
        SET updated_at = ?
        WHERE id = ?
        """,
        (
            now(),
            conversation_id
        )
    )

    db.commit()
    db.close()


def get_messages(conversation_id):

    db = database()

    rows = db.execute(
        """
        SELECT role, content, created_at
        FROM messages
        WHERE conversation_id = ?
        ORDER BY id ASC
        """,
        (conversation_id,)
    ).fetchall()

    db.close()

    return [
        {
            "role": row["role"],
            "content": row["content"],
            "created_at": row["created_at"]
        }
        for row in rows
    ]


def conversation_title_from_text(text):

    clean = " ".join(text.strip().split())

    if not clean:
        return "New conversation"

    if len(clean) > 45:
        clean = clean[:45] + "..."

    return clean


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

    paragraphs = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            paragraphs.append(paragraph.text)

    return "\n".join(paragraphs)


def extract_xlsx(path):

    workbook = load_workbook(
        path,
        read_only=True,
        data_only=True
    )

    output = []

    for sheet in workbook.worksheets:

        output.append(
            f"\n--- SHEET: {sheet.title} ---\n"
        )

        for row in sheet.iter_rows(values_only=True):

            values = []

            for value in row:

                if value is None:
                    values.append("")
                else:
                    values.append(str(value))

            output.append(" | ".join(values))

    return "\n".join(output)


def extract_text_file(path):

    encodings = [
        "utf-8",
        "utf-16",
        "latin-1"
    ]

    for encoding in encodings:

        try:

            with open(
                path,
                "r",
                encoding=encoding,
                errors="ignore"
            ) as f:

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

    if suffix in [
        ".txt",
        ".csv",
        ".md",
        ".json",
        ".py",
        ".html",
        ".css",
        ".js"
    ]:
        return extract_text_file(path)

    return ""


# ============================================================
# GEMINI TEXT
# ============================================================

async def ask_gemini(
    user_message,
    conversation_id=None,
    extra_context=""
):

    if client is None:

        return (
            "ASRIT is not connected to Gemini yet. "
            "Open asrit.py and replace "
            '"your_api_key" with your Gemini API key.'
        )

    history = ""

    if conversation_id:

        messages = get_messages(conversation_id)

        for message in messages[-20:]:

            role = message["role"]

            if role == "user":
                history += f"\nUSER: {message['content']}\n"

            else:
                history += f"\nASRIT: {message['content']}\n"

    prompt = f"""
{ASRIT_SYSTEM_PROMPT}

Previous conversation:
{history}

Additional uploaded information:
{extra_context}

Current user message:
{user_message}

Answer as ASRIT.
"""

    try:

        response = await asyncio.to_thread(
            client.models.generate_content,
            model=TEXT_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=ASRIT_SYSTEM_PROMPT,
                temperature=0.7
            )
        )

        return response.text or "I couldn't generate a response."

    except Exception as error:

        return (
            "ASRIT encountered an error while contacting Gemini.\n\n"
            + str(error)
        )


# ============================================================
# IMAGE + GEMINI
# ============================================================

async def analyze_image(
    image_bytes,
    user_message,
    conversation_id=None
):

    if client is None:

        return (
            "Please add your Gemini API key first."
        )

    try:

        image = Image.open(
            io.BytesIO(image_bytes)
        )

        image.load()

        prompt = f"""
{ASRIT_SYSTEM_PROMPT}

Analyze the provided image.

User question:
{user_message}

Give a clear answer based only on what you can actually determine
from the image.
"""

        response = await asyncio.to_thread(
            client.models.generate_content,
            model=TEXT_MODEL,
            contents=[
                prompt,
                image
            ]
        )

        return response.text or "I couldn't analyze the image."

    except Exception as error:

        return (
            "ASRIT could not analyze this image.\n\n"
            + str(error)
        )


# ============================================================
# FILE + GEMINI
# ============================================================

async def analyze_uploaded_file(
    path,
    filename,
    mime_type,
    question
):

    if client is None:

        return (
            "Please add your Gemini API key first."
        )

    suffix = Path(filename).suffix.lower()

    # --------------------------------------------------------
    # IMAGE
    # --------------------------------------------------------

    if suffix in [
        ".jpg",
        ".jpeg",
        ".png",
        ".webp"
    ]:

        with open(path, "rb") as file:
            image_bytes = file.read()

        return await analyze_image(
            image_bytes,
            question
        )

    # --------------------------------------------------------
    # TEXT/DOCUMENT FILES
    # --------------------------------------------------------

    extracted = extract_file_text(
        path,
        mime_type,
        filename
    )

    if extracted:

        # Prevent an extremely large local prompt
        extracted = extracted[:120000]

        prompt = f"""
{ASRIT_SYSTEM_PROMPT}

The user uploaded this file:

Filename:
{filename}

Extracted content:
-------------------
{extracted}
-------------------

User question:
{question}

Analyze the file and answer the question.
"""

        try:

            response = await asyncio.to_thread(
                client.models.generate_content,
                model=TEXT_MODEL,
                contents=prompt
            )

            return response.text or "No answer generated."

        except Exception as error:

            return (
                "Error analyzing the document:\n"
                + str(error)
            )

    # --------------------------------------------------------
    # MEDIA FILES
    # --------------------------------------------------------

    if suffix in [
        ".mp3",
        ".wav",
        ".m4a",
        ".aac",
        ".mp4",
        ".mov",
        ".webm"
    ]:

        try:

            uploaded = await asyncio.to_thread(
                client.files.upload,
                file=path
            )

            prompt = f"""
{ASRIT_SYSTEM_PROMPT}

The user uploaded a media file:

{filename}

Analyze the uploaded media.

User request:
{question}
"""

            response = await asyncio.to_thread(
                client.models.generate_content,
                model=TEXT_MODEL,
                contents=[
                    prompt,
                    uploaded
                ]
            )

            return response.text or "No answer generated."

        except Exception as error:

            return (
                "ASRIT could not process this media file.\n\n"
                + str(error)
            )

    # --------------------------------------------------------
    # APK / UNKNOWN FILE
    # --------------------------------------------------------

    return (
        f"I received `{filename}`.\n\n"
        "This file type is stored by ASRIT, but this version "
        "doesn't execute or deeply inspect that file type."
    )


# ============================================================
# CHAT API
# ============================================================

@app.post("/api/chat")
async def chat(
    message: str = Form(...),
    conversation_id: int = Form(0)
):

    if conversation_id == 0:

        conversation_id = create_conversation(
            conversation_title_from_text(message)
        )

    add_message(
        conversation_id,
        "user",
        message
    )

    answer = await ask_gemini(
        message,
        conversation_id
    )

    add_message(
        conversation_id,
        "assistant",
        answer
    )

    return {
        "conversation_id": conversation_id,
        "answer": answer
    }


# ============================================================
# IMAGE API
# ============================================================

@app.post("/api/image")
async def image_api(
    image: UploadFile = File(...),
    message: str = Form("What is in this image?"),
    conversation_id: int = Form(0)
):

    if conversation_id == 0:

        conversation_id = create_conversation(
            "Image conversation"
        )

    image_bytes = await image.read()

    answer = await analyze_image(
        image_bytes,
        message,
        conversation_id
    )

    add_message(
        conversation_id,
        "user",
        "[Image] " + message
    )

    add_message(
        conversation_id,
        "assistant",
        answer
    )

    return {
        "conversation_id": conversation_id,
        "answer": answer
    }




# ============================================================
# IMAGE GENERATION
# ============================================================

IMAGE_MODEL = os.environ.get("ASRIT_IMAGE_MODEL", "gemini-3.1-flash-image")


async def generate_image(prompt, conversation_id=0):
    if client is None:
        return {"error": "Please configure GEMINI_API_KEY first."}

    prompt = (prompt or "").strip()
    if not prompt:
        return {"error": "Please describe the image you want."}

    try:
        interaction = await asyncio.to_thread(
            client.interactions.create,
            model=IMAGE_MODEL,
            input=prompt,
            response_format={
                "type": "image",
                "aspect_ratio": "1:1",
                "image_size": "1K"
            }
        )

        image_data = getattr(getattr(interaction, "output_image", None), "data", None)

        if not image_data:
            for step in getattr(interaction, "steps", []) or []:
                if getattr(step, "type", "") != "model_output":
                    continue
                for block in getattr(step, "content", []) or []:
                    if getattr(block, "type", "") == "image" and getattr(block, "data", None):
                        image_data = block.data
                        break
                if image_data:
                    break

        if not image_data:
            return {"error": "The image model returned no image."}

        filename = datetime.now().strftime("asrit_%Y%m%d_%H%M%S_%f.png")
        path = UPLOAD_FOLDER / filename
        with open(path, "wb") as output:
            output.write(base64.b64decode(image_data))

        if conversation_id:
            add_message(conversation_id, "assistant", f"[Generated image] {prompt}")

        return {"url": f"/media/{filename}", "prompt": prompt}

    except Exception as error:
        return {"error": f"Image generation error: {error}"}


@app.post("/api/generate-image")
async def generate_image_api(
    prompt: str = Form(...),
    conversation_id: int = Form(0)
):
    if conversation_id == 0:
        conversation_id = create_conversation("Image generation")

    result = await generate_image(prompt, conversation_id)
    result["conversation_id"] = conversation_id
    return result


# ============================================================
# GEMINI LIVE API WEBSOCKET PROXY
# ============================================================

LIVE_MODEL = "gemini-3.8-live"
LIVE_WS_URL = (
    "wss://generativelanguage.googleapis.com/ws/"
    "google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent"
)

async def relay_browser_to_google(browser_ws, google_ws):
    while True:
        message = await browser_ws.receive_json()
        kind = message.get("type")
        if kind == "audio":
            data = message.get("data")
            if data:
                await google_ws.send(json.dumps({
                    "realtimeInput": {"audio": {"data": data, "mimeType": "audio/pcm;rate=16000"}}
                }))
        elif kind == "video":
            data = message.get("data")
            if data:
                await google_ws.send(json.dumps({
                    "realtimeInput": {"video": {"data": data, "mimeType": "image/jpeg"}}
                }))
        elif kind == "text":
            text_value = str(message.get("text", "")).strip()
            if text_value:
                # Realtime text input is documented by the Live API.
                await google_ws.send(json.dumps({"realtimeInput": {"text": text_value}}))
        elif kind == "audio_end":
            await google_ws.send(json.dumps({"realtimeInput": {"audioStreamEnd": True}}))
        elif kind == "close":
            return

async def relay_google_to_browser(browser_ws, google_ws):
    async for raw in google_ws:
        try:
            data = json.loads(raw)
        except Exception:
            continue
        await browser_ws.send_json(data)

@app.websocket("/ws/live")
async def live_websocket(websocket: WebSocket):
    await websocket.accept()
    if not GEMINI_API_KEY:
        await websocket.send_json({"type": "error", "message": "GEMINI_API_KEY is not configured on the server."})
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
                "inputAudioTranscription": {},
                "outputAudioTranscription": {},
                "generationConfig": {
                    "speechConfig": {
                        "voiceConfig": {
                            "prebuiltVoiceConfig": {"voiceName": "Puck"}
                        }
                    }
                }
            }
        }
        await google_ws.send(json.dumps(setup))
        first = await google_ws.recv()
        try:
            first_data = json.loads(first)
        except Exception:
            first_data = {"type": "error", "message": "Invalid Live API setup response."}
        await websocket.send_json(first_data)
        if "setupComplete" not in first_data:
            await websocket.close(code=1011)
            return
        a = asyncio.create_task(relay_browser_to_google(websocket, google_ws))
        b = asyncio.create_task(relay_google_to_browser(websocket, google_ws))
        done, pending = await asyncio.wait({a, b}, return_when=asyncio.FIRST_COMPLETED)
        for task in pending:
            task.cancel()
        for task in done:
            try:
                task.result()
            except Exception:
                pass
    except Exception as error:
        try:
            await websocket.send_json({"type": "error", "message": f"Live connection error: {error}"})
        except Exception:
            pass
    finally:
        if google_ws is not None:
            try:
                await google_ws.close()
            except Exception:
                pass
        try:
            await websocket.close()
        except Exception:
            pass

# ============================================================
# FILE API
# ============================================================

@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    message: str = Form("Analyze this file."),
    conversation_id: int = Form(0)
):

    if conversation_id == 0:

        conversation_id = create_conversation(
            file.filename
        )

    safe_name = (
        datetime.now().strftime("%Y%m%d_%H%M%S_")
        + Path(file.filename).name
    )

    save_path = UPLOAD_FOLDER / safe_name

    content = await file.read()

    with open(
        save_path,
        "wb"
    ) as output:

        output.write(content)

    mime_type = (
        file.content_type
        or mimetypes.guess_type(file.filename)[0]
        or "application/octet-stream"
    )

    db = database()

    db.execute(
        """
        INSERT INTO files
        (
            conversation_id,
            filename,
            mime_type,
            path,
            created_at
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            conversation_id,
            file.filename,
            mime_type,
            str(save_path),
            now()
        )
    )

    db.commit()
    db.close()

    answer = await analyze_uploaded_file(
        str(save_path),
        file.filename,
        mime_type,
        message
    )

    add_message(
        conversation_id,
        "user",
        f"[Uploaded file: {file.filename}] {message}"
    )

    add_message(
        conversation_id,
        "assistant",
        answer
    )

    return {
        "conversation_id": conversation_id,
        "filename": file.filename,
        "answer": answer
    }


# ============================================================
# HISTORY API
# ============================================================

@app.get("/api/conversations")
async def conversations():

    db = database()

    rows = db.execute(
        """
        SELECT id, title, created_at, updated_at
        FROM conversations
        ORDER BY updated_at DESC
        """
    ).fetchall()

    db.close()

    return [
        dict(row)
        for row in rows
    ]


@app.get("/api/conversations/{conversation_id}")
async def conversation(
    conversation_id: int
):

    return {
        "id": conversation_id,
        "messages": get_messages(
            conversation_id
        )
    }


@app.delete("/api/conversations/{conversation_id}")
async def delete_conversation(
    conversation_id: int
):

    db = database()

    db.execute(
        """
        DELETE FROM messages
        WHERE conversation_id = ?
        """,
        (conversation_id,)
    )

    db.execute(
        """
        DELETE FROM files
        WHERE conversation_id = ?
        """,
        (conversation_id,)
    )

    db.execute(
        """
        DELETE FROM conversations
        WHERE id = ?
        """,
        (conversation_id,)
    )

    db.commit()
    db.close()

    return {
        "success": True
    }


# ============================================================
# MAIN HTML
# ============================================================

HTML = r"""
<!DOCTYPE html>

<html lang="en">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>ASRIT AI</title>

<style>

/* =========================================================
   GLOBAL
========================================================= */

* {
    box-sizing: border-box;
}

:root {

    --background:
        radial-gradient(
            circle at 20% 0%,
            rgba(91, 109, 255, .22),
            transparent 35%
        ),
        radial-gradient(
            circle at 100% 100%,
            rgba(0, 220, 255, .15),
            transparent 35%
        ),
        #070911;

    --glass: rgba(255,255,255,.065);

    --glass-strong: rgba(255,255,255,.10);

    --border: rgba(255,255,255,.12);

    --text: #f8f9ff;

    --muted: #a7adbf;

    --accent: #8b9cff;

    --accent2: #52e5ff;

    --danger: #ff657a;
}

html,
body {

    width: 100%;
    height: 100%;

    margin: 0;

    overflow: hidden;

    font-family:
        Inter,
        ui-sans-serif,
        system-ui,
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;

    color: var(--text);

    background: var(--background);
}

/* =========================================================
   BACKGROUND
========================================================= */

body::before {

    content: "";

    position: fixed;

    width: 450px;
    height: 450px;

    left: -150px;
    top: -150px;

    background:
        radial-gradient(
            circle,
            rgba(119, 102, 255, .32),
            transparent 70%
        );

    filter: blur(40px);

    animation:
        float1 12s infinite alternate ease-in-out;

    pointer-events: none;
}

body::after {

    content: "";

    position: fixed;

    width: 500px;
    height: 500px;

    right: -200px;
    bottom: -200px;

    background:
        radial-gradient(
            circle,
            rgba(0, 210, 255, .2),
            transparent 70%
        );

    filter: blur(50px);

    animation:
        float2 14s infinite alternate ease-in-out;

    pointer-events: none;
}

@keyframes float1 {

    from {
        transform: translate(
            0,
            0
        );
    }

    to {
        transform: translate(
            180px,
            130px
        );
    }
}

@keyframes float2 {

    from {
        transform: translate(
            0,
            0
        );
    }

    to {
        transform: translate(
            -150px,
            -100px
        );
    }
}

/* =========================================================
   APP
========================================================= */

.app {

    width: 100%;
    height: 100%;

    display: flex;

    position: relative;

    z-index: 2;
}

/* =========================================================
   SIDEBAR
========================================================= */

.sidebar {

    width: 285px;

    padding: 18px;

    border-right:
        1px solid var(--border);

    background:
        rgba(5, 7, 15, .56);

    backdrop-filter:
        blur(30px);

    -webkit-backdrop-filter:
        blur(30px);

    display: flex;

    flex-direction: column;

    gap: 15px;
}

/* =========================================================
   LOGO
========================================================= */

.logo {

    display: flex;

    align-items: center;

    gap: 12px;

    padding: 8px;
}

.logo-orb {

    width: 43px;
    height: 43px;

    border-radius: 15px;

    background:
        linear-gradient(
            135deg,
            #a7b0ff,
            #57e7ff
        );

    box-shadow:
        0 0 30px rgba(93, 216, 255, .32);

    display: grid;

    place-items: center;

    color: #07101b;

    font-weight: 900;

    animation:
        logoPulse 4s infinite ease-in-out;
}

@keyframes logoPulse {

    50% {
        transform:
            scale(1.04)
            rotate(2deg);
    }
}

.logo-title {

    font-size: 20px;

    font-weight: 800;

    letter-spacing: .08em;
}

.logo-sub {

    color: var(--muted);

    font-size: 11px;

    margin-top: 2px;
}

/* =========================================================
   BUTTON
========================================================= */

.glass-button {

    border:
        1px solid var(--border);

    background:
        var(--glass);

    color: var(--text);

    padding: 13px 15px;

    border-radius: 15px;

    cursor: pointer;

    transition:
        transform .2s ease,
        background .2s ease,
        border .2s ease;

    font-size: 14px;

    text-align: left;
}

.glass-button:hover {

    background:
        var(--glass-strong);

    transform:
        translateY(-2px);

    border-color:
        rgba(255,255,255,.2);
}

.new-chat {

    background:
        linear-gradient(
            135deg,
            rgba(130,145,255,.22),
            rgba(75,220,255,.12)
        );
}

/* =========================================================
   SEARCH
========================================================= */

.search {

    width: 100%;

    border:
        1px solid var(--border);

    background:
        rgba(255,255,255,.045);

    color: var(--text);

    padding: 12px 14px;

    outline: none;

    border-radius: 13px;
}

.search::placeholder {

    color: #777f94;
}

/* =========================================================
   HISTORY
========================================================= */

.history {

    flex: 1;

    overflow-y: auto;

    padding-right: 3px;
}

.history-title {

    color: #737b91;

    font-size: 11px;

    text-transform: uppercase;

    letter-spacing: .12em;

    padding: 10px 8px;
}

.history-item {

    padding: 11px 12px;

    border-radius: 12px;

    cursor: pointer;

    color: #dce0ec;

    margin-bottom: 3px;

    white-space: nowrap;

    overflow: hidden;

    text-overflow: ellipsis;

    transition:
        background .2s ease,
        transform .2s ease;
}

.history-item:hover {

    background:
        rgba(255,255,255,.06);

    transform:
        translateX(3px);
}

.history-item.active {

    background:
        rgba(139,156,255,.13);

    border:
        1px solid rgba(139,156,255,.16);
}

/* =========================================================
   MAIN
========================================================= */

.main {

    flex: 1;

    min-width: 0;

    display: flex;

    flex-direction: column;
}

/* =========================================================
   TOPBAR
========================================================= */

.topbar {

    height: 75px;

    display: flex;

    align-items: center;

    justify-content: space-between;

    padding:
        0 28px;

    border-bottom:
        1px solid var(--border);

    background:
        rgba(8, 10, 18, .30);

    backdrop-filter:
        blur(25px);

    -webkit-backdrop-filter:
        blur(25px);
}

.status {

    display: flex;

    align-items: center;

    gap: 9px;

    color: var(--muted);

    font-size: 13px;
}

.status-dot {

    width: 8px;
    height: 8px;

    border-radius: 50%;

    background: #65ffbb;

    box-shadow:
        0 0 15px #65ffbb;
}

/* =========================================================
   CHAT
========================================================= */

.chat {

    flex: 1;

    overflow-y: auto;

    padding:
        30px
        max(5vw, 40px);

    scroll-behavior: smooth;
}

.welcome {

    min-height: 100%;

    display: flex;

    flex-direction: column;

    justify-content: center;

    align-items: center;

    text-align: center;

    animation:
        appear .7s ease;
}

@keyframes appear {

    from {

        opacity: 0;

        transform:
            translateY(18px)
            scale(.98);
    }

    to {

        opacity: 1;

        transform:
            none;
    }
}

.big-orb {

    width: 100px;
    height: 100px;

    border-radius: 34px;

    background:
        linear-gradient(
            135deg,
            #9eaaff,
            #4ee6ff
        );

    display: grid;

    place-items: center;

    color: #07101b;

    font-size: 28px;

    font-weight: 900;

    box-shadow:
        0 0 80px rgba(80,210,255,.25);

    animation:
        heroFloat 4s infinite ease-in-out;
}

@keyframes heroFloat {

    0%,
    100% {
        transform:
            translateY(0)
            rotate(0deg);
    }

    50% {
        transform:
            translateY(-10px)
            rotate(2deg);
    }
}

.welcome h1 {

    margin:
        25px 0 8px;

    font-size:
        clamp(36px, 6vw, 64px);

    letter-spacing:
        -.045em;
}

.welcome p {

    color: var(--muted);

    max-width: 550px;

    line-height: 1.6;
}

/* =========================================================
   MESSAGE
========================================================= */

.message {

    display: flex;

    margin-bottom: 22px;

    animation:
        messageIn .3s ease;
}

@keyframes messageIn {

    from {

        opacity: 0;

        transform:
            translateY(8px);
    }

    to {

        opacity: 1;

        transform:
            none;
    }
}

.message.user {

    justify-content: flex-end;
}

.message-bubble {

    max-width: min(780px, 85%);

    padding: 14px 17px;

    border-radius: 19px;

    line-height: 1.65;

    white-space: pre-wrap;

    word-break: break-word;
}

.message.assistant
.message-bubble {

    background:
        rgba(255,255,255,.045);

    border:
        1px solid var(--border);

    backdrop-filter:
        blur(15px);
}

.message.user
.message-bubble {

    background:
        linear-gradient(
            135deg,
            rgba(128,142,255,.22),
            rgba(60,220,255,.10)
        );

    border:
        1px solid rgba(139,156,255,.16);
}

/* =========================================================
   INPUT AREA
========================================================= */

.input-area {

    padding:
        16px
        max(5vw, 40px)
        24px;
}

.input-glass {

    display: flex;

    align-items: flex-end;

    gap: 10px;

    padding: 10px;

    border:
        1px solid var(--border);

    border-radius: 22px;

    background:
        rgba(255,255,255,.065);

    backdrop-filter:
        blur(30px);

    -webkit-backdrop-filter:
        blur(30px);

    box-shadow:
        0 20px 80px rgba(0,0,0,.22);
}

#messageInput {

    flex: 1;

    resize: none;

    min-height: 48px;

    max-height: 160px;

    border: 0;

    outline: none;

    background: transparent;

    color: var(--text);

    font:
        inherit;

    padding:
        13px 10px;
}

#messageInput::placeholder {

    color:
        #777f94;
}

.icon-button {

    width: 47px;
    height: 47px;

    border-radius: 15px;

    border:
        1px solid var(--border);

    background:
        rgba(255,255,255,.055);

    color: white;

    cursor: pointer;

    transition:
        .2s ease;

    display: grid;

    place-items: center;

    font-size: 18px;
}

.icon-button:hover {

    transform:
        translateY(-2px)
        scale(1.03);

    background:
        rgba(255,255,255,.11);
}

.send-button {

    background:
        linear-gradient(
            135deg,
            #9da8ff,
            #54e4ff
        );

    color: #07101b;

    border: 0;

    font-weight: 800;
}

.recording {

    background:
        rgba(255,90,110,.18);

    border-color:
        rgba(255,90,110,.35);

    animation:
        recordPulse 1s infinite;
}

@keyframes recordPulse {

    50% {
        box-shadow:
            0 0 0 8px rgba(255,90,110,.05);
    }
}

/* =========================================================
   CAMERA
========================================================= */

.camera-panel {

    display: none;

    position: fixed;

    inset: 0;

    z-index: 100;

    background:
        rgba(0,0,0,.72);

    backdrop-filter:
        blur(30px);

    align-items: center;

    justify-content: center;

    padding: 30px;
}

.camera-panel.open {

    display: flex;
}

.camera-box {

    width: min(1000px, 95vw);

    height: min(700px, 90vh);

    border:
        1px solid var(--border);

    border-radius: 30px;

    background:
        rgba(15,18,29,.7);

    backdrop-filter:
        blur(35px);

    padding: 16px;

    display: flex;

    flex-direction: column;

    gap: 12px;

    box-shadow:
        0 40px 120px rgba(0,0,0,.5);
}

#camera {

    width: 100%;

    flex: 1;

    min-height: 0;

    object-fit: cover;

    border-radius: 22px;

    background: #020306;
}

.camera-controls {

    display: flex;

    justify-content: center;

    gap: 10px;
}

/* =========================================================
   MOBILE
========================================================= */

@media (max-width: 800px) {

    .sidebar {

        position: fixed;

        left: -300px;

        top: 0;

        bottom: 0;

        z-index: 200;

        transition:
            left .3s ease;
    }

    .sidebar.mobile-open {

        left: 0;
    }

    .topbar {

        padding:
            0 16px;
    }

    .chat {

        padding:
            20px 15px;
    }

    .input-area {

        padding:
            10px 15px 15px;
    }

    .message-bubble {

        max-width: 92%;
    }
}



/* =========================================================
   ASRIT LIVE + PASTE EXTRAS
========================================================= */
.live-meta {
    display:flex;
    align-items:center;
    justify-content:space-between;
    color:#fff;
    font-size:13px;
    font-weight:800;
    letter-spacing:.08em;
    padding:2px 4px;
}
.live-meta span { color:#65ffbb; font-size:12px; letter-spacing:0; }
.live-state { text-align:center; color:var(--muted); font-size:13px; min-height:18px; }
.camera-controls .off { opacity:.55; }
.paste-hint {
    position:fixed; left:50%; bottom:110px;
    transform:translateX(-50%) translateY(12px);
    padding:11px 16px; border:1px solid rgba(255,255,255,.14);
    border-radius:14px; background:rgba(10,12,20,.72);
    backdrop-filter:blur(24px); color:#fff; opacity:0;
    pointer-events:none; transition:.25s ease; z-index:500;
}
.paste-hint.show { opacity:1; transform:translateX(-50%) translateY(0); }


/* =========================================================
   ASRIT 2.0 — iOS LIQUID GLASS OVERLAY
   Visual-only layer: existing chat/backend behavior is preserved.
========================================================= */
:root {
    --sky-0:#eef8ff;
    --sky-1:#dff2ff;
    --sky-2:#c7e9ff;
    --ink:#10243a;
    --muted2:#60758b;
    --glass-light:rgba(255,255,255,.54);
    --glass-white:rgba(255,255,255,.72);
    --glass-line:rgba(255,255,255,.78);
    --shadow-soft:0 18px 55px rgba(47,106,148,.14);
}
html, body {
    background:
      radial-gradient(circle at 15% 8%, rgba(255,255,255,.96), transparent 30%),
      radial-gradient(circle at 88% 12%, rgba(161,220,255,.62), transparent 32%),
      radial-gradient(circle at 72% 92%, rgba(188,228,255,.65), transparent 34%),
      linear-gradient(145deg,var(--sky-0),var(--sky-1) 48%,var(--sky-2));
    color:var(--ink);
}
body::before, body::after { opacity:.45; filter:blur(58px); }
.sidebar,
.topbar,
.input-glass,
.message.assistant .message-bubble,
.camera-box,
.glass-button,
.icon-button,
.search {
    background:linear-gradient(145deg,rgba(255,255,255,.70),rgba(255,255,255,.34));
    border-color:var(--glass-line);
    box-shadow:var(--shadow-soft), inset 0 1px 0 rgba(255,255,255,.9);
    -webkit-backdrop-filter:blur(28px) saturate(145%);
    backdrop-filter:blur(28px) saturate(145%);
}
.sidebar { color:var(--ink); }
.logo-title,.welcome h1 { color:#10243a; }
.logo-sub,.history-title,.status,.welcome p,.live-state { color:var(--muted2); }
.history-item { color:#17324b; }
.history-item:hover { background:rgba(255,255,255,.45); }
.history-item.active { background:rgba(255,255,255,.62); border-color:rgba(120,190,235,.45); }
.search { color:var(--ink); }
.search::placeholder,#messageInput::placeholder { color:#71879b; }
#messageInput { color:var(--ink); }
.message.user .message-bubble { 
    background:linear-gradient(145deg,rgba(255,255,255,.82),rgba(191,228,250,.58));
    border:1px solid rgba(255,255,255,.88);
    color:#17324b;
    box-shadow:0 12px 30px rgba(56,125,167,.12), inset 0 1px 0 rgba(255,255,255,.9);
}
.big-orb,.logo-orb { box-shadow:0 18px 55px rgba(68,164,222,.25), inset 0 1px 0 rgba(255,255,255,.8); }
.send-button { box-shadow:0 10px 28px rgba(50,150,215,.24); }
.icon-button:hover,.glass-button:hover { transform:translateY(-1px) scale(1.015); }

/* Premium navigation */
.side-nav { display:grid; gap:7px; }
.nav-item { display:flex; align-items:center; gap:12px; width:100%; padding:11px 13px; border:1px solid transparent; border-radius:15px; background:transparent; color:#24445e; cursor:pointer; text-align:left; font:inherit; transition:.18s ease; }
.nav-item:hover,.nav-item.active { background:rgba(255,255,255,.52); border-color:rgba(255,255,255,.75); box-shadow:0 8px 24px rgba(61,127,164,.10); }
.nav-icon { width:27px; text-align:center; font-size:17px; }
.sidebar-divider { height:1px; background:rgba(91,143,174,.18); margin:3px 0; }

/* Home / mode cards */
.quick-actions { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; width:min(760px,94%); margin-top:26px; }
.quick-card { border:1px solid rgba(255,255,255,.82); border-radius:22px; padding:16px; background:rgba(255,255,255,.42); color:#17324b; text-align:left; cursor:pointer; box-shadow:var(--shadow-soft),inset 0 1px 0 white; backdrop-filter:blur(22px); transition:.2s ease; }
.quick-card:hover { transform:translateY(-3px); background:rgba(255,255,255,.62); }
.quick-card strong { display:block; margin-bottom:4px; }
.quick-card small { color:#688096; }

/* Image studio */
.image-studio { display:none; min-height:100%; padding:28px; align-items:center; justify-content:center; }
.image-studio.open { display:flex; }
.image-studio-card { width:min(900px,100%); border:1px solid rgba(255,255,255,.84); border-radius:30px; padding:28px; background:rgba(255,255,255,.48); backdrop-filter:blur(32px) saturate(150%); box-shadow:0 25px 80px rgba(54,119,158,.16), inset 0 1px 0 white; }
.image-prompt { width:100%; min-height:130px; resize:vertical; border:1px solid rgba(140,190,220,.35); border-radius:20px; padding:16px; background:rgba(255,255,255,.58); color:#17324b; outline:none; font:inherit; }
.image-result { margin-top:18px; display:grid; gap:12px; }
.image-result img { width:100%; max-height:65vh; object-fit:contain; border-radius:22px; box-shadow:0 18px 60px rgba(39,96,135,.18); }

/* Bottom dock */
.mobile-dock { display:none; }

/* Mobile polish */
.mobile-back { display:none; }
@media (max-width:800px) {
    .sidebar { width:min(88vw,340px); left:-360px; padding:16px; }
    .topbar { height:64px; padding:0 12px; }
    .chat { padding:16px 12px 12px; }
    .input-area { padding:8px 10px calc(12px + env(safe-area-inset-bottom)); }
    .input-glass { border-radius:25px; padding:7px; gap:6px; }
    .icon-button { width:44px; height:44px; border-radius:14px; }
    .quick-actions { grid-template-columns:1fr; }
    .mobile-dock { position:fixed; display:flex; left:10px; right:10px; bottom:10px; z-index:180; padding:7px; gap:4px; justify-content:space-around; border:1px solid rgba(255,255,255,.78); border-radius:22px; background:rgba(255,255,255,.58); backdrop-filter:blur(30px) saturate(150%); box-shadow:0 18px 55px rgba(40,96,133,.20),inset 0 1px 0 white; }
    .dock-btn { border:0; background:transparent; color:#24445e; min-width:54px; padding:7px 5px; border-radius:16px; font:inherit; font-size:11px; cursor:pointer; }
    .dock-btn span { display:block; font-size:18px; margin-bottom:2px; }
    .input-area { padding-bottom:78px; }
    .mobile-back { display:grid; }
    .topbar .status { flex:1; justify-content:center; }
    .topbar .glass-button { padding:10px 12px; }
    .image-studio { padding:14px; padding-bottom:90px; }
    .image-studio-card { padding:18px; border-radius:25px; }
}

@media (prefers-reduced-motion:reduce) {
    *,*::before,*::after { animation-duration:.01ms!important; transition-duration:.01ms!important; scroll-behavior:auto!important; }
}

</style>

</head>

<body>

<div class="app">

    <!-- =====================================================
         SIDEBAR
    ====================================================== -->

    <aside class="sidebar" id="sidebar">

        <div class="logo">

            <div class="logo-orb">
                A
            </div>

            <div>

                <div class="logo-title">
                    ASRIT
                </div>

                <div class="logo-sub">
                    PERSONAL AI
                </div>

            </div>

        </div>

        <button
            class="glass-button new-chat"
            onclick="newChat()"
        >
            ＋ New conversation
        </button>

        <input
            id="search"
            class="search"
            placeholder="Search chats..."
            oninput="searchChats()"
        >

        <div class="side-nav">
            <button class="nav-item active" data-mode="chat" onclick="showMode('chat')"><span class="nav-icon">◌</span>Chat</button>
            <button class="nav-item" data-mode="images" onclick="showMode('images')"><span class="nav-icon">✦</span>Images</button>
            <button class="nav-item" onclick="openFeature('library')"><span class="nav-icon">▤</span>Library</button>
            <button class="nav-item" onclick="openFeature('scheduled')"><span class="nav-icon">◷</span>Scheduled</button>
            <button class="nav-item" onclick="openFeature('plugins')"><span class="nav-icon">⌘</span>Plugins</button>
            <button class="nav-item" onclick="openFeature('projects')"><span class="nav-icon">□</span>Projects</button>
            <button class="nav-item" onclick="openFeature('codex')"><span class="nav-icon">{ }</span>Codex</button>
        </div>

        <div class="history">

            <div class="history-title">
                Conversations
            </div>

            <div id="historyList"></div>

        </div>

    </aside>


    <!-- =====================================================
         MAIN
    ====================================================== -->

    <main class="main">

        <header class="topbar">

            <button class="icon-button mobile-back" onclick="goHome()" title="Back to home">‹</button>
            <button
                class="icon-button"
                onclick="toggleSidebar()"
                title="Menu"
            >
                ☰
            </button>

            <div class="status">

                <span class="status-dot"></span>

                <span id="statusText">
                    ASRIT online
                </span>

            </div>

            <button
                class="glass-button"
                onclick="openCamera()"
            >
                ◉ Live
            </button>

        </header>


        <section
            id="chat"
            class="chat"
        >

            <div
                id="welcome"
                class="welcome"
            >

                <div class="big-orb">
                    A
                </div>

                <h1>
                    ASRIT
                </h1>

                <p>
                    Your personal multimodal AI.
                    Ask anything, paste screenshots with Ctrl+V,
                    create images, upload files, or talk to ASRIT Live.
                </p>

                <div class="quick-actions">
                    <button class="quick-card" onclick="showMode('images')"><strong>✦ Create images</strong><small>Generate posters, art and concepts.</small></button>
                    <button class="quick-card" onclick="openCamera()"><strong>◉ ASRIT Live</strong><small>Talk naturally with camera + voice.</small></button>
                    <button class="quick-card" onclick="document.getElementById('fileInput').click()"><strong>＋ Add anything</strong><small>Files, photos and screenshots.</small></button>
                    <button class="quick-card" onclick="openFeature('codex')"><strong>{ } Codex</strong><small>Open a coding workspace.</small></button>
                </div>

            </div>

            <div id="messages"></div>

        </section>


        <section id="imageStudio" class="image-studio">
            <div class="image-studio-card">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:8px">
                    <div><div style="font-size:13px;color:#6a8196">ASRIT CREATIVE</div><h2 style="margin:3px 0 0">Create an image</h2></div>
                    <button class="icon-button" onclick="showMode('chat')">×</button>
                </div>
                <p style="color:#668096">Describe the image, poster, illustration, concept art, or edit you want.</p>
                <textarea id="imagePrompt" class="image-prompt" placeholder="Create a cinematic futuristic city at sunrise..."></textarea>
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap">
                    <button class="glass-button" onclick="useImageExample()">Example</button>
                    <button class="glass-button send-button" onclick="generateImage()">✦ Generate image</button>
                </div>
                <div id="imageResult" class="image-result"></div>
            </div>
        </section>

        <!-- =================================================
             INPUT
        ================================================== -->

        <div class="input-area">

            <div class="input-glass">

                <input id="fileInput" type="file" hidden multiple>
                <input id="cameraFileInput" type="file" hidden accept="image/*" capture="environment">

                <button
                    class="icon-button"
                    onclick="document.getElementById('fileInput').click()"
                    title="Add files, photos or screenshots"
                >
                    ＋
                </button>

                <textarea
                    id="messageInput"
                    placeholder="Ask ASRIT anything..."
                    rows="1"
                    onkeydown="handleKey(event)"
                ></textarea>

                <button
                    id="micButton"
                    class="icon-button"
                    onclick="toggleVoice()"
                    title="Voice"
                >
                    🎙
                </button>

                <button
                    class="icon-button send-button"
                    onclick="sendMessage()"
                >
                    ↑
                </button>

            </div>

        </div>

    </main>

</div>

<div id="pasteHint" class="paste-hint">Image pasted — ready to analyze</div>

<nav class="mobile-dock" aria-label="ASRIT quick navigation">
    <button class="dock-btn" onclick="goHome()"><span>⌂</span>Home</button>
    <button class="dock-btn" onclick="showMode('images')"><span>✦</span>Images</button>
    <button class="dock-btn" onclick="document.getElementById('fileInput').click()"><span>＋</span>Add</button>
    <button class="dock-btn" onclick="openCamera()"><span>◉</span>Live</button>
    <button class="dock-btn" onclick="toggleSidebar()"><span>☰</span>Menu</button>
</nav>


<!-- =========================================================
     CAMERA
========================================================= -->

<div
    id="cameraPanel"
    class="camera-panel"
>

    <div class="camera-box">

        <video
            id="camera"
            autoplay
            playsinline
            muted
        ></video>

        <div class="live-meta">
            <div>ASRIT LIVE</div>
            <span id="liveStatusPill">● Connecting…</span>
        </div>

        <div id="liveState" class="live-state">Opening camera and microphone…</div>

        <div class="camera-controls">
            <button id="liveMicButton" class="glass-button" onclick="toggleLiveMic()" title="Mute microphone">🎙</button>
            <button id="liveCameraButton" class="glass-button" onclick="toggleLiveCamera()" title="Pause camera">📷</button>
            <button class="glass-button" onclick="analyzeCamera()">✦ Ask ASRIT</button>
            <button class="glass-button" onclick="closeCamera()">End call</button>
        </div>

    </div>

</div>


<script>

/* =========================================================
   STATE
========================================================= */

let conversationId = 0;

let allConversations = [];

let recognition = null;

let isRecording = false;

let cameraStream = null;

let liveSocket = null;
let liveAudioContext = null;
let liveSource = null;
let liveProcessor = null;
let liveVideoTimer = null;
let liveMicEnabled = true;
let liveCameraEnabled = true;
let liveNextAudioTime = 0;
let liveClosing = false;
let currentMode = "chat";
let pastedImageFile = null;


/* =========================================================
   INITIALIZE
========================================================= */

window.addEventListener(
    "load",
    () => {

        loadHistory();

        const input =
            document.getElementById(
                "messageInput"
            );

        input.focus();

    }
);


/* =========================================================
   SIDEBAR
========================================================= */

function toggleSidebar() {
    document.getElementById("sidebar").classList.toggle("mobile-open");
}

function closeSidebarMobile() {
    document.getElementById("sidebar")?.classList.remove("mobile-open");
}


/* =========================================================
   NEW CHAT
========================================================= */

function newChat() {

    showMode("chat");
    closeSidebarMobile();
    conversationId = 0;

    document
        .getElementById("messages")
        .innerHTML = "";

    document
        .getElementById("welcome")
        .style.display = "flex";

    document
        .getElementById("messageInput")
        .focus();

}


/* =========================================================
   LOAD HISTORY
========================================================= */

async function loadHistory() {

    try {

        const response =
            await fetch(
                "/api/conversations"
            );

        allConversations =
            await response.json();

        renderHistory(
            allConversations
        );

    } catch(error) {

        console.error(error);

    }

}


/* =========================================================
   RENDER HISTORY
========================================================= */

function renderHistory(items) {

    const list =
        document.getElementById(
            "historyList"
        );

    list.innerHTML = "";

    for (
        const conversation
        of items
    ) {

        const item =
            document.createElement(
                "div"
            );

        item.className =
            "history-item";

        if (
            conversation.id ===
            conversationId
        ) {

            item.classList.add(
                "active"
            );

        }

        item.textContent =
            conversation.title;

        item.onclick =
            () => loadConversation(
                conversation.id
            );

        list.appendChild(item);

    }

}


/* =========================================================
   SEARCH
========================================================= */

function searchChats() {

    const query =
        document
            .getElementById("search")
            .value
            .toLowerCase()
            .trim();

    const filtered =
        allConversations.filter(
            item =>
                item.title
                    .toLowerCase()
                    .includes(query)
        );

    renderHistory(filtered);

}


/* =========================================================
   LOAD CONVERSATION
========================================================= */

async function loadConversation(id) {

    showMode("chat");
    closeSidebarMobile();
    conversationId = id;

    try {

        const response =
            await fetch(
                `/api/conversations/${id}`
            );

        const data =
            await response.json();

        document
            .getElementById("messages")
            .innerHTML = "";

        document
            .getElementById("welcome")
            .style.display = "none";

        for (
            const message
            of data.messages
        ) {

            addMessage(
                message.role ===
                    "assistant"
                    ? "assistant"
                    : "user",
                message.content
            );

        }

        renderHistory(
            allConversations
        );

    } catch(error) {

        console.error(error);

    }

}


/* =========================================================
   ADD MESSAGE
========================================================= */

function addMessage(
    role,
    text
) {

    document
        .getElementById("welcome")
        .style.display = "none";

    const wrapper =
        document.createElement(
            "div"
        );

    wrapper.className =
        `message ${role}`;

    const bubble =
        document.createElement(
            "div"
        );

    bubble.className =
        "message-bubble";

    bubble.textContent =
        text;

    wrapper.appendChild(
        bubble
    );

    document
        .getElementById("messages")
        .appendChild(
            wrapper
        );

    const chat =
        document.getElementById(
            "chat"
        );

    chat.scrollTop =
        chat.scrollHeight;

}



/* =========================================================
   IOS-LIKE NAVIGATION / MODES
========================================================= */
function goHome() {
    showMode("chat");
    conversationId = 0;
    document.getElementById("messages").innerHTML = "";
    document.getElementById("welcome").style.display = "flex";
    closeSidebarMobile();
}

function showMode(mode) {
    currentMode = mode;
    const chat = document.getElementById("chat");
    const studio = document.getElementById("imageStudio");
    const input = document.querySelector(".input-area");
    const welcome = document.getElementById("welcome");

    document.querySelectorAll(".nav-item").forEach(btn => btn.classList.toggle("active", btn.dataset.mode === mode));

    if (mode === "images") {
        chat.style.display = "none";
        input.style.display = "none";
        studio.classList.add("open");
    } else {
        studio.classList.remove("open");
        chat.style.display = "block";
        input.style.display = "block";
        if (!document.getElementById("messages").children.length) welcome.style.display = "flex";
    }
}

function useImageExample() {
    document.getElementById("imagePrompt").value = "A premium futuristic city floating above a bright sky ocean, cinematic lighting, realistic glass architecture, elegant blue and white palette, high detail";
}

async function generateImage() {
    const prompt = document.getElementById("imagePrompt").value.trim();
    const result = document.getElementById("imageResult");
    if (!prompt) return;
    result.innerHTML = '<div class="glass-button">Creating your image…</div>';

    const form = new FormData();
    form.append("prompt", prompt);
    form.append("conversation_id", conversationId);

    try {
        const response = await fetch("/api/generate-image", { method:"POST", body:form });
        const data = await response.json();
        if (data.conversation_id) conversationId = data.conversation_id;
        if (!response.ok || data.error) throw new Error(data.error || "Image generation failed");
        result.innerHTML = `<img src="${data.url}" alt="ASRIT generated image" loading="eager"><div style="color:#668096">Generated by ASRIT · ${escapeHtml(prompt)}</div>`;
        loadHistory();
    } catch (error) {
        result.innerHTML = `<div class="glass-button">${escapeHtml(error.message || error)}</div>`;
    }
}

function openFeature(name) {
    const labels = {
        library:"Library", scheduled:"Scheduled", plugins:"Plugins", projects:"Projects", codex:"Codex"
    };
    const text = labels[name] || name;
    addMessage("assistant", `${text} is ready as an ASRIT workspace. Connect the corresponding service/tool in your backend to make this section action-enabled.`);
    showMode("chat");
    closeSidebarMobile();
}

function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[c]));
}

/* =========================================================
   CLIPBOARD IMAGE PASTE — Ctrl+V / Cmd+V
========================================================= */
window.addEventListener("paste", async event => {
    const items = Array.from(event.clipboardData?.items || []);
    const imageItem = items.find(item => item.type.startsWith("image/"));
    if (!imageItem) return;

    event.preventDefault();
    const file = imageItem.getAsFile();
    if (!file) return;
    pastedImageFile = new File([file], `clipboard-${Date.now()}.png`, {type:file.type || "image/png"});

    const hint = document.getElementById("pasteHint");
    hint.classList.add("show");
    setTimeout(() => hint.classList.remove("show"), 1800);
    addMessage("user", "📋 Pasted image");
    await uploadImageBlob(pastedImageFile);
});

async function uploadImageBlob(file) {
    const form = new FormData();
    form.append("image", file);
    form.append("message", "Analyze this pasted image and explain what is visible.");
    form.append("conversation_id", conversationId);
    try {
        const response = await fetch("/api/image", {method:"POST", body:form});
        const data = await response.json();
        if (data.conversation_id) conversationId = data.conversation_id;
        addMessage("assistant", data.answer || "I couldn't analyze that image.");
        loadHistory();
    } catch (error) {
        addMessage("assistant", "Image paste error: " + error);
    }
}

/* =========================================================
   SEND MESSAGE
========================================================= */

async function sendMessage() {

    const input =
        document.getElementById(
            "messageInput"
        );

    const text =
        input.value.trim();

    if (!text) {
        return;
    }

    input.value = "";

    addMessage(
        "user",
        text
    );

    document.getElementById(
        "statusText"
    ).textContent =
        "ASRIT is thinking...";

    const thinking =
        document.createElement(
            "div"
        );

    thinking.className =
        "message assistant";

    thinking.innerHTML =
        `
        <div class="message-bubble">
            ASRIT is thinking...
        </div>
        `;

    document
        .getElementById("messages")
        .appendChild(
            thinking
        );

    try {

        const form =
            new FormData();

        form.append(
            "message",
            text
        );

        form.append(
            "conversation_id",
            conversationId
        );

        const response =
            await fetch(
                "/api/chat",
                {
                    method: "POST",
                    body: form
                }
            );

        const data =
            await response.json();

        thinking.remove();

        conversationId =
            data.conversation_id;

        addMessage(
            "assistant",
            data.answer
        );

        speak(
            data.answer
        );

        loadHistory();

    } catch(error) {

        thinking.remove();

        addMessage(
            "assistant",
            "ASRIT connection error: " +
            error
        );

    }

    document.getElementById(
        "statusText"
    ).textContent =
        "ASRIT online";

}


/* =========================================================
   ENTER KEY
========================================================= */

function handleKey(event) {

    if (
        event.key === "Enter" &&
        !event.shiftKey
    ) {

        event.preventDefault();

        sendMessage();

    }

}


/* =========================================================
   FILE UPLOAD
========================================================= */

document
    .getElementById("fileInput")
    .addEventListener(
        "change",
        async function() {

            const files =
                Array.from(
                    this.files
                );

            for (
                const file
                of files
            ) {

                await uploadFile(
                    file
                );

            }

            this.value = "";

        }
    );


async function uploadFile(file) {

    addMessage(
        "user",
        "📎 " + file.name
    );

    document.getElementById(
        "statusText"
    ).textContent =
        "ASRIT is reading " +
        file.name + "...";

    const question =
        prompt(
            "What should ASRIT do with this file?",
            "Analyze and summarize this file."
        )
        ||
        "Analyze this file.";

    const form =
        new FormData();

    form.append(
        "file",
        file
    );

    form.append(
        "message",
        question
    );

    form.append(
        "conversation_id",
        conversationId
    );

    try {

        const response =
            await fetch(
                "/api/upload",
                {
                    method: "POST",
                    body: form
                }
            );

        const data =
            await response.json();

        conversationId =
            data.conversation_id;

        addMessage(
            "assistant",
            data.answer
        );

        speak(
            data.answer
        );

        loadHistory();

    } catch(error) {

        addMessage(
            "assistant",
            "File processing error: " +
            error
        );

    }

    document.getElementById(
        "statusText"
    ).textContent =
        "ASRIT online";

}


/* =========================================================
   CAMERA
========================================================= */

async function openCamera() {

    const panel =
        document.getElementById(
            "cameraPanel"
        );

    const video =
        document.getElementById(
            "camera"
        );

    try {

        cameraStream =
            await navigator
                .mediaDevices
                .getUserMedia(
                    {
                        video: true,
                        audio: false
                    }
                );

        video.srcObject =
            cameraStream;

        panel.classList.add(
            "open"
        );

    } catch(error) {

        alert(
            "Camera permission was denied or camera is unavailable.\n\n" +
            error
        );

    }

}


function closeCamera() {

    if (cameraStream) {

        cameraStream
            .getTracks()
            .forEach(
                track =>
                    track.stop()
            );

        cameraStream = null;

    }

    document
        .getElementById(
            "cameraPanel"
        )
        .classList
        .remove("open");

}


/* =========================================================
   CAPTURE CAMERA IMAGE
========================================================= */

async function analyzeCamera() {

    const video =
        document.getElementById(
            "camera"
        );

    if (
        !video.videoWidth ||
        !video.videoHeight
    ) {

        alert(
            "Camera is not ready yet."
        );

        return;

    }

    const canvas =
        document.createElement(
            "canvas"
        );

    canvas.width =
        video.videoWidth;

    canvas.height =
        video.videoHeight;

    const context =
        canvas.getContext(
            "2d"
        );

    context.drawImage(
        video,
        0,
        0
    );

    canvas.toBlob(
        async function(blob) {

            const question =
                prompt(
                    "Ask ASRIT about the camera view:",
                    "What do you see?"
                )
                ||
                "What do you see?";

            const form =
                new FormData();

            form.append(
                "image",
                blob,
                "camera.jpg"
            );

            form.append(
                "message",
                question
            );

            form.append(
                "conversation_id",
                conversationId
            );

            document.getElementById(
                "statusText"
            ).textContent =
                "ASRIT is looking...";

            try {

                const response =
                    await fetch(
                        "/api/image",
                        {
                            method: "POST",
                            body: form
                        }
                    );

                const data =
                    await response.json();

                conversationId =
                    data.conversation_id;

                closeCamera();

                addMessage(
                    "user",
                    "📷 " + question
                );

                addMessage(
                    "assistant",
                    data.answer
                );

                speak(
                    data.answer
                );

                loadHistory();

            } catch(error) {

                alert(
                    "Camera analysis failed:\n" +
                    error
                );

            }

            document.getElementById(
                "statusText"
            ).textContent =
                "ASRIT online";

        },
        "image/jpeg",
        .85
    );

}


/* =========================================================
   VOICE RECOGNITION
========================================================= */

function toggleVoice() {

    const SpeechRecognition =
        window.SpeechRecognition ||
        window.webkitSpeechRecognition;

    if (!SpeechRecognition) {

        alert(
            "Your browser does not support speech recognition."
        );

        return;

    }

    if (isRecording) {

        if (recognition) {
            recognition.stop();
        }

        return;

    }

    recognition =
        new SpeechRecognition();

    recognition.lang =
        "en-IN";

    recognition.continuous =
        false;

    recognition.interimResults =
        false;

    recognition.onstart =
        function() {

            isRecording = true;

            document
                .getElementById(
                    "micButton"
                )
                .classList
                .add("recording");

            document.getElementById(
                "statusText"
            ).textContent =
                "Listening...";

        };

    recognition.onresult =
        function(event) {

            const text =
                event
                    .results[0][0]
                    .transcript;

            document
                .getElementById(
                    "messageInput"
                )
                .value =
                text;

            sendMessage();

        };

    recognition.onerror =
        function(event) {

            console.error(
                event.error
            );

        };

    recognition.onend =
        function() {

            isRecording = false;

            document
                .getElementById(
                    "micButton"
                )
                .classList
                .remove(
                    "recording"
                );

            document.getElementById(
                "statusText"
            ).textContent =
                "ASRIT online";

        };

    recognition.start();

}


/* =========================================================
   SPEECH
========================================================= */

function speak(text) {

    if (
        !("speechSynthesis" in window)
    ) {

        return;

    }

    window.speechSynthesis.cancel();

    const utterance =
        new SpeechSynthesisUtterance(
            text
        );

    utterance.lang =
        "en-IN";

    utterance.rate =
        .98;

    utterance.pitch =
        1.0;

    window.speechSynthesis.speak(
        utterance
    );

}



/* =========================================================
   DIRECT IMAGE PASTE
========================================================= */

document.addEventListener("paste", async function(event) {
    const items = Array.from(event.clipboardData?.items || []);
    const imageItem = items.find(item => item.type.startsWith("image/"));

    if (!imageItem) return;

    const file = imageItem.getAsFile();
    if (!file) return;

    event.preventDefault();
    showPasteHint("Pasted image — ASRIT is analyzing it...");
    await uploadPastedImage(file);
});

async function uploadPastedImage(file) {
    const question = "Analyze this pasted image and tell me what you can see.";

    addMessage("user", "🖼️ Pasted image");

    const form = new FormData();
    form.append("image", file, "pasted-image.png");
    form.append("message", question);
    form.append("conversation_id", conversationId);

    document.getElementById("statusText").textContent =
        "ASRIT is analyzing the pasted image...";

    try {
        const response = await fetch("/api/image", {
            method: "POST",
            body: form
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.detail || data.answer || "Image request failed");
        }

        conversationId = data.conversation_id;
        addMessage("assistant", data.answer);
        speak(data.answer);
        loadHistory();
    } catch (error) {
        addMessage(
            "assistant",
            "Image processing error: " + (error.message || error)
        );
    } finally {
        document.getElementById("statusText").textContent = "ASRIT online";
    }
}

function showPasteHint(text) {
    let el = document.getElementById("pasteHint");
    if (!el) {
        el = document.createElement("div");
        el.id = "pasteHint";
        el.className = "paste-hint";
        document.body.appendChild(el);
    }
    el.textContent = text;
    el.classList.add("show");
    clearTimeout(el._timer);
    el._timer = setTimeout(() => el.classList.remove("show"), 2200);
}




function liveSetState(text, connected = false) {
    const state = document.getElementById("liveState");
    const pill = document.getElementById("liveStatusPill");

    if (state) state.textContent = text;
    if (pill) {
        pill.textContent = connected ? "● Live" : "● " + text;
    }
}

function base64FromBytes(bytes) {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
}

function floatTo16BitPCM(float32, fromRate, toRate) {
    let samples = float32;

    if (fromRate !== toRate) {
        const ratio = fromRate / toRate;
        const newLength = Math.round(samples.length / ratio);
        const result = new Float32Array(newLength);

        let offsetResult = 0;
        let offsetBuffer = 0;

        while (offsetResult < result.length) {
            const nextOffsetBuffer = Math.min(
                Math.round((offsetResult + 1) * ratio),
                samples.length
            );

            let accum = 0;
            let count = 0;

            for (
                let i = offsetBuffer;
                i < nextOffsetBuffer;
                i++
            ) {
                accum += samples[i];
                count++;
            }

            result[offsetResult] = count ? accum / count : 0;
            offsetResult++;
            offsetBuffer = nextOffsetBuffer;
        }

        samples = result;
    }

    const pcm = new Int16Array(samples.length);

    for (let i = 0; i < samples.length; i++) {
        const sample = Math.max(-1, Math.min(1, samples[i]));
        pcm[i] = sample < 0
            ? sample * 0x8000
            : sample * 0x7fff;
    }

    return new Uint8Array(pcm.buffer);
}

function playLivePCM(base64) {
    try {
        if (!liveAudioContext) return;

        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);

        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }

        const pcm = new Int16Array(
            bytes.buffer,
            bytes.byteOffset,
            Math.floor(bytes.byteLength / 2)
        );

        const audioBuffer = liveAudioContext.createBuffer(
            1,
            pcm.length,
            24000
        );

        const channel = audioBuffer.getChannelData(0);

        for (let i = 0; i < pcm.length; i++) {
            channel[i] = pcm[i] / 32768;
        }

        const source = liveAudioContext.createBufferSource();
        source.buffer = audioBuffer;
        source.connect(liveAudioContext.destination);

        const nowTime = liveAudioContext.currentTime;
        if (liveNextAudioTime < nowTime + 0.03) {
            liveNextAudioTime = nowTime + 0.03;
        }

        source.start(liveNextAudioTime);
        liveNextAudioTime += audioBuffer.duration;
    } catch (error) {
        console.error("Live audio playback error:", error);
    }
}

function handleLiveServerMessage(data) {
    if (data.type === "error") {
        liveSetState(data.message || "Live error");
        console.error(data);
        return;
    }

    if (data.setupComplete) {
        liveSetState("You can talk now", true);
        return;
    }

    const content = data.serverContent;
    if (!content) return;

    if (content.inputTranscription?.text) {
        liveSetState("Listening…", true);
    }

    if (content.outputTranscription?.text) {
        liveSetState("ASRIT speaking…", true);
    }

    const parts = content.modelTurn?.parts || [];

    for (const part of parts) {
        const inline = part.inlineData;
        if (!inline?.data) continue;

        const mime = inline.mimeType || "";
        if (mime.startsWith("audio/pcm")) {
            playLivePCM(inline.data);
            liveSetState("ASRIT speaking…", true);
        }
    }

    if (content.turnComplete) {
        liveSetState("Listening…", true);
    }
}

async function openCamera() {
    if (liveSocket || cameraStream) return;

    const panel = document.getElementById("cameraPanel");
    const video = document.getElementById("camera");

    liveClosing = false;
    liveSetState("Opening camera and microphone…", false);
    panel.classList.add("open");

    try {
        cameraStream = await navigator.mediaDevices.getUserMedia({
            video: {
                facingMode: "user",
                width: { ideal: 1280 },
                height: { ideal: 720 }
            },
            audio: {
                echoCancellation: true,
                noiseSuppression: true,
                autoGainControl: true
            }
        });

        video.srcObject = cameraStream;
        liveMicEnabled = true;
        liveCameraEnabled = true;

        await connectLiveSocket();
        await startLiveAudio();

        startLiveVideoFrames();
        updateLiveControls();

    } catch (error) {
        console.error(error);
        liveSetState("Permission or device error", false);

        alert(
            "ASRIT Live could not start.\n\n" +
            "Please allow camera + microphone permissions and use HTTPS (or localhost).\n\n" +
            (error.message || error)
        );

        closeCamera();
    }
}

async function startLiveAudio() {
    if (!cameraStream) return;

    liveAudioContext = new (
        window.AudioContext ||
        window.webkitAudioContext
    )();

    await liveAudioContext.resume();

    liveSource = liveAudioContext.createMediaStreamSource(cameraStream);

    // ScriptProcessor is intentionally used here so this one-file app
    // needs no separate AudioWorklet file.
    liveProcessor = liveAudioContext.createScriptProcessor(
        4096,
        1,
        1
    );

    liveProcessor.onaudioprocess = function(event) {
        if (!liveMicEnabled) return;
        if (!liveSocket || liveSocket.readyState !== WebSocket.OPEN) return;

        const input = event.inputBuffer.getChannelData(0);
        const pcmBytes = floatTo16BitPCM(
            input,
            liveAudioContext.sampleRate,
            16000
        );

        liveSocket.send(JSON.stringify({
            type: "audio",
            data: base64FromBytes(pcmBytes)
        }));
    };

    liveSource.connect(liveProcessor);
    liveProcessor.connect(liveAudioContext.destination);
}

async function connectLiveSocket() {
    const protocol = location.protocol === "https:" ? "wss:" : "ws:";
    liveSocket = new WebSocket(
        protocol + "//" + location.host + "/ws/live"
    );

    await new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
            reject(new Error("Live connection timed out."));
        }, 12000);

        liveSocket.onopen = () => {
            clearTimeout(timeout);
            resolve();
        };

        liveSocket.onerror = () => {
            clearTimeout(timeout);
            reject(new Error("Live WebSocket connection failed."));
        };
    });

    liveSocket.onmessage = event => {
        try {
            handleLiveServerMessage(JSON.parse(event.data));
        } catch (error) {
            console.error("Invalid Live response:", error);
        }
    };

    liveSocket.onclose = () => {
        if (!liveClosing) {
            liveSetState("Live disconnected", false);
        }
    };

    liveSocket.onerror = error => {
        console.error("Live WebSocket error:", error);
        liveSetState("Live connection error", false);
    };
}

function startLiveVideoFrames() {
    stopLiveVideoFrames();

    const video = document.getElementById("camera");
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", {
        alpha: false,
        desynchronized: true
    });

    liveVideoTimer = setInterval(() => {
        if (!liveCameraEnabled) return;
        if (!liveSocket || liveSocket.readyState !== WebSocket.OPEN) return;
        if (!video.videoWidth || !video.videoHeight) return;

        const maxWidth = 640;
        const scale = Math.min(1, maxWidth / video.videoWidth);

        canvas.width = Math.max(1, Math.round(video.videoWidth * scale));
        canvas.height = Math.max(1, Math.round(video.videoHeight * scale));

        context.drawImage(
            video,
            0,
            0,
            canvas.width,
            canvas.height
        );

        canvas.toBlob(
            blob => {
                if (!blob) return;
                if (!liveSocket || liveSocket.readyState !== WebSocket.OPEN) return;

                const reader = new FileReader();

                reader.onload = () => {
                    const base64 = reader.result.split(",")[1];

                    liveSocket.send(JSON.stringify({
                        type: "video",
                        data: base64
                    }));
                };

                reader.readAsDataURL(blob);
            },
            "image/jpeg",
            0.62
        );
    }, 1000);
}

function stopLiveVideoFrames() {
    if (liveVideoTimer) {
        clearInterval(liveVideoTimer);
        liveVideoTimer = null;
    }
}

function toggleLiveMic() {
    if (!cameraStream) return;

    liveMicEnabled = !liveMicEnabled;

    const tracks = cameraStream.getAudioTracks();
    tracks.forEach(track => {
        track.enabled = liveMicEnabled;
    });

    updateLiveControls();

    if (liveSocket?.readyState === WebSocket.OPEN && !liveMicEnabled) {
        liveSocket.send(JSON.stringify({ type: "audio_end" }));
    }

    liveSetState(
        liveMicEnabled ? "Listening…" : "Microphone muted",
        true
    );
}

function toggleLiveCamera() {
    if (!cameraStream) return;

    liveCameraEnabled = !liveCameraEnabled;

    const tracks = cameraStream.getVideoTracks();
    tracks.forEach(track => {
        track.enabled = liveCameraEnabled;
    });

    updateLiveControls();

    liveSetState(
        liveCameraEnabled ? "Camera live" : "Camera paused",
        true
    );
}

function updateLiveControls() {
    const mic = document.getElementById("liveMicButton");
    const cam = document.getElementById("liveCameraButton");

    if (mic) {
        mic.textContent = liveMicEnabled ? "🎙" : "🔇";
        mic.classList.toggle("off", !liveMicEnabled);
    }

    if (cam) {
        cam.textContent = liveCameraEnabled ? "📷" : "🚫";
        cam.classList.toggle("off", !liveCameraEnabled);
    }
}

function closeCamera() {
    liveClosing = true;

    stopLiveVideoFrames();

    if (liveSocket) {
        try {
            if (liveSocket.readyState === WebSocket.OPEN) {
                liveSocket.send(JSON.stringify({ type: "close" }));
            }
            liveSocket.close();
        } catch (_) {}
        liveSocket = null;
    }

    if (liveProcessor) {
        try { liveProcessor.disconnect(); } catch (_) {}
        liveProcessor = null;
    }

    if (liveSource) {
        try { liveSource.disconnect(); } catch (_) {}
        liveSource = null;
    }

    if (liveAudioContext) {
        try { liveAudioContext.close(); } catch (_) {}
        liveAudioContext = null;
    }

    if (cameraStream) {
        cameraStream.getTracks().forEach(track => track.stop());
        cameraStream = null;
    }

    document.getElementById("camera").srcObject = null;
    document.getElementById("cameraPanel").classList.remove("open");

    liveNextAudioTime = 0;
    liveClosing = false;

    document.getElementById("statusText").textContent = "ASRIT online";
}


</script>

</body>

</html>
"""


# ============================================================
# HTML ROUTE
# ============================================================

@app.get("/", response_class=HTMLResponse)
async def home():

    return HTML


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":

    import uvicorn

    print()
    print("=" * 60)
    print("                 ASRIT AI")
    print("=" * 60)
    print()

    if not GEMINI_API_KEY:

        print(
            "⚠️ GEMINI API KEY NOT SET"
        )

        print()
        print(
            'Open asrit.py and replace:'
        )

        print()
        print(
            'GEMINI_API_KEY = "your_api_key"'
        )

        print()
        print(
            'with your actual Gemini API key.'
        )

    else:

        print(
            "✓ Gemini API key detected"
        )

    print()
    print(
        "Starting ASRIT..."
    )

    print()
    print(
        "Open:"
    )

    print(
        "http://127.0.0.1:8000"
    )

    print()
    print("=" * 60)
    print()

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "8000")),
        reload=False
    )
