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
import urllib.request
import urllib.error
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
LIVE_MODEL = os.environ.get("ASRIT_LIVE_MODEL", "gemini-3.8-live")

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

    db.execute("""
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            prompt TEXT NOT NULL,
            run_at TEXT NOT NULL,
            repeat TEXT NOT NULL DEFAULT 'once',
            enabled INTEGER NOT NULL DEFAULT 1,
            last_run TEXT,
            created_at TEXT NOT NULL
        )
    """)

    db.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT NOT NULL DEFAULT '',
            content TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    db.execute("""
        CREATE TABLE IF NOT EXISTS plugins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            description TEXT NOT NULL,
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)

    for plugin_name, plugin_desc in [
        ("Image Generator", "Generate images with Gemini."),
        ("File Analyzer", "Analyze uploaded documents and images."),
        ("ASRIT Live", "Real-time camera, microphone and audio replies."),
        ("Codex", "Gemini-powered coding workspace."),
    ]:
        db.execute("""
            INSERT OR IGNORE INTO plugins (name, description, enabled, created_at)
            VALUES (?, ?, 1, ?)
        """, (plugin_name, plugin_desc, datetime.now().isoformat(timespec="seconds")))

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


async def _generate_image_rest(prompt):
    """Fallback for google-genai versions where Interactions is unavailable."""
    if not GEMINI_API_KEY:
        raise RuntimeError("GEMINI_API_KEY is not configured.")

    payload = json.dumps({
        "model": IMAGE_MODEL,
        "input": [{"type": "text", "text": prompt}],
        "response_format": {"type": "image", "aspect_ratio": "1:1", "image_size": "1K"},
    }).encode("utf-8")

    def request():
        req = urllib.request.Request(
            "https://generativelanguage.googleapis.com/v1beta/interactions",
            data=payload,
            headers={
                "x-goog-api-key": GEMINI_API_KEY,
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=180) as response:
            return json.loads(response.read().decode("utf-8"))

    return await asyncio.to_thread(request)


async def generate_image(prompt, conversation_id=0):
    if not GEMINI_API_KEY:
        return {"error": "Please configure GEMINI_API_KEY first."}

    prompt = (prompt or "").strip()
    if not prompt:
        return {"error": "Please describe the image you want."}

    try:
        interaction = None
        # Prefer the current google-genai Interactions API.
        if client is not None and hasattr(client, "interactions"):
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
            except Exception as sdk_error:
                print("Image SDK fallback:", sdk_error)

        if interaction is None:
            interaction = await _generate_image_rest(prompt)

        image_data = None
        output_image = interaction.get("output_image") if isinstance(interaction, dict) else getattr(interaction, "output_image", None)
        if isinstance(output_image, dict):
            image_data = output_image.get("data")
        elif output_image is not None:
            image_data = getattr(output_image, "data", None)

        if not image_data:
            steps = interaction.get("steps", []) if isinstance(interaction, dict) else getattr(interaction, "steps", [])
            for step in steps or []:
                blocks = step.get("content", []) if isinstance(step, dict) else getattr(step, "content", [])
                for block in blocks or []:
                    btype = block.get("type") if isinstance(block, dict) else getattr(block, "type", "")
                    bdata = block.get("data") if isinstance(block, dict) else getattr(block, "data", None)
                    if btype == "image" and bdata:
                        image_data = bdata
                        break
                if image_data:
                    break

        # Some REST responses expose output_image nested in output.
        if not image_data and isinstance(interaction, dict):
            output = interaction.get("output") or []
            for block in output if isinstance(output, list) else []:
                if isinstance(block, dict) and block.get("type") == "image" and block.get("data"):
                    image_data = block["data"]
                    break

        if not image_data:
            return {"error": "Gemini returned no image data. Check the Image model/API access on your Gemini project."}

        filename = datetime.now().strftime("asrit_%Y%m%d_%H%M%S_%f.png")
        path = UPLOAD_FOLDER / filename
        with open(path, "wb") as output:
            output.write(base64.b64decode(image_data))

        if conversation_id:
            add_message(conversation_id, "assistant", f"[Generated image] {prompt}")

        return {"url": f"/media/{filename}", "prompt": prompt}

    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        return {"error": f"Image API error {error.code}: {detail[:1200]}"}
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

LIVE_MODEL = os.environ.get("ASRIT_LIVE_MODEL", "gemini-3.8-live")
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
        # Gemini 3.8 Live expects response modalities and speech configuration
        # inside generationConfig. Keeping this shape aligned with Google's
        # current Live WebSocket protocol avoids immediate setup disconnects.
        setup = {
            "setup": {
                "model": f"models/{LIVE_MODEL}",
                "systemInstruction": {"parts": [{"text": ASRIT_SYSTEM_PROMPT}]},
                "inputAudioTranscription": {},
                "outputAudioTranscription": {},
                "generationConfig": {
                    "responseModalities": ["AUDIO"],
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
            # Forward the exact provider response to the UI. Do not hide the
            # real reason behind a generic "Live disconnected" message.
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
# SCHEDULE WORKER
# ============================================================

async def scheduled_task_worker():
    while True:
        try:
            current = datetime.now()
            db = database()
            rows = db.execute("SELECT * FROM schedules WHERE enabled=1").fetchall()
            for row in rows:
                try:
                    due = datetime.fromisoformat(row["run_at"])
                except Exception:
                    continue
                if due > current:
                    continue

                schedule_id = row["id"]
                title = row["title"]
                prompt = row["prompt"]
                conversation_id = create_conversation(f"Scheduled · {title}")
                add_message(conversation_id, "user", f"[Scheduled task] {prompt}")
                answer = await ask_gemini(prompt, conversation_id)
                add_message(conversation_id, "assistant", answer)

                repeat = row["repeat"]
                if repeat == "daily":
                    next_run = due + __import__("datetime").timedelta(days=1)
                    db.execute("UPDATE schedules SET run_at=?, last_run=? WHERE id=?", (next_run.isoformat(timespec="minutes"), now(), schedule_id))
                elif repeat == "weekly":
                    next_run = due + __import__("datetime").timedelta(days=7)
                    db.execute("UPDATE schedules SET run_at=?, last_run=? WHERE id=?", (next_run.isoformat(timespec="minutes"), now(), schedule_id))
                else:
                    db.execute("UPDATE schedules SET enabled=0, last_run=? WHERE id=?", (now(), schedule_id))
            db.commit(); db.close()
        except Exception as error:
            print("Schedule worker error:", error)
        await asyncio.sleep(20)


@app.on_event("startup")
async def start_asrit_scheduler():
    asyncio.create_task(scheduled_task_worker())


# ============================================================
# ASRIT WORKSPACE APIs
# ============================================================

@app.get("/api/library")
async def library_api():
    db = database()
    rows = db.execute("SELECT id, conversation_id, filename, mime_type, path, created_at FROM files ORDER BY id DESC").fetchall()
    db.close()
    return [
        {"id": r["id"], "conversation_id": r["conversation_id"], "filename": r["filename"], "mime_type": r["mime_type"], "url": "/media/" + Path(r["path"]).name, "created_at": r["created_at"]}
        for r in rows
    ]


@app.get("/api/schedules")
async def schedules_api():
    db = database(); rows = db.execute("SELECT * FROM schedules ORDER BY run_at ASC").fetchall(); db.close()
    return [dict(r) for r in rows]


@app.post("/api/schedules")
async def create_schedule(payload: dict):
    title = str(payload.get("title", "ASRIT reminder")).strip() or "ASRIT reminder"
    prompt = str(payload.get("prompt", "")).strip()
    run_at = str(payload.get("run_at", "")).strip()
    repeat = str(payload.get("repeat", "once")).strip() or "once"
    if not prompt or not run_at:
        return JSONResponse({"error": "Prompt and run time are required."}, status_code=400)
    db = database()
    cur = db.execute("INSERT INTO schedules (title,prompt,run_at,repeat,created_at) VALUES (?,?,?,?,?)", (title,prompt,run_at,repeat,now()))
    db.commit(); db.close()
    return {"id": cur.lastrowid, "success": True}


@app.delete("/api/schedules/{schedule_id}")
async def delete_schedule(schedule_id: int):
    db = database(); db.execute("DELETE FROM schedules WHERE id=?", (schedule_id,)); db.commit(); db.close(); return {"success": True}


@app.post("/api/schedules/{schedule_id}/run")
async def run_schedule_now(schedule_id: int):
    db = database(); row = db.execute("SELECT * FROM schedules WHERE id=?", (schedule_id,)).fetchone(); db.close()
    if not row: return JSONResponse({"error":"Schedule not found"}, status_code=404)
    conversation_id = create_conversation(f"Scheduled · {row['title']}")
    add_message(conversation_id, "user", f"[Scheduled task] {row['prompt']}")
    answer = await ask_gemini(row["prompt"], conversation_id)
    add_message(conversation_id, "assistant", answer)
    return {"success": True, "conversation_id": conversation_id, "answer": answer}


@app.get("/api/projects")
async def projects_api():
    db = database(); rows = db.execute("SELECT * FROM projects ORDER BY updated_at DESC").fetchall(); db.close(); return [dict(r) for r in rows]


@app.post("/api/projects")
async def create_project(payload: dict):
    name = str(payload.get("name", "New project")).strip() or "New project"
    desc = str(payload.get("description", "")).strip()
    content = str(payload.get("content", "")).strip()
    timestamp = now(); db = database()
    cur = db.execute("INSERT INTO projects (name,description,content,created_at,updated_at) VALUES (?,?,?,?,?)", (name,desc,content,timestamp,timestamp))
    db.commit(); db.close(); return {"id": cur.lastrowid, "success": True}


@app.put("/api/projects/{project_id}")
async def update_project(project_id: int, payload: dict):
    db = database(); db.execute("UPDATE projects SET name=?, description=?, content=?, updated_at=? WHERE id=?", (str(payload.get("name","New project")), str(payload.get("description","")), str(payload.get("content","")), now(), project_id)); db.commit(); db.close(); return {"success": True}


@app.delete("/api/projects/{project_id}")
async def delete_project(project_id: int):
    db = database(); db.execute("DELETE FROM projects WHERE id=?", (project_id,)); db.commit(); db.close(); return {"success": True}


@app.get("/api/plugins")
async def plugins_api():
    db = database(); rows = db.execute("SELECT * FROM plugins ORDER BY name").fetchall(); db.close(); return [dict(r) for r in rows]


@app.post("/api/plugins/{plugin_id}/toggle")
async def toggle_plugin(plugin_id: int):
    db = database(); row = db.execute("SELECT enabled FROM plugins WHERE id=?", (plugin_id,)).fetchone()
    if not row:
        db.close(); return JSONResponse({"error":"Plugin not found"}, status_code=404)
    enabled = 0 if row["enabled"] else 1
    db.execute("UPDATE plugins SET enabled=? WHERE id=?", (enabled,plugin_id)); db.commit(); db.close(); return {"enabled": bool(enabled)}


@app.post("/api/codex")
async def codex_api(payload: dict):
    prompt = str(payload.get("prompt", "")).strip()
    code = str(payload.get("code", ""))
    if not prompt:
        return JSONResponse({"error": "Tell Codex what you want to build or change."}, status_code=400)
    answer = await ask_gemini(prompt + "\n\nExisting code:\n" + code, 0)
    return {"answer": answer}


@app.get("/api/search")
async def search_api(q: str = ""):
    q = q.strip()
    if not q: return []
    like = f"%{q}%"; db = database()
    rows = db.execute("""
        SELECT DISTINCT c.id, c.title, c.created_at, c.updated_at
        FROM conversations c LEFT JOIN messages m ON m.conversation_id=c.id
        WHERE c.title LIKE ? OR m.content LIKE ?
        ORDER BY c.updated_at DESC
    """, (like,like)).fetchall(); db.close(); return [dict(r) for r in rows]


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

    pointer-events: none;
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

    font-size: 14px;

    text-align: left;
}

.glass-button:hover {

    background:
        var(--glass-strong);

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
}

.history-item:hover {

    background:
        rgba(255,255,255,.06);
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

    scroll-behavior: auto;
}

.welcome {

    min-height: 100%;

    display: flex;

    flex-direction: column;

    justify-content: center;

    align-items: center;

    text-align: center;
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
    pointer-events:none;  z-index:500;
}
.paste-hint.show { opacity:1; transform:translateX(-50%) translateY(0); }
.pasted-image-card{padding:10px!important}.pasted-image-card img{display:block;max-width:min(520px,78vw);max-height:360px;object-fit:contain;border-radius:16px;margin-bottom:9px}.pasted-image-actions{display:flex;align-items:center;justify-content:space-between;gap:10px;font-size:12px}.pasted-image-actions .glass-button{padding:7px 11px!important;text-decoration:none!important;color:#2f211b!important}


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
.nav-item { display:flex; align-items:center; gap:12px; width:100%; padding:11px 13px; border:1px solid transparent; border-radius:15px; background:transparent; color:#24445e; cursor:pointer; text-align:left; font:inherit;  }
.nav-item:hover,.nav-item.active { background:rgba(255,255,255,.52); border-color:rgba(255,255,255,.75); box-shadow:0 8px 24px rgba(61,127,164,.10); }
.nav-icon { width:27px; text-align:center; font-size:17px; }
.sidebar-divider { height:1px; background:rgba(91,143,174,.18); margin:3px 0; }

/* Home / mode cards */
.quick-actions { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; width:min(760px,94%); margin-top:26px; }
.quick-card { border:1px solid rgba(255,255,255,.82); border-radius:22px; padding:16px; background:rgba(255,255,255,.42); color:#17324b; text-align:left; cursor:pointer; box-shadow:var(--shadow-soft),inset 0 1px 0 white; backdrop-filter:blur(22px);  }
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


/* =========================================================
   LIGHT LIQUID GLASS OVERRIDES
========================================================= */
:root{
  --background:linear-gradient(145deg,#f4c7aa 0%,#f7d7c2 42%,#e9b28f 100%);
  --glass:rgba(255,248,242,.66); --glass-strong:rgba(255,248,242,.84);
  --border:rgba(91,58,43,.20); --text:#2f211b; --muted:#6f5548;
  --accent:#c45b35; --accent2:#e58b62; --danger:#bd4350;
}
html,body{color:var(--text);background:var(--background)}
.sidebar{background:rgba(244,251,255,.66);border-right-color:rgba(64,113,150,.16);box-shadow:10px 0 40px rgba(47,111,145,.08)}
.topbar{background:rgba(255,255,255,.46);border-bottom-color:rgba(64,113,150,.14)}
.glass-button,.search,.input-glass,.camera-box,.workspace-panel,.image-studio-card,.quick-card{color:var(--text);background:rgba(255,255,255,.56);border-color:rgba(75,125,160,.18)}
.search::placeholder,textarea::placeholder,input::placeholder{color:#7590a3!important}
textarea,input,select{color:#17354b!important;background:rgba(255,255,255,.68)!important}
.nav-item,.history-item,.history-main,.history-title,.message-bubble,.status{color:#17354b!important}
.nav-item:hover,.history-item:hover{background:rgba(255,255,255,.72)}
.message-bubble{background:rgba(255,255,255,.72);border:1px solid rgba(72,120,155,.14);box-shadow:0 8px 24px rgba(55,100,130,.08)}
.message.user .message-bubble{background:linear-gradient(135deg,rgba(222,244,255,.95),rgba(255,255,255,.82))}
.eyebrow{font-size:11px;letter-spacing:.12em;color:#4d83a1}.workspace-head h2{color:#17354b}.workspace-head p,.workspace-row small,.empty-state{color:#5c788d}.workspace-panel{max-width:980px;margin:20px auto;padding:24px;border-radius:28px;box-shadow:0 18px 60px rgba(50,100,130,.12)}
.workspace-head{display:flex;justify-content:space-between;gap:16px;align-items:center}.workspace-head h2{margin:4px 0}.workspace-head p{margin:4px 0 16px}.workspace-list{display:grid;gap:10px}.workspace-row{display:flex;align-items:center;gap:12px;padding:13px 15px;border:1px solid rgba(75,125,160,.14);border-radius:18px;background:rgba(255,255,255,.48);color:#17354b;text-decoration:none}.workspace-row>div{flex:1}.workspace-row small{display:block;margin-top:4px;line-height:1.35}.schedule-form,.codex-grid{display:grid;gap:10px;margin:10px 0 18px}.schedule-form{grid-template-columns:1.2fr 1.2fr .7fr}.schedule-form textarea{grid-column:1/-1;min-height:90px}.codex-grid{grid-template-columns:1fr 1fr}.codex-grid textarea{min-height:260px}.codex-grid .glass-button{margin-top:10px;width:100%}.codex-result{white-space:pre-wrap;background:rgba(20,50,70,.05);border:1px solid rgba(75,125,160,.14);padding:16px;border-radius:18px;overflow:auto;color:#17354b}.project-card{padding:14px;border:1px solid rgba(75,125,160,.14);border-radius:20px;background:rgba(255,255,255,.48)}.project-card input,.project-card textarea{width:100%;border:1px solid rgba(75,125,160,.14);border-radius:14px;padding:11px;margin-bottom:8px}.project-card textarea{min-height:130px}.danger-button,.toggle{border:0;border-radius:12px;padding:9px 12px;cursor:pointer}.danger-button{background:rgba(217,74,92,.1);color:#b33c4d}.toggle{background:rgba(80,110,130,.12);color:#416176}.toggle.on{background:rgba(32,189,233,.16);color:#13718d}.history-item{display:flex;gap:5px;align-items:stretch}.history-main{border:0;background:transparent;text-align:left;flex:1;min-width:0;padding:10px;border-radius:12px;cursor:pointer}.history-main strong,.history-main small{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.history-main small{font-size:10px;color:#7891a2!important;margin-top:3px}.history-delete{width:34px;border:0;background:transparent;color:#7d96a8;border-radius:10px;cursor:pointer}.history-delete:hover{background:rgba(217,74,92,.1);color:#b33c4d}.empty-history{padding:14px;color:#7891a2;font-size:12px}.image-result img{max-width:100%;border-radius:22px;display:block;box-shadow:0 15px 40px rgba(35,95,130,.16)}
@media(max-width:760px){.schedule-form,.codex-grid{grid-template-columns:1fr}.workspace-panel{margin:10px;padding:16px}.workspace-head{align-items:flex-start;flex-direction:column}}

/* ASRIT fast peach glass polish */
html,body{background:linear-gradient(145deg,#efbfa0 0%,#f5d0b9 45%,#e5aa84 100%)!important;color:#2f211b!important}
body::before,body::after{animation:none!important;display:none!important}
.sidebar,.topbar,.input-glass,.glass-button,.search,.workspace-panel,.image-studio-card,.quick-card,.camera-box{background:rgba(255,248,242,.70)!important;border-color:rgba(91,58,43,.18)!important;color:#2f211b!important}
.nav-item,.history-item,.history-main,.history-title,.dock-btn,.status,.topbar{color:#3a281f!important}
.nav-item:hover,.history-item:hover{background:rgba(255,255,255,.50)!important}
.nav-item.active,.history-item.active{background:rgba(255,255,255,.66)!important;color:#2b1d17!important}
.message-bubble,.workspace-row,.codex-result{color:#2f211b!important}
.message.assistant .message-bubble{background:rgba(255,250,246,.82)!important;color:#2f211b!important}
.message.user .message-bubble{background:linear-gradient(135deg,rgba(255,236,222,.96),rgba(255,249,243,.88))!important;color:#2f211b!important}
.workspace-head h2,.workspace-head p,.workspace-row,.workspace-row small,.empty-state,.eyebrow{color:#3b2921!important}
#messageInput,textarea,input,select{color:#2f211b!important}
#messageInput::placeholder,textarea::placeholder,input::placeholder,.search::placeholder{color:#7b5e50!important}
.history{scrollbar-width:thin}
.history-item{margin:5px 0;border:1px solid transparent}
.history-main{padding:11px 10px}
.history-main strong{color:#35241d!important;font-size:13px}
.history-main small{color:#806556!important}
.history-delete{color:#8a6757!important}
.glass-button,.nav-item,.history-main,.history-delete,.dock-btn{border-color:rgba(91,58,43,.18)!important}
.paste-hint{background:rgba(57,35,25,.92)!important;color:#fff!important}
.history-refresh{border:0;background:transparent;color:#7a594b;cursor:pointer;font-size:16px;padding:2px 5px;border-radius:8px}.history-refresh:hover{background:rgba(120,70,45,.08)}
/* Keep history usable on small screens. */
@media(max-width:800px){.history{min-height:150px}.history-title{display:flex;justify-content:space-between;align-items:center}.history-item{padding:2px}.history-main{min-width:0}.sidebar.mobile-open{box-shadow:12px 0 45px rgba(75,42,28,.20)!important}}


/* ASRIT performance mode: completely static UI — no animations or transitions */
*, *::before, *::after {
    animation: none !important;
    transition: none !important;
    scroll-behavior: auto !important;
}
body::before, body::after { animation: none !important; }

/* ===== iOS / macOS INSPIRED MOTION ===== */
:root{--motion-fast:.16s;--motion-med:.28s;--motion-spring:cubic-bezier(.22,1,.36,1);--motion-soft:cubic-bezier(.25,.8,.25,1)}
@keyframes iosPageIn{from{opacity:0;transform:translateY(10px) scale(.985);filter:blur(2px)}to{opacity:1;transform:none;filter:blur(0)}}
@keyframes iosMessageIn{from{opacity:0;transform:translateY(8px) scale(.97)}to{opacity:1;transform:none}}
@keyframes iosPop{0%{opacity:0;transform:scale(.88)}70%{transform:scale(1.025)}100%{opacity:1;transform:scale(1)}}
@keyframes iosDockIn{from{opacity:0;transform:translateY(22px) scale(.96)}to{opacity:1;transform:none}}
@keyframes glassShimmer{0%{transform:translateX(-120%);opacity:0}35%{opacity:.16}70%,100%{transform:translateX(120%);opacity:0}}
@keyframes livePulse{0%,100%{box-shadow:0 0 0 0 rgba(255,90,80,.18)}50%{box-shadow:0 0 0 8px rgba(255,90,80,0)}}
@keyframes typingDot{0%,60%,100%{transform:translateY(0);opacity:.45}30%{transform:translateY(-4px);opacity:1}}
.app,.main-content,.sidebar,.workspace-panel{animation:iosPageIn .42s var(--motion-spring) both}
.sidebar{animation-delay:.02s}.main-content{animation-delay:.05s}
button,.glass-button,.nav-item,.history-item,.project-card,.plugin-card,.schedule-card,.library-card,.message-bubble,input,textarea,select{
 transition:transform var(--motion-fast) var(--motion-spring),box-shadow var(--motion-med) var(--motion-soft),background-color var(--motion-fast) ease,border-color var(--motion-fast) ease,opacity var(--motion-fast) ease;
 will-change:transform
}
button:hover,.glass-button:hover{transform:translateY(-2px) scale(1.015)}
button:active,.glass-button:active{transform:translateY(1px) scale(.965)}
.nav-item:hover,.history-item:hover,.project-card:hover,.plugin-card:hover,.schedule-card:hover,.library-card:hover{transform:translateY(-2px);box-shadow:0 10px 28px rgba(80,45,25,.12)}
.glass-button,button{position:relative;overflow:hidden}
.glass-button::after,button::after{content:"";position:absolute;inset:0;width:38%;background:linear-gradient(105deg,transparent,rgba(255,255,255,.22),transparent);transform:translateX(-150%);pointer-events:none}
.glass-button:hover::after,button:hover::after{animation:glassShimmer .7s var(--motion-soft) both}
.message,.message-row,.chat-message{animation:iosMessageIn .30s var(--motion-spring) both}
.pasted-image-card,.modal,.toast,.dropdown,.context-menu{animation:iosPop .24s var(--motion-spring) both}
.mobile-dock,.bottom-dock,.floating-dock{animation:iosDockIn .45s var(--motion-spring) .12s both}
.mobile-dock button:hover,.bottom-dock button:hover,.floating-dock button:hover{transform:translateY(-5px) scale(1.07)}
.mobile-dock button:active,.bottom-dock button:active,.floating-dock button:active{transform:translateY(-1px) scale(.92)}
.live-indicator,.live-dot,.voice-active{animation:livePulse 1.5s ease-in-out infinite}
.typing-dot,.loading-dot{animation:typingDot .9s ease-in-out infinite}
.typing-dot:nth-child(2),.loading-dot:nth-child(2){animation-delay:.12s}
.typing-dot:nth-child(3),.loading-dot:nth-child(3){animation-delay:.24s}
input:focus,textarea:focus,select:focus{transform:translateY(-1px);box-shadow:0 0 0 3px rgba(120,80,55,.12),0 8px 24px rgba(80,45,25,.08)}
@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.001ms!important;animation-iteration-count:1!important;transition-duration:.001ms!important;scroll-behavior:auto!important}}


/* ===== iOS HOME + DOCUMENTS ===== */

/* ===== ASRIT HISTORY SCREEN / READABILITY FIXES ===== */
.search{
  color:#111111!important;
  background:rgba(255,255,255,.82)!important;
  border-color:rgba(70,45,35,.24)!important;
}
.search::placeholder{color:#6b5146!important;opacity:1!important}
#messageInput,.message-bubble,.message-bubble *,.nav-item,.history-item,.history-main,
.history-title,.status,.live-state,.live-meta,.camera-controls,.dock-btn,.glass-button,
.workspace-panel,.workspace-panel *{
  color:#111111!important;
}
#messageInput::placeholder{color:#6b5146!important;opacity:1!important}
.sidebar .search{margin-bottom:14px!important}
.sidebar .side-nav{margin-top:4px!important;margin-bottom:14px!important}
.history-screen{
  position:fixed;inset:0;z-index:5200;display:none;overflow:auto;
  padding:clamp(18px,4vw,42px);
  background:linear-gradient(145deg,#efbfa0 0%,#f5d0b9 45%,#e5aa84 100%);
  color:#111111;
}
.history-screen.is-visible{display:block}
.history-shell{width:min(900px,100%);margin:0 auto}
.history-screen-head{
  display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:18px;
}
.history-screen-head h2{margin:0;color:#111111;font-size:clamp(24px,4vw,34px)}
.history-screen-head p{margin:4px 0 0;color:#3f3029}
.history-screen-search{
  width:100%;padding:14px 16px;border:1px solid rgba(80,50,38,.20);
  border-radius:18px;background:rgba(255,255,255,.72);color:#111111!important;
  outline:0;margin-bottom:18px;
}
.history-screen-list{display:grid;gap:10px}
.history-screen-item{
  display:flex;align-items:center;gap:10px;padding:8px;
  border:1px solid rgba(255,255,255,.62);border-radius:18px;
  background:rgba(255,255,255,.52);box-shadow:0 10px 28px rgba(70,40,20,.08);
}
.history-screen-item .history-main{padding:12px;background:transparent;border:0}
.history-screen-item .history-delete{flex:0 0 38px}
.history-screen-empty{text-align:center;padding:50px 20px;color:#4d3930}
.history-back{border:0;border-radius:14px;padding:10px 14px;background:rgba(255,255,255,.60);color:#111111!important;cursor:pointer}
@media(max-width:800px){
  .history-screen{padding:18px 14px 95px}
  .history-screen-head{align-items:flex-start}
  .history-screen-head h2{font-size:25px}
  .mobile-dock{gap:2px!important;padding:6px!important}
  .dock-btn{min-width:46px!important;padding:6px 3px!important;font-size:10px!important}
}

.ios-home-screen,.ios-docs-screen{
  position:fixed;inset:0;z-index:5000;display:none;overflow:auto;
  padding:clamp(22px,5vw,52px);
  background:linear-gradient(145deg,#efbfa0 0%,#f5d0b9 48%,#e5aa84 100%);
  color:#2f211b;
}
.ios-home-screen.is-visible,.ios-docs-screen.is-visible{display:block;animation:iosPageIn .38s var(--motion-spring) both}
.ios-home-top{display:flex;align-items:center;justify-content:space-between;max-width:1050px;margin:0 auto 28px}
.ios-home-greeting{font-size:clamp(30px,5vw,52px);font-weight:800;letter-spacing:-1.5px}
.ios-home-subtitle{opacity:.68;font-size:15px;margin-top:3px}
.ios-profile{width:48px;height:48px;border:0;border-radius:50%;font-size:22px;background:rgba(255,255,255,.5);backdrop-filter:blur(18px);box-shadow:0 10px 30px rgba(70,40,20,.12);cursor:pointer}
.ios-app-grid{max-width:1050px;margin:auto;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:20px}
.ios-app-icon{border:0;background:transparent;color:inherit;display:flex;flex-direction:column;align-items:center;gap:9px;font-size:14px;cursor:pointer}
.ios-icon{width:76px;height:76px;border-radius:22px;display:grid;place-items:center;font-size:32px;background:rgba(255,255,255,.48);border:1px solid rgba(255,255,255,.62);box-shadow:0 14px 30px rgba(70,40,20,.12),inset 0 1px rgba(255,255,255,.8);backdrop-filter:blur(18px);transition:transform .2s var(--motion-spring),box-shadow .2s}
.ios-app-icon:hover .ios-icon{transform:translateY(-5px) scale(1.06);box-shadow:0 20px 38px rgba(70,40,20,.18)}
.ios-app-icon:active .ios-icon{transform:scale(.9)}
.ios-widget-row{max-width:1050px;margin:38px auto 110px;display:grid;grid-template-columns:1fr 1fr;gap:18px}
.ios-widget{min-height:150px;text-align:left;border:1px solid rgba(255,255,255,.55);border-radius:28px;padding:24px;color:inherit;background:rgba(255,255,255,.36);backdrop-filter:blur(24px);box-shadow:0 18px 45px rgba(70,40,20,.12);cursor:pointer}
.ios-widget-title{display:block;font-size:13px;opacity:.62;margin-bottom:22px}.ios-widget strong{display:block;font-size:24px}.ios-widget small{display:block;margin-top:8px;opacity:.65}
.ios-home-dock{position:fixed;left:50%;bottom:20px;transform:translateX(-50%);display:flex;gap:10px;padding:10px 13px;border:1px solid rgba(255,255,255,.55);border-radius:28px;background:rgba(255,255,255,.38);backdrop-filter:blur(28px);box-shadow:0 18px 45px rgba(70,40,20,.16);animation:iosDockIn .45s var(--motion-spring) both}
.ios-home-dock button{width:56px;height:56px;border:0;border-radius:18px;background:rgba(255,255,255,.48);font-size:25px;cursor:pointer}
.ios-docs-screen{background:linear-gradient(145deg,#edbc9b,#f4d0b8,#e3a981)}
.ios-docs-header{max-width:1050px;margin:auto;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:15px}
.ios-docs-header>button{border:0;background:rgba(255,255,255,.4);border-radius:14px;padding:10px 14px;color:inherit;cursor:pointer;justify-self:start}
.ios-docs-header>button:last-child{justify-self:end}
.ios-docs-header strong,.ios-docs-header small{display:block;text-align:center}.ios-docs-header small{opacity:.6;font-size:12px;margin-top:3px}
.ios-docs-toolbar{max-width:1050px;margin:26px auto 18px;display:flex;gap:12px}
.ios-docs-toolbar input{flex:1;border:1px solid rgba(255,255,255,.55);border-radius:17px;padding:14px 17px;background:rgba(255,255,255,.45);color:#2f211b;outline:0}
.ios-docs-toolbar button{border:0;border-radius:17px;padding:0 20px;background:rgba(255,255,255,.55);color:#2f211b;cursor:pointer}
.ios-docs-list{max-width:1050px;margin:auto;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
.ios-doc-card{display:flex;align-items:center;gap:14px;padding:18px;border-radius:22px;background:rgba(255,255,255,.4);border:1px solid rgba(255,255,255,.5);box-shadow:0 10px 28px rgba(70,40,20,.08);transition:transform .2s var(--motion-spring),box-shadow .2s}
.ios-doc-card:hover{transform:translateY(-3px);box-shadow:0 16px 34px rgba(70,40,20,.14)}
.ios-doc-card-icon{font-size:30px}.ios-doc-card-name{font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ios-doc-card-meta{font-size:12px;opacity:.58;margin-top:4px}.ios-empty-docs{grid-column:1/-1;text-align:center;padding:60px 20px;opacity:.6}
@media(max-width:700px){.ios-home-screen,.ios-docs-screen{padding:20px}.ios-app-grid{grid-template-columns:repeat(4,1fr);gap:16px 8px}.ios-icon{width:62px;height:62px;border-radius:19px;font-size:26px}.ios-widget-row{grid-template-columns:1fr;margin-top:28px}.ios-widget{min-height:125px}.ios-docs-list{grid-template-columns:1fr}.ios-docs-toolbar{flex-direction:column}.ios-docs-toolbar button{height:46px}.ios-home-dock button{width:50px;height:50px}}

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
            title="Start a new conversation"
        >
            <span class="new-chat-plus">+</span>
            <span>New conversation</span>
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
             <button class="nav-item" data-mode="history" onclick="openHistoryScreen()"><span class="nav-icon">◷</span>History</button>
            <button class="nav-item" onclick="openFeature('library')"><span class="nav-icon">▤</span>Library</button>
            <button class="nav-item" onclick="openFeature('scheduled')"><span class="nav-icon">◷</span>Scheduled</button>
            <button class="nav-item" onclick="openFeature('plugins')"><span class="nav-icon">⌘</span>Plugins</button>
            <button class="nav-item" onclick="openFeature('projects')"><span class="nav-icon">□</span>Projects</button>
            <button class="nav-item" onclick="openFeature('codex')"><span class="nav-icon">{ }</span>Codex</button>
        </div>

        <div class="history">

            <div class="history-title">
                <span>Conversations</span>
                <button class="history-refresh" onclick="loadHistory()" title="Refresh history">↻</button>
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
                aria-label="Open menu"
            >
                <span class="menu-bars" aria-hidden="true"><i></i><i></i><i></i></span>
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
                    <div><div style="font-size:13px;color:#4d7188">ASRIT CREATIVE</div><h2 style="margin:3px 0 0">Create an image</h2></div>
                    <button class="icon-button" onclick="showMode('chat')">×</button>
                </div>
                <p style="color:#5c788d">Describe the image, poster, illustration, concept art, or edit you want.</p>
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
                    class="icon-button add-button"
                    onclick="document.getElementById('fileInput').click()"
                    title="Add files, photos or screenshots"
                    aria-label="Add files, photos or screenshots"
                >
                    <span class="plus-mark" aria-hidden="true">+</span>
                </button>

                <textarea
                    id="messageInput"
                    placeholder="Ask ASRIT anything..."
                    rows="1"
                    onkeydown="handleKey(event)"
                ></textarea>

                <button
                    id="micButton"
                    class="icon-button mic-button"
                    onclick="toggleVoice()"
                    title="Voice"
                    aria-label="Voice input"
                >
                    <span class="mic-mark" aria-hidden="true"></span>
                </button>

                <button
                    class="icon-button send-button"
                    onclick="sendMessage()"
                    title="Send message"
                    aria-label="Send message"
                >
                    <span class="send-mark" aria-hidden="true">↑</span>
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
    <button class="dock-btn" onclick="openHistoryScreen()"><span>◷</span>History</button>
    <button class="dock-btn" onclick="toggleSidebar()"><span>☰</span>Menu</button>
</nav>



<!-- Dedicated History column/screen -->
<section id="historyScreen" class="history-screen" aria-label="ASRIT History">
  <div class="history-shell">
    <div class="history-screen-head">
      <div>
        <h2>History</h2>
        <p>Your saved ASRIT conversations</p>
      </div>
      <button class="history-back" type="button" onclick="closeHistoryScreen()">‹ Back</button>
    </div>
    <input id="historyScreenSearch" class="history-screen-search" type="search"
           placeholder="Search conversation history..." autocomplete="off">
    <div id="historyScreenList" class="history-screen-list">
      <div class="history-screen-empty">Loading history…</div>
    </div>
  </div>
</section>

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
let lastUserPrompt = "";

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
let liveMuteGain = null;
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

        document.getElementById("historyScreenSearch")?.addEventListener("input", event => {
            loadHistoryScreen(event.target.value);
        });

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
    const list = document.getElementById("historyList");
    list.innerHTML = "";
    if (!items.length) { list.innerHTML = `<div class="empty-history">No conversations found</div>`; return; }
    for (const conversation of items) {
        const item = document.createElement("div");
        item.className = "history-item" + (conversation.id === conversationId ? " active" : "");
        const main = document.createElement("button"); main.className="history-main";
        main.innerHTML = `<strong>${escapeHtml(conversation.title || "New conversation")}</strong><small>${escapeHtml((conversation.updated_at||"").replace("T"," "))}</small>`;
        main.onclick = () => loadConversation(conversation.id);
        const del = document.createElement("button"); del.className="history-delete"; del.title="Delete conversation"; del.textContent="×";
        del.onclick = async e => { e.stopPropagation(); if(!confirm("Delete this conversation?")) return; await fetch(`/api/conversations/${conversation.id}`,{method:"DELETE"}); if(conversation.id===conversationId)newChat(); await loadHistory(); };
        item.append(main,del); list.appendChild(item);
    }
}

/* =========================================================
   SEARCH
========================================================= */

async function searchChats() {
    const query = document.getElementById("search").value.trim();
    if (!query) { renderHistory(allConversations); return; }
    try {
        const r = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
        renderHistory(await r.json());
    } catch (_) { renderHistory(allConversations.filter(x=>x.title.toLowerCase().includes(query.toLowerCase()))); }
}

function openHistoryScreen() {
    const screen = document.getElementById("historyScreen");
    if (!screen) return;
    screen.classList.add("is-visible");
    document.body.classList.add("history-open");
    closeSidebarMobile();
    loadHistoryScreen();
}

function closeHistoryScreen() {
    const screen = document.getElementById("historyScreen");
    if (screen) screen.classList.remove("is-visible");
    document.body.classList.remove("history-open");
}

async function loadHistoryScreen(query = "") {
    const list = document.getElementById("historyScreenList");
    if (!list) return;
    list.innerHTML = `<div class="history-screen-empty">Loading history…</div>`;
    try {
        if (!allConversations.length) {
            const r = await fetch("/api/conversations");
            allConversations = await r.json();
        }
        const q = String(query || "").trim().toLowerCase();
        const items = q
            ? allConversations.filter(x =>
                String(x.title || "").toLowerCase().includes(q) ||
                String(x.updated_at || "").toLowerCase().includes(q)
              )
            : allConversations;

        if (!items.length) {
            list.innerHTML = `<div class="history-screen-empty">No conversations found.</div>`;
            return;
        }

        list.innerHTML = "";
        for (const conversation of items) {
            const item = document.createElement("div");
            item.className = "history-screen-item";
            const main = document.createElement("button");
            main.className = "history-main";
            main.innerHTML =
                `<strong>${escapeHtml(conversation.title || "New conversation")}</strong>` +
                `<small>${escapeHtml((conversation.updated_at || "").replace("T", " "))}</small>`;
            main.onclick = () => {
                closeHistoryScreen();
                loadConversation(conversation.id);
            };
            const del = document.createElement("button");
            del.className = "history-delete";
            del.title = "Delete conversation";
            del.textContent = "×";
            del.onclick = async event => {
                event.stopPropagation();
                if (!confirm("Delete this conversation?")) return;
                await fetch(`/api/conversations/${conversation.id}`, {method:"DELETE"});
                if (conversation.id === conversationId) newChat();
                await loadHistory();
                await loadHistoryScreen(
                    document.getElementById("historyScreenSearch")?.value || ""
                );
            };
            item.append(main, del);
            list.appendChild(item);
        }
    } catch (error) {
        console.error(error);
        list.innerHTML = `<div class="history-screen-empty">History could not be loaded.</div>`;
    }
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

            if (message.role === "user") {
                lastUserPrompt = message.content || "";
            }

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

    const wrapper = document.createElement("div");
    wrapper.className = `message ${role}`;

    const bubble = document.createElement("div");
    bubble.className = "message-bubble";
    bubble.textContent = text;

    wrapper.appendChild(bubble);

    if (role === "assistant") {
        // Keep the user prompt that produced this answer so Try again
        // works even when several assistant messages are on screen.
        wrapper.dataset.prompt = lastUserPrompt || "";

        const actions = document.createElement("div");
        actions.className = "response-actions";
        actions.setAttribute("aria-label", "Response actions");

        actions.innerHTML = `
            <button class="response-action" data-action="copy" title="Copy response" aria-label="Copy response">
                <span class="ra-icon">▣</span>
            </button>
            <button class="response-action" data-action="like" title="Good response" aria-label="Good response">
                <span class="ra-icon">♧</span>
            </button>
            <button class="response-action" data-action="dislike" title="Bad response" aria-label="Bad response">
                <span class="ra-icon">♧</span>
            </button>
            <button class="response-action" data-action="share" title="Share response" aria-label="Share response">
                <span class="ra-icon">↗</span>
            </button>
            <button class="response-action" data-action="retry" title="Try again" aria-label="Try again">
                <span class="ra-icon">↻</span>
            </button>
            <button class="response-action response-more" data-action="more" title="More" aria-label="More response actions">
                <span class="ra-icon">•••</span>
            </button>
        `;

        actions.addEventListener("click", async (event) => {
            const button = event.target.closest(".response-action");
            if (!button) return;

            const action = button.dataset.action;

            if (action === "copy") {
                await copyResponseText(text, button);
            } else if (action === "like" || action === "dislike") {
                const siblingAction = action === "like" ? "dislike" : "like";
                button.classList.toggle("selected");
                const other = actions.querySelector(`[data-action="${siblingAction}"]`);
                if (button.classList.contains("selected")) other?.classList.remove("selected");
            } else if (action === "share") {
                await shareResponseText(text, button);
            } else if (action === "retry") {
                const prompt = wrapper.dataset.prompt || findPromptBefore(wrapper);
                if (prompt) {
                    retryAssistant(prompt, wrapper);
                } else {
                    flashResponseAction(button, "No prompt");
                }
            } else if (action === "more") {
                toggleResponseMoreMenu(wrapper, text, button);
            }
        });

        wrapper.appendChild(actions);
    }

    document.getElementById("messages").appendChild(wrapper);

    const chat = document.getElementById("chat");
    chat.scrollTop = chat.scrollHeight;
}

function findPromptBefore(assistantWrapper) {
    let node = assistantWrapper.previousElementSibling;
    while (node) {
        if (node.classList.contains("message") && node.classList.contains("user")) {
            return node.querySelector(".message-bubble")?.textContent?.trim() || "";
        }
        node = node.previousElementSibling;
    }
    return "";
}

function flashResponseAction(button, label = "Copied") {
    const original = button.innerHTML;
    button.innerHTML = `<span class="ra-label">${escapeHtml(label)}</span>`;
    setTimeout(() => {
        button.innerHTML = original;
    }, 1200);
}

async function copyResponseText(text, button) {
    try {
        await navigator.clipboard.writeText(text);
        flashResponseAction(button, "Copied");
    } catch (error) {
        const area = document.createElement("textarea");
        area.value = text;
        area.style.position = "fixed";
        area.style.opacity = "0";
        document.body.appendChild(area);
        area.select();
        document.execCommand("copy");
        area.remove();
        flashResponseAction(button, "Copied");
    }
}

async function shareResponseText(text, button) {
    try {
        if (navigator.share) {
            await navigator.share({
                title: "ASRIT response",
                text
            });
            flashResponseAction(button, "Shared");
        } else {
            await navigator.clipboard.writeText(text);
            flashResponseAction(button, "Copied");
        }
    } catch (error) {
        // User cancelling the native share sheet is not an error worth surfacing.
        if (error?.name !== "AbortError") {
            flashResponseAction(button, "Copied");
        }
    }
}

function toggleResponseMoreMenu(wrapper, text, button) {
    document.querySelectorAll(".response-more-menu").forEach(el => el.remove());

    const menu = document.createElement("div");
    menu.className = "response-more-menu";
    menu.innerHTML = `
        <button type="button" data-more-action="copy">Copy response</button>
        <button type="button" data-more-action="share">Share response</button>
        <button type="button" data-more-action="retry">Try again</button>
    `;

    menu.addEventListener("click", async (event) => {
        const item = event.target.closest("[data-more-action]");
        if (!item) return;
        const action = item.dataset.moreAction;
        menu.remove();

        if (action === "copy") await copyResponseText(text, button);
        if (action === "share") await shareResponseText(text, button);
        if (action === "retry") {
            const prompt = wrapper.dataset.prompt || findPromptBefore(wrapper);
            if (prompt) retryAssistant(prompt, wrapper);
        }
    });

    wrapper.appendChild(menu);
    requestAnimationFrame(() => menu.classList.add("open"));

    setTimeout(() => {
        const close = (event) => {
            if (!menu.contains(event.target) && event.target !== button) {
                menu.remove();
                document.removeEventListener("click", close);
            }
        };
        document.addEventListener("click", close);
    }, 0);
}

async function retryAssistant(prompt, oldWrapper) {
    if (!prompt) return;

    const input = document.getElementById("messageInput");
    if (input) input.value = "";

    oldWrapper?.classList.add("retrying");
    document.getElementById("statusText").textContent = "ASRIT is thinking...";

    const thinking = document.createElement("div");
    thinking.className = "message assistant";
    thinking.innerHTML = `<div class="message-bubble">ASRIT is thinking...</div>`;
    document.getElementById("messages").appendChild(thinking);

    try {
        const form = new FormData();
        form.append("message", prompt);
        form.append("conversation_id", conversationId);

        const response = await fetch("/api/chat", {
            method: "POST",
            body: form
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || data.answer || "Request failed");

        thinking.remove();
        conversationId = data.conversation_id;
        lastUserPrompt = prompt;
        addMessage("assistant", data.answer);
        speak(data.answer);
        loadHistory();
    } catch (error) {
        thinking.remove();
        addMessage("assistant", "ASRIT connection error: " + (error.message || error));
    } finally {
        oldWrapper?.classList.remove("retrying");
        document.getElementById("statusText").textContent = "ASRIT online";
    }
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
    if (mode === "history") {
        openHistoryScreen();
        return;
    }
    closeHistoryScreen();

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
        result.innerHTML = `<img src="${data.url}" alt="ASRIT generated image" loading="eager"><div style="color:#5c788d">Generated by ASRIT · ${escapeHtml(prompt)}</div>`;
        loadHistory();
    } catch (error) {
        result.innerHTML = `<div class="glass-button">${escapeHtml(error.message || error)}</div>`;
    }
}

async function openFeature(name) {
    closeSidebarMobile();
    document.getElementById("messages").innerHTML = "";
    document.getElementById("welcome").style.display = "none";
    showMode("chat");

    const messages = document.getElementById("messages");
    const panel = document.createElement("div");
    panel.className = "workspace-panel";
    messages.appendChild(panel);

    if (name === "library") {
        const data = await (await fetch("/api/library")).json();
        panel.innerHTML = `<div class="workspace-head"><div><span class="eyebrow">ASRIT WORKSPACE</span><h2>Library</h2><p>Your uploaded files and generated images.</p></div><button class="glass-button" onclick="document.getElementById('fileInput').click()">＋ Add file</button></div>`;
        const list = document.createElement("div"); list.className="workspace-list";
        list.innerHTML = data.length ? data.map(x=>`<a class="workspace-row" href="${x.url}" target="_blank"><span>▤</span><div><strong>${escapeHtml(x.filename)}</strong><small>${escapeHtml(x.mime_type||"file")} · ${escapeHtml(x.created_at)}</small></div><b>↗</b></a>`).join("") : `<div class="empty-state">No files yet. Add a photo, document, or generate an image.</div>`;
        panel.appendChild(list); return;
    }

    if (name === "scheduled") {
        const data = await (await fetch("/api/schedules")).json();
        panel.innerHTML = `<div class="workspace-head"><div><span class="eyebrow">ASRIT AUTOMATIONS</span><h2>Scheduled</h2><p>Create reminders/tasks that ASRIT keeps in its local database.</p></div></div><div class="schedule-form"><input id="schTitle" placeholder="Title"><input id="schWhen" type="datetime-local"><select id="schRepeat"><option value="once">Once</option><option value="daily">Daily</option><option value="weekly">Weekly</option></select><textarea id="schPrompt" placeholder="What should ASRIT remember or run?"></textarea><button class="glass-button send-button" onclick="createSchedule()">＋ Schedule</button></div><div id="scheduleList" class="workspace-list"></div>`;
        renderSchedules(data); return;
    }

    if (name === "plugins") {
        const data = await (await fetch("/api/plugins")).json();
        panel.innerHTML = `<div class="workspace-head"><div><span class="eyebrow">ASRIT TOOLS</span><h2>Plugins</h2><p>Built-in ASRIT tools with on/off controls.</p></div></div><div id="pluginList" class="workspace-list"></div>`;
        document.getElementById("pluginList").innerHTML = data.map(x=>`<div class="workspace-row"><span>⌘</span><div><strong>${escapeHtml(x.name)}</strong><small>${escapeHtml(x.description)}</small></div><button class="toggle ${x.enabled?'on':''}" onclick="togglePlugin(${x.id})">${x.enabled?'Enabled':'Disabled'}</button></div>`).join(""); return;
    }

    if (name === "projects") {
        const data = await (await fetch("/api/projects")).json();
        panel.innerHTML = `<div class="workspace-head"><div><span class="eyebrow">ASRIT WORKSPACE</span><h2>Projects</h2><p>Persistent project spaces stored in ASRIT.</p></div><button class="glass-button" onclick="createProject()">＋ New project</button></div><div id="projectList" class="workspace-list"></div>`;
        renderProjects(data); return;
    }

    if (name === "codex") {
        panel.innerHTML = `<div class="workspace-head"><div><span class="eyebrow">ASRIT CODEX</span><h2>Codex</h2><p>Describe a coding task and ASRIT will work against the code you provide.</p></div></div><div class="codex-grid"><textarea id="codexCode" placeholder="Paste your code here…"></textarea><div><textarea id="codexPrompt" placeholder="What should I build, debug, explain, or change?"></textarea><button class="glass-button send-button" onclick="runCodex()">Run Codex</button></div></div><pre id="codexResult" class="codex-result"></pre>`; return;
    }
}

async function createSchedule(){
    const payload={title:document.getElementById("schTitle").value,prompt:document.getElementById("schPrompt").value,run_at:document.getElementById("schWhen").value,repeat:document.getElementById("schRepeat").value};
    const r=await fetch("/api/schedules",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)}); const d=await r.json(); if(d.error){alert(d.error);return;} openFeature("scheduled");
}
function renderSchedules(data){ const el=document.getElementById("scheduleList"); el.innerHTML=data.length?data.map(x=>`<div class="workspace-row"><span>◷</span><div><strong>${escapeHtml(x.title)}</strong><small>${escapeHtml(x.run_at)} · ${escapeHtml(x.repeat)}<br>${escapeHtml(x.prompt)}</small></div><button class="glass-button" onclick="runSchedule(${x.id})">Run now</button><button class="danger-button" onclick="deleteSchedule(${x.id})">Delete</button></div>`).join(""): `<div class="empty-state">No scheduled tasks yet.</div>`; }
async function runSchedule(id){await fetch(`/api/schedules/${id}/run`,{method:"POST"});openFeature("scheduled");}
async function deleteSchedule(id){await fetch(`/api/schedules/${id}`,{method:"DELETE"});openFeature("scheduled");}
async function togglePlugin(id){await fetch(`/api/plugins/${id}/toggle`,{method:"POST"});openFeature("plugins");}
async function createProject(){ const name=prompt("Project name:","My ASRIT project"); if(!name)return; await fetch("/api/projects",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name})}); openFeature("projects"); }
function renderProjects(data){ const el=document.getElementById("projectList"); el.innerHTML=data.length?data.map(x=>`<div class="project-card"><input value="${escapeHtml(x.name)}" id="pn${x.id}"><textarea id="pc${x.id}" placeholder="Project notes / code…">${escapeHtml(x.content)}</textarea><div><button class="glass-button" onclick="saveProject(${x.id})">Save</button><button class="danger-button" onclick="deleteProject(${x.id})">Delete</button></div></div>`).join(""): `<div class="empty-state">No projects yet. Create one to keep code and notes together.</div>`; }
async function saveProject(id){await fetch(`/api/projects/${id}`,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:document.getElementById(`pn${id}`).value,content:document.getElementById(`pc${id}`).value})});openFeature("projects");}
async function deleteProject(id){await fetch(`/api/projects/${id}`,{method:"DELETE"});openFeature("projects");}
async function runCodex(){ const result=document.getElementById("codexResult"); result.textContent="Working…"; const r=await fetch("/api/codex",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt:document.getElementById("codexPrompt").value,code:document.getElementById("codexCode").value})}); const d=await r.json(); result.textContent=d.answer||d.error||"No response"; }


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

    const objectUrl = URL.createObjectURL(pastedImageFile);
    const hint = document.getElementById("pasteHint");
    hint.classList.add("show");
    setTimeout(() => hint.classList.remove("show"), 1800);

    // Show the pasted image in chat with an explicit Open button before analysis.
    const wrapper = document.createElement("div");
    wrapper.className = "message user";
    wrapper.innerHTML = `<div class="message-bubble pasted-image-card"><img src="${objectUrl}" alt="Pasted image" loading="eager"><div class="pasted-image-actions"><span>📋 Pasted image</span><a class="glass-button" href="${objectUrl}" target="_blank" rel="noopener">Open</a></div></div>`;
    document.getElementById("messages").appendChild(wrapper);
    document.getElementById("welcome").style.display = "none";
    document.getElementById("chat").scrollTop = document.getElementById("chat").scrollHeight;
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
    lastUserPrompt = text;

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
    liveSetState("Requesting camera + microphone…", false);
    panel.classList.add("open");

    try {
        if (!navigator.mediaDevices?.getUserMedia) {
            throw new Error("Camera/microphone access is unavailable in this browser. Use HTTPS or localhost.");
        }

        cameraStream = await navigator.mediaDevices.getUserMedia({
            video: {
                facingMode: "user",
                width: { ideal: 1280 },
                height: { ideal: 720 }
            },
            audio: {
                echoCancellation: true,
                noiseSuppression: true,
                autoGainControl: true,
                channelCount: 1
            }
        });

        video.srcObject = cameraStream;
        video.muted = true;
        await video.play().catch(() => {});
        liveMicEnabled = true;
        liveCameraEnabled = true;
        updateLiveControls();

        liveSetState("Connecting to ASRIT Live…", false);
        await connectLiveSocket();

        liveSetState("Starting microphone…", true);
        await startLiveAudio();

        startLiveVideoFrames();
        updateLiveControls();
        liveSetState("You can talk now", true);

    } catch (error) {
        console.error("ASRIT Live startup error:", error);
        liveSetState("Live failed — see the message below", false);

        alert(
            "ASRIT Live could not start.\n\n" +
            (error.message || error) +
            "\n\nIf the browser asks, allow camera + microphone access."
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
    liveMuteGain = liveAudioContext.createGain();
    liveMuteGain.gain.value = 0;
    liveProcessor.connect(liveMuteGain);
    liveMuteGain.connect(liveAudioContext.destination);
}

async function connectLiveSocket() {
    const protocol = location.protocol === "https:" ? "wss:" : "ws:";
    const url = protocol + "//" + location.host + "/ws/live";
    liveSocket = new WebSocket(url);

    await new Promise((resolve, reject) => {
        let setupReady = false;
        let settled = false;
        const timeout = setTimeout(() => {
            if (settled) return;
            settled = true;
            reject(new Error("Live setup timed out. Check the ASRIT server log and Gemini API key."));
        }, 20000);

        liveSocket.onopen = () => {
            liveSetState("Connected — waiting for Gemini…", false);
        };

        liveSocket.onmessage = event => {
            try {
                const data = JSON.parse(event.data);

                if (data.setupComplete) {
                    setupReady = true;
                    settled = true;
                    clearTimeout(timeout);
                    liveSetState("Live connected", true);
                    resolve();
                    return;
                }

                if (data.type === "error" || data.error) {
                    const message =
                        data.message ||
                        data.error?.message ||
                        data.error ||
                        "Gemini Live setup failed.";
                    if (!settled) {
                        settled = true;
                        clearTimeout(timeout);
                        reject(new Error(message));
                    } else {
                        liveSetState("Live error", false);
                    }
                    console.error("Gemini Live error:", data);
                    return;
                }

                if (setupReady) handleLiveServerMessage(data);
            } catch (error) {
                console.error("Invalid Live response:", error, event.data);
            }
        };

        liveSocket.onerror = () => {
            if (!settled) {
                settled = true;
                clearTimeout(timeout);
                reject(new Error("Live WebSocket connection failed before Gemini setup completed."));
            }
        };

        liveSocket.onclose = event => {
            if (!settled) {
                settled = true;
                clearTimeout(timeout);
                reject(new Error(
                    "Live WebSocket closed before setup completed." +
                    (event.reason ? " " + event.reason : "")
                ));
            } else if (!liveClosing) {
                liveSetState("Live disconnected", false);
            }
        };
    });

    // Keep one message handler after setup; setupComplete has already been consumed.
    liveSocket.onmessage = event => {
        try {
            const data = JSON.parse(event.data);
            if (data.type === "error" || data.error) {
                const message = data.message || data.error?.message || data.error || "Live error";
                liveSetState(message, false);
                console.error("Gemini Live error:", data);
                return;
            }
            handleLiveServerMessage(data);
        } catch (error) {
            console.error("Invalid Live response:", error, event.data);
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

    if (liveMuteGain) { try { liveMuteGain.disconnect(); } catch (_) {} liveMuteGain = null; }

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



document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeHistoryScreen();
});

</script>


<!-- iOS-style Home / Documents shell -->
<section id="iosHomeScreen" class="ios-home-screen" aria-label="ASRIT Home">
  <div class="ios-home-top">
    <div>
      <div class="ios-home-greeting">ASRIT AI</div>
      <div class="ios-home-subtitle">Good to see you</div>
    </div>
    <button class="ios-profile" type="button" onclick="document.body.classList.toggle('ios-home-focus')">✦</button>
  </div>

  <div class="ios-app-grid">
    <button class="ios-app-icon" data-home-action="chat"><span class="ios-icon">💬</span><b>Chat</b></button>
    <button class="ios-app-icon" data-home-action="images"><span class="ios-icon">🖼️</span><b>Images</b></button>
    <button class="ios-app-icon" data-home-action="documents"><span class="ios-icon">📁</span><b>Documents</b></button>
    <button class="ios-app-icon" data-home-action="library"><span class="ios-icon">📚</span><b>Library</b></button>
    <button class="ios-app-icon" data-home-action="projects"><span class="ios-icon">🗂️</span><b>Projects</b></button>
    <button class="ios-app-icon" data-home-action="scheduled"><span class="ios-icon">🕘</span><b>Scheduled</b></button>
    <button class="ios-app-icon" data-home-action="plugins"><span class="ios-icon">🧩</span><b>Plugins</b></button>
    <button class="ios-app-icon" data-home-action="codex"><span class="ios-icon">⌘</span><b>Codex</b></button>
  </div>

  <div class="ios-widget-row">
    <button class="ios-widget ios-doc-widget" data-home-action="documents">
      <span class="ios-widget-title">Documents</span>
      <strong id="iosDocCount">0 files</strong>
      <small>Browse your workspace</small>
    </button>
    <button class="ios-widget ios-ai-widget" data-home-action="chat">
      <span class="ios-widget-title">ASRIT</span>
      <strong>Ask anything</strong>
      <small>Chat, analyze, create</small>
    </button>
  </div>

  <div class="ios-home-dock">
    <button data-home-action="chat">💬</button>
    <button data-home-action="documents">📁</button>
    <button data-home-action="images">✨</button>
    <button data-home-action="library">📚</button>
  </div>
</section>

<section id="iosDocumentsScreen" class="ios-docs-screen" aria-label="Documents">
  <div class="ios-docs-header">
    <button id="iosDocsBack" type="button">‹ Home</button>
    <div>
      <strong>Documents</strong>
      <small>ASRIT workspace</small>
    </div>
    <button id="iosDocsRefresh" type="button">↻</button>
  </div>
  <div class="ios-docs-toolbar">
    <input id="iosDocSearch" type="search" placeholder="Search documents">
    <button id="iosDocUpload" type="button">＋ Add</button>
  </div>
  <div id="iosDocsList" class="ios-docs-list">
    <div class="ios-empty-docs">No documents yet.</div>
  </div>
</section>


<script>
(function(){
  const home = document.getElementById('iosHomeScreen');
  const docs = document.getElementById('iosDocumentsScreen');
  if(!home || !docs) return;

  function showHome(){
    docs.classList.remove('is-visible');
    home.classList.add('is-visible');
  }
  function showDocs(){
    home.classList.remove('is-visible');
    docs.classList.add('is-visible');
    loadDocs();
  }

  document.querySelectorAll('[data-home-action]').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const action=btn.dataset.homeAction;
      if(action==='documents'){showDocs();return;}
      home.classList.remove('is-visible');
      // Use existing ASRIT navigation where available.
      const nav=document.querySelector(`[data-section="${action}"], [data-page="${action}"], [onclick*="${action}"]`);
      if(nav) nav.click();
      else if(typeof window.showSection==='function') window.showSection(action);
    });
  });
  document.getElementById('iosDocsBack')?.addEventListener('click',showHome);
  document.getElementById('iosDocsRefresh')?.addEventListener('click',loadDocs);
  document.getElementById('iosDocSearch')?.addEventListener('input',loadDocs);

  async function loadDocs(){
    const list=document.getElementById('iosDocsList');
    const q=(document.getElementById('iosDocSearch')?.value||'').toLowerCase();
    try{
      const r=await fetch('/api/library');
      if(!r.ok) throw new Error('library');
      const data=await r.json();
      const items=Array.isArray(data)?data:(data.items||data.files||[]);
      const filtered=items.filter(x=>JSON.stringify(x).toLowerCase().includes(q));
      document.getElementById('iosDocCount').textContent=`${items.length} file${items.length===1?'':'s'}`;
      if(!filtered.length){list.innerHTML='<div class="ios-empty-docs">No documents found.</div>';return;}
      list.innerHTML=filtered.map(x=>{
        const name=x.name||x.filename||x.title||'Document';
        const size=x.size?`${Math.round(x.size/1024)} KB`:'Workspace file';
        return `<div class="ios-doc-card"><span class="ios-doc-card-icon">📄</span><div><div class="ios-doc-card-name">${escapeHtml(name)}</div><div class="ios-doc-card-meta">${escapeHtml(size)}</div></div></div>`;
      }).join('');
    }catch(e){
      list.innerHTML='<div class="ios-empty-docs">Documents are ready when files are added to your ASRIT library.</div>';
    }
  }
  function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}

  // Home is the first screen on initial load, but the existing app remains underneath.
  window.openASRITHome=showHome;
  window.openASRITDocuments=showDocs;
})();
</script>

</body>

</html>

/* =========================================================
   ASRIT 2026 UI REFINEMENT — FINAL OVERRIDES
   These rules intentionally sit last so they win over the older
   glass/legacy styles above them.
========================================================= */
:root{
  --asrit-ink:#172534;
  --asrit-muted:#687889;
  --asrit-blue:#42bdf5;
  --asrit-blue-2:#8ccfff;
  --asrit-card:rgba(255,255,255,.72);
  --asrit-line:rgba(45,76,98,.12);
}

/* Sidebar: cleaner hierarchy */
.sidebar{
  width:285px!important;
  padding:20px 18px!important;
  gap:13px!important;
  background:rgba(255,248,242,.72)!important;
  border-right:1px solid rgba(71,91,104,.12)!important;
}
.logo{padding:4px 7px 10px!important;gap:12px!important}
.logo-orb{width:42px!important;height:42px!important;border-radius:14px!important}
.logo-title{color:#152538!important;font-size:20px!important;letter-spacing:.07em!important}
.logo-sub{color:#718294!important}

/* New conversation — proper pill/button, no clipped baseline */
.new-chat.glass-button{
  width:100%!important;
  min-height:46px!important;
  height:46px!important;
  display:flex!important;
  align-items:center!important;
  justify-content:flex-start!important;
  gap:10px!important;
  padding:0 14px!important;
  border:1px solid rgba(75,105,124,.14)!important;
  border-radius:15px!important;
  background:rgba(255,255,255,.82)!important;
  color:#18293a!important;
  box-shadow:0 5px 18px rgba(65,91,105,.06)!important;
  font-size:14px!important;
  line-height:1!important;
}
.new-chat:hover{background:#fff!important;transform:translateY(-1px)!important}
.new-chat-plus{
  width:25px;height:25px;border-radius:8px;display:grid;place-items:center;
  flex:0 0 25px;background:linear-gradient(135deg,#73cfff,#45b8ef);
  color:#082235!important;font-size:19px;font-weight:700;line-height:1;
}

/* Sidebar search — one clean surface, no inner rectangle */
.sidebar .search{
  width:100%!important;height:43px!important;min-height:43px!important;
  padding:0 15px 0 16px!important;margin:0 0 9px!important;
  border:1px solid rgba(72,95,112,.14)!important;
  outline:none!important;box-shadow:0 5px 18px rgba(65,91,105,.045)!important;
  border-radius:14px!important;background:rgba(255,255,255,.82)!important;
  color:#172534!important;font-size:13px!important;
}
.sidebar .search:focus{border-color:rgba(65,178,233,.42)!important;box-shadow:0 0 0 3px rgba(65,178,233,.10)!important}
.sidebar .search::placeholder{color:#80909e!important}

/* Sidebar nav */
.side-nav{gap:7px!important;margin:0!important}
.nav-item{
  min-height:47px!important;padding:0 13px!important;border-radius:15px!important;
  border:1px solid transparent!important;background:transparent!important;
  color:#263645!important;font-size:14px!important;
}
.nav-item:hover{background:rgba(255,255,255,.64)!important;border-color:rgba(75,105,124,.10)!important;box-shadow:none!important}
.nav-item.active{background:rgba(255,255,255,.78)!important;border-color:rgba(75,105,124,.12)!important;box-shadow:0 6px 18px rgba(65,91,105,.055)!important}
.nav-icon{width:25px!important;font-size:16px!important;color:#263645!important}

/* Top menu button — unmistakable */
.topbar .icon-button[title="Menu"]{
  width:48px!important;height:48px!important;border-radius:15px!important;
  display:grid!important;place-items:center!important;
  background:rgba(255,255,255,.88)!important;
  border:1px solid rgba(70,93,108,.13)!important;
  color:#172534!important;box-shadow:0 7px 20px rgba(55,82,98,.07)!important;
}
.menu-bars{width:20px;height:16px;display:flex;flex-direction:column;justify-content:space-between}
.menu-bars i{display:block;width:20px;height:2px;border-radius:4px;background:#172534}

/* Bottom composer: no nested rectangle around the textarea */
.input-area{padding:12px max(5vw,40px) 22px!important}
.input-glass{
  width:100%!important;min-height:64px!important;padding:8px!important;gap:8px!important;
  align-items:center!important;border:1px solid rgba(71,91,104,.15)!important;
  border-radius:21px!important;background:rgba(255,249,244,.88)!important;
  box-shadow:0 12px 34px rgba(72,56,45,.10)!important;
}
#messageInput{
  min-width:0!important;height:46px!important;min-height:46px!important;
  margin:0!important;padding:12px 10px!important;
  border:0!important;border-width:0!important;outline:none!important;
  box-shadow:none!important;border-radius:12px!important;
  background:transparent!important;color:#172534!important;
}
#messageInput:focus{border:0!important;outline:none!important;box-shadow:none!important;transform:none!important}
#messageInput::placeholder{color:#7b8792!important}
.input-glass .icon-button{
  width:46px!important;height:46px!important;min-width:46px!important;
  flex:0 0 46px!important;border-radius:14px!important;
  display:grid!important;place-items:center!important;
  border:1px solid rgba(67,91,107,.13)!important;
  background:rgba(255,255,255,.94)!important;color:#183042!important;
  font-size:0!important;box-shadow:0 4px 12px rgba(55,80,95,.055)!important;
}
.input-glass .icon-button:hover{background:#fff!important;transform:translateY(-1px)!important}
.plus-mark{font-size:25px!important;font-weight:400!important;line-height:1!important;color:#1a3447!important}
.mic-mark{
  position:relative;width:14px;height:19px;border:2px solid #1a3447;border-radius:8px;
  display:block!important;font-size:0!important;
}
.mic-mark::after{content:"";position:absolute;left:50%;bottom:-7px;width:16px;height:8px;
  transform:translateX(-50%);border:2px solid #1a3447;border-top:0;border-radius:0 0 10px 10px}
.mic-mark::before{content:"";position:absolute;left:50%;bottom:-10px;width:2px;height:5px;
  transform:translateX(-50%);background:#1a3447;border-radius:2px}
.send-button{background:linear-gradient(135deg,#7dcdf5,#46b7ee)!important;border:0!important}
.send-mark{font-size:20px!important;font-weight:800!important;color:#092337!important}

/* Make the voice control impossible to hide on desktop */
#micButton{display:grid!important;visibility:visible!important;opacity:1!important}

/* Live button */
.topbar .glass-button[onclick="openCamera()"]{
  min-height:44px!important;padding:0 17px!important;border-radius:15px!important;
  background:rgba(255,255,255,.84)!important;border:1px solid rgba(70,93,108,.13)!important;
  color:#172534!important;box-shadow:0 6px 18px rgba(55,80,95,.055)!important;
}

/* Codex modern workspace */
.codex-grid{gap:14px!important;margin:18px 0!important;grid-template-columns:minmax(0,1.25fr) minmax(280px,.75fr)!important}
.codex-grid textarea{
  min-height:300px!important;border:1px solid rgba(70,95,112,.13)!important;
  border-radius:18px!important;background:rgba(255,255,255,.82)!important;
  padding:16px!important;color:#172534!important;box-shadow:none!important;
}
.workspace-panel{border-radius:28px!important;background:rgba(255,249,244,.78)!important;box-shadow:0 20px 60px rgba(67,55,45,.09)!important}
.codex-result{border-radius:18px!important;background:rgba(255,255,255,.72)!important}

@media(max-width:800px){
  .sidebar{width:min(88vw,340px)!important;padding:16px!important}
  .input-area{padding:8px 10px calc(82px + env(safe-area-inset-bottom))!important}
  .input-glass{min-height:60px!important;padding:7px!important;border-radius:20px!important}
  .input-glass .icon-button{width:43px!important;height:43px!important;min-width:43px!important;flex-basis:43px!important}
  #messageInput{height:43px!important;min-height:43px!important;padding:11px 7px!important}
  .codex-grid{grid-template-columns:1fr!important}
}


/* =========================================================
   RESPONSE ACTION BAR — modern AI reply controls
========================================================= */
.message.assistant{
  position:relative!important;
  flex-direction:column!important;
  align-items:flex-start!important;
  gap:8px!important;
}
.message.assistant .message-bubble{
  margin:0!important;
}
.response-actions{
  display:flex!important;
  align-items:center!important;
  gap:2px!important;
  margin-left:8px!important;
  min-height:30px!important;
  opacity:.72;
  transition:opacity .18s ease, transform .18s ease;
}
.message.assistant:hover .response-actions,
.response-actions:focus-within{
  opacity:1;
}
.response-action{
  width:31px!important;
  height:31px!important;
  min-width:31px!important;
  padding:0!important;
  border:0!important;
  border-radius:9px!important;
  display:grid!important;
  place-items:center!important;
  background:transparent!important;
  color:#566574!important;
  cursor:pointer!important;
  box-shadow:none!important;
  font:inherit!important;
  transition:background .16s ease,color .16s ease,transform .16s ease;
}
.response-action:hover{
  background:rgba(75,98,115,.10)!important;
  color:#172534!important;
  transform:translateY(-1px)!important;
}
.response-action.selected{
  background:rgba(67,188,241,.15)!important;
  color:#168fc9!important;
}
.ra-icon{
  display:block!important;
  line-height:1!important;
  font-size:17px!important;
  font-weight:500!important;
}
.response-action[data-action="like"] .ra-icon{
  transform:rotate(180deg);
}
.response-action[data-action="dislike"] .ra-icon{
  transform:none;
}
.response-action[data-action="copy"] .ra-icon{
  font-size:16px!important;
}
.response-action[data-action="more"] .ra-icon{
  font-size:12px!important;
  letter-spacing:1px!important;
  font-weight:800!important;
}
.ra-label{
  font-size:10px!important;
  font-weight:700!important;
  white-space:nowrap!important;
}
.message.assistant.retrying .message-bubble{
  opacity:.55;
}
.response-more-menu{
  position:absolute!important;
  left:8px!important;
  top:100%!important;
  z-index:1000!important;
  display:flex!important;
  flex-direction:column!important;
  min-width:160px!important;
  padding:6px!important;
  border:1px solid rgba(60,82,98,.13)!important;
  border-radius:13px!important;
  background:rgba(255,251,247,.98)!important;
  box-shadow:0 14px 35px rgba(52,63,70,.16)!important;
  backdrop-filter:blur(18px)!important;
  transform:translateY(-3px)!important;
  opacity:0!important;
  transition:opacity .15s ease,transform .15s ease;
}
.response-more-menu.open{
  opacity:1!important;
  transform:translateY(0)!important;
}
.response-more-menu button{
  border:0!important;
  background:transparent!important;
  color:#263746!important;
  text-align:left!important;
  border-radius:9px!important;
  padding:9px 11px!important;
  cursor:pointer!important;
  font-size:12px!important;
}
.response-more-menu button:hover{
  background:rgba(70,100,120,.08)!important;
}

/* Make the actual icons unmistakable, even if an older stylesheet is present. */
.new-chat-plus{
  font-family:Arial,sans-serif!important;
  font-size:18px!important;
  font-weight:600!important;
}
.topbar .icon-button[title="Menu"]{
  overflow:visible!important;
}
.menu-bars{
  width:19px!important;height:16px!important;
  display:flex!important;flex-direction:column!important;
  justify-content:space-between!important;align-items:stretch!important;
}
.menu-bars i{
  width:19px!important;height:2px!important;
  min-height:2px!important;display:block!important;
  background:#172534!important;border-radius:5px!important;
}
.input-glass .add-button,
.input-glass .mic-button{
  position:relative!important;
  overflow:visible!important;
}
.input-glass .add-button .plus-mark{
  display:block!important;
  font-family:Arial,sans-serif!important;
  font-size:25px!important;
  font-weight:400!important;
  line-height:1!important;
  color:#172f42!important;
}
.input-glass .mic-button .mic-mark{
  display:block!important;
  width:13px!important;height:18px!important;
  border:2px solid #172f42!important;
  border-radius:8px!important;
}
.input-glass .mic-button .mic-mark::after{
  content:""!important;
  position:absolute!important;
  left:50%!important;
  bottom:-7px!important;
  width:15px!important;height:8px!important;
  transform:translateX(-50%)!important;
  border:2px solid #172f42!important;
  border-top:0!important;
  border-radius:0 0 9px 9px!important;
}
.input-glass .mic-button .mic-mark::before{
  content:""!important;
  position:absolute!important;
  left:50%!important;
  bottom:-10px!important;
  width:2px!important;height:5px!important;
  transform:translateX(-50%)!important;
  background:#172f42!important;
  border-radius:2px!important;
}
@media(max-width:700px){
  .response-actions{margin-left:3px!important;opacity:.9!important}
  .response-action{width:30px!important;height:30px!important;min-width:30px!important}
}

/* =========================================================
   ASRIT FINAL UI PATCH — force visible controls
   This block intentionally comes LAST so legacy CSS cannot hide
   the composer controls or response action bar.
========================================================= */

/* Clean composer: one surface, no nested input rectangle */
.input-area{position:relative!important;z-index:50!important}
.input-glass{
  display:flex!important;align-items:center!important;gap:8px!important;
  width:100%!important;min-height:64px!important;padding:8px!important;
  border:1px solid rgba(50,65,78,.14)!important;border-radius:20px!important;
  background:rgba(255,250,246,.96)!important;
  box-shadow:0 10px 30px rgba(42,55,65,.12)!important;
}
.input-glass #messageInput{
  display:block!important;visibility:visible!important;opacity:1!important;
  flex:1 1 auto!important;min-width:0!important;width:auto!important;
  height:46px!important;min-height:46px!important;margin:0!important;
  padding:12px 8px!important;border:0!important;outline:0!important;
  background:transparent!important;box-shadow:none!important;border-radius:10px!important;
  color:#172534!important;font-size:15px!important;
}
.input-glass #messageInput::placeholder{color:#7c6b63!important;opacity:1!important}

/* +, mic and send are real fixed-size controls */
.input-glass .add-button,
.input-glass #micButton,
.input-glass .send-button{
  display:grid!important;visibility:visible!important;opacity:1!important;
  flex:0 0 46px!important;width:46px!important;height:46px!important;
  min-width:46px!important;min-height:46px!important;
  place-items:center!important;border-radius:14px!important;
  border:1px solid rgba(50,65,78,.12)!important;
  background:#fff!important;color:#172534!important;
  box-shadow:0 4px 12px rgba(42,55,65,.08)!important;
  position:relative!important;overflow:visible!important;
}
.input-glass .add-button{order:1!important}
.input-glass #messageInput{order:2!important}
.input-glass #micButton{order:3!important}
.input-glass .send-button{order:4!important}
.input-glass .add-button:hover,.input-glass #micButton:hover{background:#f5fbff!important}
.input-glass .send-button{background:linear-gradient(135deg,#83d5f8,#4bb9ef)!important;border:0!important}

.plus-mark{display:block!important;font-family:Arial,sans-serif!important;font-size:27px!important;font-weight:400!important;line-height:1!important;color:#172f42!important}
.mic-mark{display:block!important;position:relative!important;width:13px!important;height:18px!important;border:2px solid #172f42!important;border-radius:8px!important;font-size:0!important}
.mic-mark:after{content:""!important;position:absolute!important;left:50%!important;bottom:-7px!important;width:15px!important;height:8px!important;transform:translateX(-50%)!important;border:2px solid #172f42!important;border-top:0!important;border-radius:0 0 9px 9px!important}
.mic-mark:before{content:""!important;position:absolute!important;left:50%!important;bottom:-10px!important;width:2px!important;height:5px!important;transform:translateX(-50%)!important;background:#172f42!important;border-radius:2px!important}
.send-mark{font-size:21px!important;font-weight:800!important;color:#092337!important}

/* Top-left menu: three unmistakable bars */
.topbar .icon-button[title="Menu"]{
  display:grid!important;visibility:visible!important;opacity:1!important;
  width:48px!important;height:48px!important;place-items:center!important;
  border-radius:15px!important;background:#fff!important;
  border:1px solid rgba(50,65,78,.12)!important;color:#172534!important;
  box-shadow:0 6px 18px rgba(42,55,65,.08)!important;
}
.menu-bars{display:flex!important;width:20px!important;height:16px!important;flex-direction:column!important;justify-content:space-between!important}
.menu-bars i{display:block!important;width:20px!important;height:2px!important;min-height:2px!important;background:#172534!important;border-radius:5px!important}

/* New conversation */
.new-chat{
  display:flex!important;align-items:center!important;justify-content:flex-start!important;gap:9px!important;
  width:100%!important;min-height:44px!important;padding:0 14px!important;
  border-radius:14px!important;border:1px solid rgba(50,65,78,.12)!important;
  background:#fff!important;color:#243542!important;font-size:13px!important;font-weight:600!important;
  box-shadow:0 5px 15px rgba(42,55,65,.06)!important;
}
.new-chat:hover{background:#f9fcfe!important;transform:translateY(-1px)!important}
.new-chat-plus{display:grid!important;place-items:center!important;width:24px!important;height:24px!important;border-radius:8px!important;background:#55bfee!important;color:#092337!important;font-size:20px!important;line-height:1!important;font-weight:500!important}

/* Response action bar — always visible, matching the supplied reference */
.message.assistant{display:flex!important;flex-direction:column!important;align-items:flex-start!important;gap:5px!important;position:relative!important}
.response-actions{
  display:flex!important;visibility:visible!important;opacity:1!important;
  align-items:center!important;gap:1px!important;margin:0 0 4px 6px!important;
  min-height:32px!important;padding:0!important;
}
.response-action{
  display:grid!important;visibility:visible!important;opacity:1!important;place-items:center!important;
  width:32px!important;height:32px!important;min-width:32px!important;padding:0!important;
  border:0!important;border-radius:8px!important;background:transparent!important;
  color:#596671!important;cursor:pointer!important;font-family:Arial,sans-serif!important;
}
.response-action:hover{background:rgba(70,90,105,.09)!important;color:#182733!important}
.response-action .ra-icon{display:block!important;font-size:17px!important;line-height:1!important;font-weight:400!important}
.response-action[data-action="more"] .ra-icon{font-size:13px!important;letter-spacing:1px!important;font-weight:700!important}
.response-action[data-action="copy"] .ra-icon{font-size:18px!important}
.response-action[data-action="retry"] .ra-icon{font-size:19px!important}

/* Make response action bar visible on touch devices too */
@media (hover:none){.response-actions{opacity:1!important}}

/* Mobile: preserve all three controls */
@media(max-width:700px){
  .input-glass{min-height:58px!important;padding:6px!important;gap:5px!important}
  .input-glass .add-button,.input-glass #micButton,.input-glass .send-button{width:42px!important;height:42px!important;min-width:42px!important;flex-basis:42px!important}
  .input-glass #messageInput{height:42px!important;min-height:42px!important;padding:10px 5px!important}
}


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
