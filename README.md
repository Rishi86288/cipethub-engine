# 🎬 CIPETHUB Engine — Automated YouTube Video Generator

A Docker-based Python API service that **automatically generates professional educational YouTube lecture videos** for the **CIPETHUB** YouTube channel, targeting CIPET (Central Institute of Petrochemicals Engineering & Technology) students across India.

The service is triggered by **n8n** via HTTP. It accepts a topic, department, and video type, then:
1. Generates a structured lecture script with Google Gemini AI
2. Creates a professional PowerPoint presentation (python-pptx)
3. Converts slides to images (LibreOffice headless)
4. Generates Indian English voice narration (edge-tts — free, no API key)
5. Combines everything into an MP4 lecture video (ffmpeg)
6. Returns JSON metadata + a download URL

---

## Architecture

```
n8n (existing Docker container)
    → HTTP POST  cipethub-engine:8899/generate
        → Python generates:
            1. AI script via Google Gemini API
            2. Professional PPT  (python-pptx)
            3. Converts PPT slides to images  (LibreOffice headless)
            4. Indian English voiceover  (edge-tts)
            5. Slides + audio → MP4 video  (ffmpeg)
        → Returns JSON with video metadata + download URL
    → n8n downloads the video and uploads to YouTube
```

---

## Quick Start (Portainer Stack)

### Prerequisites
- Docker & Docker Compose installed on your server (ARM64 or AMD64)
- A Google Gemini API key (free tier available)
- Portainer running (optional but recommended)

### Step-by-Step via Portainer

1. **Log in** to Portainer → **Stacks** → **Add stack**
2. Set **Name**: `cipethub-engine`
3. In the **Web editor**, paste the contents of `docker-compose.yml`
4. Scroll down to **Environment variables** → click **Add environment variable**
   - Name: `GEMINI_API_KEY`  |  Value: `your_actual_gemini_key`
5. Click **Deploy the stack**
6. Wait 2-3 minutes for the image to build (first run downloads LibreOffice)
7. Verify with: `curl http://localhost:8899/health`

### Step-by-Step via CLI

```bash
git clone https://github.com/Rishi86288/cipethub-engine.git
cd cipethub-engine
cp .env.example .env
# Edit .env and set your GEMINI_API_KEY
nano .env
docker compose up -d --build
```

---

## API Endpoints

### `GET /health`
Returns service health status.

```bash
curl http://localhost:8899/health
```
```json
{"status": "ok", "service": "cipethub-engine", "port": 8899}
```

---

### `POST /generate`
Generates a complete lecture video. **Long-running** (up to 10 minutes).

```bash
curl -X POST http://localhost:8899/generate \
  -H "Content-Type: application/json" \
  -d '{
    "topic": "Injection Moulding Process",
    "department": "Plastics",
    "video_type": "PPT Lecture"
  }'
```

**Body parameters:**

| Field | Type | Values | Required |
|-------|------|--------|----------|
| `topic` | string | Any lecture topic | ✅ Yes |
| `department` | string | `Plastics`, `Mechanical`, `Manufacturing`, `All` | ✅ Yes |
| `video_type` | string | `PPT Lecture`, `Simulation` | ✅ Yes |

**Response:**
```json
{
  "status": "success",
  "title": "Injection Moulding Process | Plastics Engineering | CIPETHUB",
  "description": "Complete lecture on Injection Moulding... #CIPETHUB #CIPET",
  "tags": ["injection moulding", "CIPET", "Plastics", ...],
  "topic": "Injection Moulding Process",
  "department": "Plastics",
  "video_type": "PPT Lecture",
  "slides_count": 12,
  "duration_seconds": 420.5,
  "duration_minutes": 7.01,
  "size_mb": 45.2,
  "channel": "CIPETHUB",
  "video_path": "/videos/20240301_120000_plastics/lecture.mp4",
  "ppt_path": "/videos/20240301_120000_plastics/lecture.pptx"
}
```

---

### `GET /download`
Downloads the latest generated video as `cipethub_lecture.mp4`.

```bash
curl -O -J http://localhost:8899/download
```

---

### `GET /status`
Returns current generation status.

```bash
curl http://localhost:8899/status
```
```json
{
  "running": false,
  "last_result": { ... }
}
```

---

## Connecting to n8n

### Join the same Docker network

If n8n is running in a Docker container on the same host:

```bash
docker network connect <n8n_network_name> cipethub-engine
```

Then in n8n HTTP Request nodes, use the URL:
```
http://cipethub-engine:8899/generate
```

### n8n Workflow (8 Nodes)

| # | Node | Type | Purpose |
|---|------|------|---------|
| 1 | **Schedule** | Schedule Trigger | Runs daily at configured time |
| 2 | **Read Sheet** | Google Sheets | Reads pending topics from Google Sheet |
| 3 | **HTTP Generate** | HTTP Request | POST to `/generate` with topic/dept/type |
| 4 | **IF Success** | IF | Checks `status == "success"` |
| 5 | **HTTP Download** | HTTP Request | GET `/download` → saves video file |
| 6 | **Update Status** | Google Sheets | Writes title, description to sheet |
| 7 | **YouTube Upload** | YouTube | Uploads video with AI-generated metadata |
| 8 | **Mark Done** | Google Sheets | Sets Status column to "Done" |

---

## Google Sheet Setup

Create a Google Sheet with these columns:

| A: Date | B: Topic | C: Department | D: Video Type | E: Title | F: Description | G: Status |
|---------|----------|---------------|---------------|----------|----------------|-----------|
| 2024-03-01 | Injection Moulding | Plastics | PPT Lecture | | | Pending |

---

## Sample Topics

### Plastics Department
- Injection Moulding Process
- Extrusion Technology
- Blow Moulding Principles
- Thermoforming Process
- Polymer Properties and Testing

### Mechanical Department
- Heat Transfer Fundamentals
- Fluid Mechanics Basics
- Machine Design Principles
- Manufacturing Processes Overview
- CAD/CAM Integration

### Manufacturing Department
- Quality Control Methods
- CNC Machining Process
- Lean Manufacturing Principles
- Industrial Safety Standards
- Production Planning and Control

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `GEMINI_API_KEY` | *(required)* | Google Gemini API key |
| `TTS_VOICE_EN` | `en-IN-PrabhatNeural` | Indian English TTS voice |
| `TTS_VOICE_HI` | `hi-IN-MadhurNeural` | Hindi TTS voice (reserved) |
| `TZ` | `Asia/Kolkata` | Timezone |

---

## Troubleshooting

**Container not starting:**
```bash
docker logs cipethub-engine
```

**Video generation fails / Gemini error:**
- Check `GEMINI_API_KEY` is set correctly
- The fallback script will be used automatically if Gemini is unavailable

**LibreOffice conversion fails:**
```bash
docker exec cipethub-engine libreoffice --version
```

**Port already in use:**
- Change the host port in `docker-compose.yml`: `"127.0.0.1:8900:8899"`

**Check generation status:**
```bash
curl http://localhost:8899/status
```

**View generated videos:**
```bash
docker exec cipethub-engine ls /videos/
```

---

## License

MIT — Free to use for educational purposes. Go CIPET! 🇮🇳
