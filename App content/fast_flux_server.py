import os
import io
import json
import base64
import uvicorn
import tempfile
import shutil
import requests
import subprocess
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import Response
from pydantic import BaseModel
from mflux.models.flux.variants.txt2img.flux import Flux1
from mflux.models.common.config.model_config import ModelConfig
from mflux.models.common.config.config import Config
import mlx_whisper
import torch
from transformers import VitsModel, AutoTokenizer
import scipy.io.wavfile
from PIL import Image
import fitz  # PyMuPDF
import re
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def resize_and_pad_image(image_bytes, max_dim=1024, ratio=None, center=True):
    """
    Resizes image to fit max_dim. 
    If ratio is provided (e.g. (16, 9)), pads the image to that aspect ratio.
    Returns: (processed_bytes, original_w, original_h, canvas_w, canvas_h, scale_factor, offset_x, offset_y)
    """
    img = Image.open(io.BytesIO(image_bytes))
    orig_w, orig_h = img.size
    
    if ratio:
        target_ratio = ratio[0] / ratio[1]
        if target_ratio > 1: # Landscape
            canvas_w = max_dim
            canvas_h = int(max_dim / target_ratio)
        else: # Portrait
            canvas_h = max_dim
            canvas_w = int(max_dim * target_ratio)
            
        scale = min(canvas_w / orig_w, canvas_h / orig_h)
        new_w, new_h = int(orig_w * scale), int(orig_h * scale)
        img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        
        offset_x = (canvas_w - new_w) // 2 if center else 0
        offset_y = (canvas_h - new_h) // 2 if center else 0
        
        canvas = Image.new("RGB", (canvas_w, canvas_h), (0, 0, 0))
        canvas.paste(img_resized, (offset_x, offset_y))
        
        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        return buf.getvalue(), orig_w, orig_h, canvas_w, canvas_h, scale, offset_x, offset_y
    else:
        # Standard resize maintaining aspect ratio (no padding)
        if max(orig_w, orig_h) > max_dim:
            print(f"Resizing image from {orig_w}x{orig_h} to max {max_dim}...")
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            new_w, new_h = img.size
            scale = new_w / orig_w
            return buf.getvalue(), orig_w, orig_h, new_w, new_h, scale, 0, 0
        return image_bytes, orig_w, orig_h, orig_w, orig_h, 1.0, 0, 0

def clean_html_noise(html):
    """
    Strips out scripts, styles, and bulky attributes to reduce noise for the LLM.
    """
    # 1. Remove scripts and styles
    html = re.sub(r'<script.*?>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<style.*?>.*?</style>', '', html, flags=re.DOTALL | re.IGNORECASE)
    # 2. Remove common bulky attributes that don't help extraction
    html = re.sub(r' (class|style|srcset|data-[^=]*|aria-[^=]*|role|tabindex|decoding|loading|fetchpriority|width|height|width|height|alt|title)="[^"]*"', '', html, flags=re.IGNORECASE)
    # 3. Collapse multiple spaces and newlines
    html = re.sub(r'\n\s*\n', '\n', html)
    html = re.sub(r' +', ' ', html)
    return html.strip()

app = FastAPI(title="Flux & Transcription AI API")

# Configuration
OLLAMA_BASE_URL = "http://localhost:11434"

# Initialize models (Lazy loading)
flux_model = None

def get_flux_model():
    global flux_model
    if flux_model is None:
        print("Loading Flux.1-schnell model...")
        flux_model = Flux1(
            model_config=ModelConfig.schnell(),
            quantize=4
        )
        print("Flux model loaded.")
    return flux_model

# TTS Model loading (Lazy loading)
tts_model = None
tts_tokenizer = None
tts_model_name = "facebook/mms-tts-deu"

def get_tts_model():
    global tts_model, tts_tokenizer
    if tts_model is None:
        print(f"Loading TTS model {tts_model_name}...")
        device = "mps" if torch.backends.mps.is_available() else "cpu"
        tts_tokenizer = AutoTokenizer.from_pretrained(tts_model_name)
        tts_model = VitsModel.from_pretrained(tts_model_name).to(device)
        print("TTS model loaded.")
    return tts_model, tts_tokenizer

# Models for Request Bodies
class GenerateRequest(BaseModel):
    prompt: str
    seed: int = 42
    steps: int = 4
    width: int = 512
    height: int = 512

class StoryRequest(BaseModel):
    words: list[str]

class TTSRequest(BaseModel):
    text: str

class InvoiceRequest(BaseModel):
    prompt: str | None = None

class PromptRequest(BaseModel):
    word: str
    translation: str

class ProductRequest(BaseModel):
    content: str | None = None
    url: str | None = None
    myproduct: str | None = None
    prompt: str | None = None

def extract_json(text):
    """
    Robustly parse JSON from model output.
    """
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    for fence in ["```json", "```"]:
        if fence in text:
            try:
                inner = text.split(fence)[1].split("```")[0].strip()
                return json.loads(inner)
            except Exception:
                pass
    # Search for anything resembling a list or dict
    match = re.search(r'(\[.*\]|\{.*\})', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    
    # Last resort: search for 4 numbers in a list or 2 pairs of coordinates
    # Pattern 1: [n, n, n, n]
    list_match = re.search(r'\[\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\]', text)
    if list_match:
        return {"bbox": [int(x) for x in list_match.groups()]}
    
    # Pattern 2: (x1, y1) and (x2, y2)
    pairs = re.findall(r'\(\s*(\d+)\s*,\s*(\d+)\s*\)', text)
    if len(pairs) >= 2:
        return {"bbox": [int(pairs[0][0]), int(pairs[0][1]), int(pairs[1][0]), int(pairs[1][1])]}

    raise ValueError(f"Cannot parse JSON or coordinates from: {text}")

@app.post("/generate")
async def generate_image(request: GenerateRequest):
    try:
        print(f"Generating image for prompt: {request.prompt}")
        model = get_flux_model()
        output = model.generate_image(
            seed=request.seed,
            prompt=request.prompt,
            width=request.width,
            height=request.height,
            num_inference_steps=request.steps
        )
        pil_image = output.image
        img_byte_arr = io.BytesIO()
        pil_image.save(img_byte_arr, format='PNG')
        return Response(content=img_byte_arr.getvalue(), media_type="image/png")
    except Exception as e:
        print(f"Error generating image: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/transcribe")
async def transcribe_audio(file: UploadFile = File(...)):
    temp_path = None
    try:
        _, suffix = os.path.splitext(file.filename)
        suffix = suffix.lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(file.file, tmp)
            temp_path = tmp.name
        result = mlx_whisper.transcribe(temp_path, path_or_hf_repo="mlx-community/whisper-base-mlx")
        return {"text": result["text"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Transcription error: {str(e)}")
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/story")
async def generate_story(request: StoryRequest):
    try:
        words_str = ", ".join(request.words)
        prompt = f"Schreibe eine kurze, kreative Geschichte auf Deutsch, die die folgenden Wörter verwendet: {words_str}."
        ollama_payload = {"model": "llama3", "prompt": prompt, "stream": False}
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        return {"story": response.json()["response"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Story generation error: {str(e)}")

@app.post("/tts")
async def text_to_speech(request: TTSRequest):
    temp_wav = None
    temp_mp3 = None
    try:
        model, tokenizer = get_tts_model()
        device = model.device
        inputs = tokenizer(request.text, return_tensors="pt").to(device)
        with torch.no_grad():
            output = model(**inputs).waveform
        output_np = output.cpu().numpy().squeeze()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp_wav:
            temp_wav = tmp_wav.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_mp3:
            temp_mp3 = tmp_mp3.name
        scipy.io.wavfile.write(temp_wav, rate=model.config.sampling_rate, data=output_np)
        subprocess.run(["ffmpeg", "-y", "-i", temp_wav, "-codec:a", "libmp3lame", "-qscale:a", "2", temp_mp3], check=True, capture_output=True)
        with open(temp_mp3, "rb") as f:
            return Response(content=f.read(), media_type="audio/mpeg")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"TTS error: {str(e)}")
    finally:
        for p in [temp_wav, temp_mp3]:
            if p and os.path.exists(p): os.remove(p)

@app.post("/ocr")
async def extract_text(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        contents, _, _, _, _, _, _, _ = resize_and_pad_image(contents)
        image_base64 = base64.b64encode(contents).decode('utf-8')
        ollama_payload = {
            "model": "llava:7b-v1.6-mistral-q4_0",
            "prompt": "Read all the text in this image line by line.",
            "images": [image_base64],
            "stream": False
        }
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        return {"text": [l.strip() for l in response.json()["response"].split('\n') if l.strip()]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OCR error: {str(e)}")

@app.post("/invoice")
async def process_invoice(file: UploadFile = File(...), prompt: str | None = Form(None)):
    try:
        contents = await file.read()
        images_base64 = []
        
        if file.filename.lower().endswith('.pdf'):
            pdf_document = fitz.open(stream=contents, filetype="pdf")
            for page in pdf_document:
                # Convert page to high-quality PNG
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_bytes = pix.tobytes("png")
                # Resize and pad
                proc_bytes, _, _, _, _, _, _, _ = resize_and_pad_image(img_bytes)
                images_base64.append(base64.b64encode(proc_bytes).decode('utf-8'))
            pdf_document.close()
        else:
            # Handle standard image files
            proc_bytes, _, _, _, _, _, _, _ = resize_and_pad_image(contents)
            images_base64.append(base64.b64encode(proc_bytes).decode('utf-8'))

        if not prompt:
            prompt = (
                "Please extract all data from the attached invoice/transaction statement into a structured JSON format.\n\n"
                "Mandatory Requirements:\n"
                "1. Translation: If the invoice is in a language other than English (e.g., Korean, Japanese, Vietnamese), "
                "translate all extracted values—specifically product descriptions, company names, and storage locations—into English.\n"
                "2. Completeness: Every field listed in the JSON structure below is mandatory. If a specific piece of data is not found, "
                "populate it with an empty string (\"\").\n"
                "3. Format: Ensure all quantities and amounts are represented as numbers (remove commas and currency symbols). "
                "Dates should be in YYYY-MM-DD format.\n\n"
                "JSON Structure:\n"
                "{\n"
                "  \"vendor_name\": \"...\",\n"
                "  \"invoice_number\": \"...\",\n"
                "  \"invoice_date\": \"...\",\n"
                "  \"currency\": \"...\",\n"
                "  \"line_items\": [\n"
                "    {\n"
                "      \"contract_no\": \"...\",\n"
                "      \"product_model\": \"...\",\n"
                "      \"description\": \"...\",\n"
                "      \"quantity\": 0,\n"
                "      \"unit_price\": 0.0,\n"
                "      \"total_amount\": 0.0\n"
                "    }\n"
                "  ],\n"
                "  \"total_invoice_amount\": 0.0\n"
                "}\n\n"
                "Return ONLY a JSON object."
            )

        ollama_payload = {
            "model": "qwen2.5vl:7b",
            "prompt": prompt,
            "images": images_base64,
            "stream": False
        }
        
        print(f"Calling Qwen2.5-VL for invoice extraction ({len(images_base64)} pages)...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        return extract_json(response.json()["response"])
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Invoice error: {str(e)}")

@app.post("/detect")
async def detect_objects(file: UploadFile = File(...), prompt: str | None = Form(None)):
    try:
        contents = await file.read()
        contents, orig_w, orig_h, canvas_w, canvas_h, scale, offset_x, offset_y = resize_and_pad_image(contents, ratio=(16, 9))
        image_base64 = base64.b64encode(contents).decode('utf-8')
        if not prompt: prompt = "Detect all visible UI elements."
        p = f"{prompt}. Return JSON list: [{{'label': '...', 'bbox': [xmin, ymin, xmax, ymax]}}]."
        payload = {"model": "qwen2.5vl:7b", "prompt": p, "images": [image_base64], "stream": False}
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=payload)
        response.raise_for_status()
        raw = response.json()["response"]
        parsed = extract_json(raw)
        if isinstance(parsed, dict): 
            for v in parsed.values(): 
                if isinstance(v, list): parsed = v; break
        elements = []
        for item in parsed:
            label = item.get("label") or "unknown"
            bbox = item.get("bbox") or item.get("box") or item.get("coordinates")
            if bbox and len(bbox) == 4:
                xmin, ymin, xmax, ymax = bbox
                is_normalized = all(0 <= v <= 1000 for v in bbox)
                if is_normalized:
                    px = {
                        "xmin": int(((xmin * canvas_w / 1000) - offset_x) / scale),
                        "ymin": int(((ymin * canvas_h / 1000) - offset_y) / scale),
                        "xmax": int(((xmax * canvas_w / 1000) - offset_x) / scale),
                        "ymax": int(((ymax * canvas_h / 1000) - offset_y) / scale)
                    }
                else:
                    px = {"xmin": int((xmin-offset_x)/scale), "ymin": int((ymin-offset_y)/scale), "xmax": int((xmax-offset_x)/scale), "ymax": int((ymax-offset_y)/scale)}
                elements.append({"label": label, "bbox_pixels": px})
        return {"elements": elements, "image_size": [orig_w, orig_h]}
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[detect-precise] Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/detect-precise")
async def detect_precise(file: UploadFile = File(...), labels: str = Form(...)):
    """
    High-precision multi-stage detection pipeline for multiple labels.
    """
    try:
        label_list = [l.strip() for l in labels.split(",") if l.strip()]
        contents = await file.read()
        orig_img = Image.open(io.BytesIO(contents)).convert("RGB")
        orig_w, orig_h = orig_img.size

        # Stage 1: Bulk Coarse
        print(f"[detect-precise] Stage 1: Bulk Coarse search for {label_list}...")
        s1_bytes, _, _, canvas_w, canvas_h, scale, off_x, off_y = resize_and_pad_image(contents, max_dim=1536, ratio=None, center=False)
        s1_b64 = base64.b64encode(s1_bytes).decode("utf-8")
        coarse_prompt = (
            f"This image resolution is {canvas_w}x{canvas_h}. "
            f"Find the bounding boxes [xmin, ymin, xmax, ymax] in pixels for: {', '.join(label_list)}. "
            f"Return ONLY a JSON list: [{{\"label\": \"...\", \"bbox\": [xmin, ymin, xmax, ymax]}}, ...]"
        )
        payload1 = {"model": "qwen2.5vl:7b", "prompt": coarse_prompt, "images": [s1_b64], "stream": False}
        resp1 = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=payload1)
        resp1.raise_for_status()
        parsed1 = extract_json(resp1.json()["response"])
        if not isinstance(parsed1, list):
            if isinstance(parsed1, dict):
                for v in parsed1.values():
                    if isinstance(v, list): parsed1 = v; break
            if not isinstance(parsed1, list): parsed1 = [parsed1]

        final_results = []
        for item in parsed1:
            if not isinstance(item, dict): continue
            curr_label = item.get("label", "unknown")
            bbox = item.get("bbox") or item.get("bbox_2d") or item.get("box") or item.get("coordinates")
            if not bbox or len(bbox) != 4: continue
            
            n1, n2, n3, n4 = bbox
            cx1_n, cx2_n = min(n1, n3), max(n1, n3)
            cy1_n, cy2_n = min(n2, n4), max(n2, n4)
            
            is_normalized = all(v <= 1000 for v in [cx1_n, cy1_n, cx2_n, cy2_n])
            if max(canvas_w, canvas_h) > 1000 and max(cx2_n, cy2_n) > 800: is_normalized = False
            
            if is_normalized:
                cx1 = max(0, min(orig_w, int(((cx1_n * canvas_w / 1000) - off_x) / scale)))
                cy1 = max(0, min(orig_h, int(((cy1_n * canvas_h / 1000) - off_y) / scale)))
                cx2 = max(0, min(orig_w, int(((cx2_n * canvas_w / 1000) - off_x) / scale)))
                cy2 = max(0, min(orig_h, int(((cy2_n * canvas_h / 1000) - off_y) / scale)))
            else:
                cx1 = max(0, min(orig_w, int((cx1_n - off_x) / scale)))
                cy1 = max(0, min(orig_h, int((cy1_n - off_y) / scale)))
                cx2 = max(0, min(orig_w, int((cx2_n - off_x) / scale)))
                cy2 = max(0, min(orig_h, int((cy2_n - off_y) / scale)))

            # Stage 2: Individual Fine
            w_px, h_px = cx2 - cx1, cy2 - cy1
            mx, my = max(150, int(w_px * 1.5)), max(100, int(h_px * 1.5))
            c1x, c1y = max(0, cx1 - mx), max(0, cy1 - my)
            c2x, c2y = min(orig_w, cx2 + mx), min(orig_h, cy2 + my)
            
            crop = orig_img.crop((c1x, c1y, c2x, c2y))
            cw, ch = crop.size
            z_scale = 1024 / max(cw, ch)
            zw, zh = int(cw * z_scale), int(ch * z_scale)
            zoomed = crop.resize((zw, zh), Image.Resampling.LANCZOS)
            debug_crop_path = f"/Users/hung_macbook_pro/Ollama/debug_crop_{curr_label.replace(' ', '_')}.png"
            zoomed.save(debug_crop_path)
            print(f"[detect-precise] Debug: Saved zoomed crop to {debug_crop_path}")
            
            buf2 = io.BytesIO()
            zoomed.save(buf2, format="PNG")
            s2_b64 = base64.b64encode(buf2.getvalue()).decode("utf-8")
            
            f_prompt = (
                f"This zoomed image resolution is {zw}x{zh}. Find the exact bounding box [xmin, ymin, xmax, ymax] "
                f"in pixels of the element with text '{curr_label}'. "
                f"Return ONLY a JSON object: {{\"bbox\": [xmin, ymin, xmax, ymax]}}. "
                f"No explanations."
            )
            payload2 = {"model": "qwen2.5vl:7b", "prompt": f_prompt, "images": [s2_b64], "stream": False}
            resp2 = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=payload2)
            parsed2 = extract_json(resp2.json()["response"])
            f_bbox = None
            if isinstance(parsed2, dict):
                for k in ("bbox", "bbox_2d", "box", "coordinates"):
                    if isinstance(parsed2.get(k), list) and len(parsed2[k]) == 4: f_bbox = parsed2[k]; break
            if not f_bbox: continue

            f1, f2, f3, f4 = f_bbox
            fx1_n, fx2_n = min(f1, f3), max(f1, f3)
            fy1_n, fy2_n = min(f2, f4), max(f2, f4)
            # Stage 2: We explicitly ask for pixels, and 7B is good at it.
            # We assume pixels [0-zw, 0-zh]
            fx1_z, fy1_z, fx2_z, fy2_z = fx1_n, fy1_n, fx2_n, fy2_n

            final_xmin = int((fx1_z / z_scale) + c1x)
            final_ymin = int((fy1_z / z_scale) + c1y)
            final_xmax = int((fx2_z / z_scale) + c1x)
            final_ymax = int((fy2_z / z_scale) + c1y)

            final_results.append({
                "label": curr_label,
                "bbox_pixels": {"xmin": final_xmin, "ymin": final_ymin, "xmax": final_xmax, "ymax": final_ymax},
                "stage1_coarse_pixels": {"xmin": cx1, "ymin": cy1, "xmax": cx2, "ymax": cy2},
                "crop_region": {"x1": c1x, "y1": c1y, "x2": c2x, "y2": c2y}
            })
            
        return {"status": "success", "results": final_results, "image_size": [orig_w, orig_h]}
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[detect-precise] Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/locate-keywords")
async def locate_keywords(file: UploadFile = File(...), keywords: str = Form(...)):
    try:
        kw_list = [k.strip() for k in keywords.split(",") if k.strip()]
        contents = await file.read()
        contents, orig_w, orig_h, canvas_w, canvas_h, scale, off_x, off_y = resize_and_pad_image(contents, ratio=(1, 1))
        image_base64 = base64.b64encode(contents).decode('utf-8')
        p = f"Find bounding boxes for: {', '.join(kw_list)}. JSON list with 'keyword' and 'bbox' (normalized 0-1000)."
        payload = {"model": "qwen2.5vl:7b", "prompt": p, "images": [image_base64], "stream": False}
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=payload)
        detected_items = extract_json(response.json()["response"])
        final_results = []
        if isinstance(detected_items, list):
            for item in detected_items:
                kw = item.get("keyword") or item.get("label") or "unknown"
                bbox = item.get("bbox")
                if bbox and len(bbox) == 4:
                    xmin, ymin, xmax, ymax = bbox
                    px = {
                        "xmin": int(((xmin * canvas_w / 1000) - off_x) / scale),
                        "ymin": int(((ymin * canvas_h / 1000) - off_y) / scale),
                        "xmax": int(((xmax * canvas_w / 1000) - off_x) / scale),
                        "ymax": int(((ymax * canvas_h / 1000) - off_y) / scale)
                    }
                    final_results.append({"keyword": kw, "bbox_pixels": px})
        return {"results": final_results, "image_size": [orig_w, orig_h]}
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[detect-precise] Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/prompt")
async def create_prompt_text(request: PromptRequest):
    try:
        p = f"Describe a simple, clear, educational illustration for '{request.word}' ({request.translation})."
        payload = {"model": "llama3", "prompt": p, "stream": False}
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=payload)
        return {"prompt": response.json()["response"]}
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[detect-precise] Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/extract-product")
async def extract_product_info(request: ProductRequest):
    try:
        html_content = request.content
        
        # If URL is provided, fetch the content
        if request.url:
            print(f"Fetching HTML from URL: {request.url}")
            headers = {
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": "https://www.google.com/",
                "Upgrade-Insecure-Requests": "1"
            }
            resp = requests.get(request.url, headers=headers, timeout=15)
            resp.raise_for_status()
            html_content = resp.text
            
        if not html_content:
            return {"error": "No content or URL provided"}

        # Clean the HTML to remove noise (scripts, styles, tracking)
        cleaned_html = clean_html_noise(html_content)
        
        # Adaptive slicing on content
        my_prod = request.myproduct or "the product"
        
        import re
        # Handle multi-line names by allowing whitespace/newlines between words
        search_words = my_prod[:40].split()
        search_regex = r'\s+'.join([re.escape(w) for w in search_words])
        
        # 1. Find all name matches
        name_matches = [m.start() for m in re.finditer(search_regex, cleaned_html, re.IGNORECASE)]
        
        # 2. Find all currency/price matches
        # This regex looks for digits followed by common Vietnamese currency markers
        price_regex = r'(\d+[\.,])+\d+\s*[₫đ]|VNĐ|VND'
        currency_matches = [m.start() for m in re.finditer(price_regex, cleaned_html, re.IGNORECASE)]
        
        target_idx = -1
        if name_matches and currency_matches:
            # 3. Find the pair with smallest distance
            min_dist = float('inf')
            best_idx = name_matches[0]
            for n_idx in name_matches:
                for c_idx in currency_matches:
                    dist = abs(n_idx - c_idx)
                    if dist < min_dist:
                        min_dist = dist
                        best_idx = n_idx
            target_idx = best_idx
            print(f"Found product name near price marker (Distance: {min_dist}, Offset: {target_idx})")
        elif name_matches:
            target_idx = name_matches[0]
            print("Found product name but no price markers nearby.")
        elif currency_matches:
            target_idx = currency_matches[0]
            print("No product name found, anchoring on price marker.")
        else:
            # Fallback to identifiers
            for marker in ['price', 'giá', 'sku', 'thông tin']:
                target_idx = cleaned_html.lower().find(marker)
                if target_idx != -1: break
            if target_idx == -1: target_idx = 0
            print("No names or prices found, using fallback markers.")
        
        # Slice the content (center around target_idx)
        # We give 15k before and 65k after (total 80k)
        start_idx = max(0, target_idx - 15000)
        end_idx = min(len(cleaned_html), start_idx + 80000)
        content_to_send = cleaned_html[start_idx:end_idx]
        
        print(f"Extracting product info from sliced content (Range: {start_idx}-{end_idx}, Total: {len(cleaned_html)})")
        
        # Determine the prompt
        if request.prompt:
            extraction_prompt = request.prompt
        else:
            extraction_prompt = (
                "You are a precise data extraction engine. "
                f"Target Product to Extract: '{my_prod}'\n\n"
                "INSTRUCTIONS:\n"
                "1. Find the main product section in the provided HTML. Ignore 'Related Products' or 'Suggestions' sections.\n"
                "2. Extract the 'product_name' (this should be the actual product name found in the content that best matches the target product name). Prioritize official names from meta tags or main headings.\n"
                "3. Extract the 'price' for this specific target product. The price may be a formatted string (e.g., '18.690.000₫') or a numeric value (e.g., 18690000). Prioritize the main price shown for the target product. If the price is not found, return 'not_found'.\n"
                "4. Extract 'status' (e.g., 'Còn hàng').\n"
                "5. Calculate 'similarity' (match percentage of name found vs target name).\n"
                "6. Identify 'popup_xpath' for any blocking overlays (return empty string if none).\n\n"
                "Return ONLY a JSON object with this exact structure:\n"
                "{\n"
                "  \"product_name\": \"...\",\n"
                "  \"price\": \"...\",\n"
                "  \"status\": \"...\",\n"
                "  \"similarity\": \"...%\",\n"
                "  \"popup_xpath\": \"...\"\n"
                "}\n"
                "Do not include prices of similar or recommended products."
            )
        
        # Debug: Save what we're sending to the model
        with open("last_extraction_context.txt", "w") as f:
            f.write(f"PROMPT:\n{extraction_prompt}\n\nCONTENT:\n{content_to_send}")
        
        ollama_payload = {
            "model": "llama3",
            "prompt": f"HTML Content:\n{content_to_send}\n\n###\n\nINSTRUCTION:\n{extraction_prompt}",
            "stream": False,
            "format": "json"
        }
        
        print("Calling Ollama (Llama3) for product extraction...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        raw_res = result["response"]
        print(f"Raw Ollama Response: {raw_res}")
        
        try:
            extracted_data = json.loads(raw_res.strip())
            
            # Ensure all required keys exist to avoid missing values in UI
            expected_keys = ["product_name", "price", "status", "similarity", "popup_xpath"]
            for key in expected_keys:
                if key not in extracted_data:
                    extracted_data[key] = ""
            
            return extracted_data
        except Exception as parse_error:
            print(f"JSON Parse Error: {parse_error}. Attempting to clean response...")
            # Fallback: try to find JSON block if model returned conversational text
            match = re.search(r'\{.*\}', raw_res, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group())
                except: pass
            return {"error": "Failed to parse JSON", "raw": raw_res}
        
    except Exception as e:
        print(f"Error extracting product info: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Extraction error: {str(e)}")

@app.post("/extract-product-from-text")
async def extract_product_from_text(request: ProductRequest):
    try:
        content = request.content
        if not content:
            return {"error": "No content provided"}
        
        my_prod = request.myproduct or "the product"
        
        # Instruction for AI
        extraction_prompt = (
            "You are a precise data extraction engine for unstructured text.\n"
            f"Target Product to Extract: '{my_prod}'\n\n"
            "INSTRUCTIONS:\n"
            "1. Search the text from TOP to BOTTOM to find the best matching product name.\n"
            "2. Once the best match is found, continue searching from that point downwards to find the corresponding 'price' and 'status'.\n"
            "3. Extract 'product_name' (this must be the actual name found in the text that best matches the target product name).\n"
            "4. Extract 'price' (numeric value preferred). If the price is not found, return 'not_found'.\n"
            "5. Extract 'status' (e.g., 'Còn hàng').\n"
            "6. Calculate 'similarity' (match percentage of name found vs target name).\n"
            "7. Output 'popup_xpath' as an empty string (not applicable for text).\n\n"
            "Return ONLY a JSON object with this exact structure:\n"
            "{\n"
            "  \"product_name\": \"...\",\n"
            "  \"price\": \"...\",\n"
            "  \"status\": \"...\",\n"
            "  \"similarity\": \"...%\",\n"
            "  \"popup_xpath\": \"\"\n"
            "}"
        )
        
        ollama_payload = {
            "model": "llama3",
            "prompt": f"Text Content:\n{content}\n\n###\n\nINSTRUCTION:\n{extraction_prompt}",
            "stream": False,
            "format": "json"
        }
        
        print("Calling Ollama (Llama3) for plaintext product extraction...")
        # Since this is likely inside the same file, we use OLLAMA_BASE_URL
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        raw_res = result["response"]
        print(f"Raw Ollama Response: {raw_res}")
        
        try:
            extracted_data = json.loads(raw_res.strip())
            
            # Ensure all required keys exist
            expected_keys = ["product_name", "price", "status", "similarity", "popup_xpath"]
            for key in expected_keys:
                if key not in extracted_data:
                    extracted_data[key] = ""
            
            return extracted_data
        except Exception as parse_error:
            print(f"JSON Parse Error: {parse_error}. Attempting fallback extraction...")
            match = re.search(r'\{.*\}', raw_res, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group())
                except: pass
            return {"error": "Failed to parse JSON", "raw": raw_res}
            
    except Exception as e:
        print(f"Error in plaintext extraction: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Extraction error: {str(e)}")

@app.post("/extract-product-from-image")
async def extract_product_from_image(file: UploadFile = File(...), myproduct: str | None = Form(None), prompt: str | None = Form(None)):
    try:
        print(f"Extracting product from image: {file.filename}")
        
        # Read image content
        contents = await file.read()
        
        # Resize and Pad to 16:9 for consistent model performance
        contents, orig_w, orig_h, canvas_w, canvas_h, scale, off_x, off_y = resize_and_pad_image(contents, ratio=(16, 9))

        # Encode to base64 for Ollama
        image_base64 = base64.b64encode(contents).decode('utf-8')
        
        target_product = myproduct or "the main product in the image"
        
        # Determine the prompt
        if not prompt:
            prompt = (
                "You are a precise data extraction engine for product screenshots.\n"
                f"Target Product: '{target_product}'\n\n"
                "INSTRUCTIONS:\n"
                "1. Identify the main product in the center/main area of the image that matches the target product.\n"
                "2. Extract 'product_name' (the exact name displayed in the UI for the main product).\n"
                "3. Extract 'price' (the main price associated with this product). IMPORTANT: Price must be a monetary value (e.g., '25.290.000đ'). Ignore ratings (e.g., '4.9/5'), review counts, or technical specs. Ignore prices inside top banners, side advertisements, or 'Related Products' lists. If the main product is discontinued or has no price, return 'not_found'.\n"
                "4. Extract 'status' (e.g., 'In Stock', 'Còn hàng', 'Ngừng kinh doanh', return empty string if not found).\n"
                "5. Calculate 'similarity' (match percentage of name found vs target name).\n\n"
                "Return ONLY a JSON object with this exact structure:\n"
                "{\n"
                "  \"product_name\": \"...\",\n"
                "  \"price\": \"...\",\n"
                "  \"status\": \"...\",\n"
                "  \"similarity\": \"...%\"\n"
                "}"
            )
            
        ollama_payload = {
            "model": "qwen2.5vl:7b",
            "prompt": prompt,
            "images": [image_base64],
            "stream": False,
            "format": "json"
        }
        
        print(f"Calling Qwen2.5-VL for visual product extraction of '{target_product}'...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        raw_res = result["response"]
        print(f"Raw Ollama Response: {raw_res}")
        
        try:
            extracted_data = json.loads(raw_res.strip())
            
            # Ensure all required keys exist
            expected_keys = ["product_name", "price", "status", "similarity"]
            for key in expected_keys:
                if key not in extracted_data:
                    extracted_data[key] = ""
            
            return extracted_data
        except Exception as parse_error:
            print(f"JSON Parse Error: {parse_error}. Attempting fallback...")
            match = re.search(r'\{.*\}', raw_res, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group())
                except: pass
            return {"error": "Failed to parse JSON", "raw": raw_res}
            
    except Exception as e:
        print(f"Error in visual product extraction: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Extraction error: {str(e)}")


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)

