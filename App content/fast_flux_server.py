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
from diffusers import FluxPipeline
import whisper
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

def resize_and_pad_image(image_bytes, max_dim=1024, ratio=None):
    """
    Resizes image to fit max_dim. 
    If ratio is provided (e.g. (16, 9)), pads the image to that aspect ratio.
    Returns: (processed_bytes, original_w, original_h, canvas_w, canvas_h, scale_factor)
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
        
        canvas = Image.new("RGB", (canvas_w, canvas_h), (0, 0, 0))
        canvas.paste(img_resized, (0, 0))
        
        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        return buf.getvalue(), orig_w, orig_h, canvas_w, canvas_h, scale
    else:
        # Standard resize maintaining aspect ratio (no padding)
        if max(orig_w, orig_h) > max_dim:
            print(f"Resizing image from {orig_w}x{orig_h} to max {max_dim}...")
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            new_w, new_h = img.size
            scale = new_w / orig_w
            return buf.getvalue(), orig_w, orig_h, new_w, new_h, scale
        return image_bytes, orig_w, orig_h, orig_w, orig_h, 1.0

def clean_html_noise(html):
    """
    Strips out scripts, styles, and bulky attributes to reduce noise for the LLM.
    """
    # 1. Remove scripts and styles
    html = re.sub(r'<script.*?>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<style.*?>.*?</style>', '', html, flags=re.DOTALL | re.IGNORECASE)
    # 2. Remove common bulky attributes that don't help extraction
    # This removes class, style, srcset, data-*, aria-*, role, tabindex, etc.
    html = re.sub(r' (class|style|srcset|data-[^=]*|aria-[^=]*|role|tabindex|decoding|loading|fetchpriority|width|height|width|height|alt|title)="[^"]*"', '', html, flags=re.IGNORECASE)
    # 3. Collapse multiple spaces and newlines
    html = re.sub(r'\n\s*\n', '\n', html)
    html = re.sub(r' +', ' ', html)
    return html.strip()

app = FastAPI(title="Flux & Transcription AI API")

# Configuration
OLLAMA_BASE_URL = "http://localhost:11434"

# Initialize models
device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Loading Flux.1-schnell model on {device}...")
try:
    # Use bfloat16 for better performance if on CUDA, otherwise float32 for CPU
    flux_dtype = torch.bfloat16 if device == "cuda" else torch.float32
    flux_model = FluxPipeline.from_pretrained(
        "black-forest-labs/FLUX.1-schnell", 
        torch_dtype=flux_dtype
    ).to(device)
    print("Flux model loaded.")
except Exception as e:
    print(f"Error loading Flux model: {e}")
    flux_model = None

print(f"Loading Whisper model on {device}...")
try:
    whisper_model = whisper.load_model("base", device=device)
    print("Whisper model loaded.")
except Exception as e:
    print(f"Error loading Whisper model: {e}")
    whisper_model = None

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

# Image Generation Request Model
class GenerateRequest(BaseModel):
    prompt: str
    seed: int = 42
    steps: int = 4
    width: int = 512
    height: int = 512

# Story Generation Request Model
class StoryRequest(BaseModel):
    words: list[str]

# TTS Request Model
class TTSRequest(BaseModel):
    text: str

# Invoice Request Model
class InvoiceRequest(BaseModel):
    prompt: str | None = None

# Prompt Generation Request Model
class PromptRequest(BaseModel):
    word: str
    translation: str

# Product Extraction Request Model
class ProductRequest(BaseModel):
    content: str | None = None
    url: str | None = None
    myproduct: str | None = None
    prompt: str | None = None

@app.post("/generate")
async def generate_image(request: GenerateRequest):
    try:
        if flux_model is None:
            raise HTTPException(status_code=500, detail="Flux model is not loaded.")
            
        print(f"Generating image for prompt: {request.prompt}")
        output = flux_model(
            prompt=request.prompt,
            guidance_scale=0.0,
            num_inference_steps=request.steps,
            width=request.width,
            height=request.height,
            generator=torch.Generator(device=flux_model.device).manual_seed(request.seed)
        )
        pil_image = output.images[0]
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
        # Get extension and log it
        _, suffix = os.path.splitext(file.filename)
        suffix = suffix.lower()
        print(f"Received audio file: {file.filename} (Format: {suffix})")
        
        # Supported formats check (Optional, but good for logging)
        if suffix not in ['.mp3', '.m4a', '.wav', '.flac', '.ogg', '.aac']:
            print(f"Warning: format {suffix} is being passed to Whisper via FFmpeg.")

        # Create a temporary file to store the uploaded audio
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(file.file, tmp)
            temp_path = tmp.name
        
        # Transcribe using openai-whisper
        if whisper_model is None:
            raise HTTPException(status_code=500, detail="Whisper model is not loaded.")
            
        print(f"Transcribing {temp_path}...")
        result = whisper_model.transcribe(temp_path)
        
        return {"text": result["text"]}
    
    except Exception as e:
        print(f"Error transcribing audio: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Transcription error: {str(e)}")
    
    finally:
        # Clean up the temporary file
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

@app.post("/story")
async def generate_story(request: StoryRequest):
    try:
        words_str = ", ".join(request.words)
        print(f"Generating German story for words: {words_str}")
        
        prompt = f"Schreibe eine kurze, kreative Geschichte auf Deutsch, die die folgenden Wörter verwendet: {words_str}. Die Geschichte sollte kohärent sein und einen interessanten Handlungsbogen haben."
        
        ollama_payload = {
            "model": "llama3",
            "prompt": prompt,
            "stream": False
        }
        
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        return {"story": result["response"]}
        
    except Exception as e:
        print(f"Error generating story: {e}")
        raise HTTPException(status_code=500, detail=f"Story generation error: {str(e)}")

@app.post("/tts")
async def text_to_speech(request: TTSRequest):
    temp_wav = None
    temp_mp3 = None
    try:
        print(f"Converting text to neural speech: {request.text[:50]}...")
        
        # Get model and tokenizer
        model, tokenizer = get_tts_model()
        device = model.device
        
        # Prepare inputs
        inputs = tokenizer(request.text, return_tensors="pt").to(device)
        
        # Generate waveform
        with torch.no_grad():
            output = model(**inputs).waveform
            
        # Move to CPU and convert to numpy
        output_np = output.cpu().numpy().squeeze()
        
        # Create temp files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp_wav:
            temp_wav = tmp_wav.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_mp3:
            temp_mp3 = tmp_mp3.name
            
        # Save temporary WAV
        scipy.io.wavfile.write(temp_wav, rate=model.config.sampling_rate, data=output_np)
        
        # Convert WAV to MP3 using FFmpeg
        print(f"Converting {temp_wav} to {temp_mp3}...")
        subprocess.run(
            ["ffmpeg", "-y", "-i", temp_wav, "-codec:a", "libmp3lame", "-qscale:a", "2", temp_mp3],
            check=True,
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True
        )
        
        # Read the MP3 file
        with open(temp_mp3, "rb") as f:
            mp3_content = f.read()
            
        return Response(content=mp3_content, media_type="audio/mpeg", headers={"Content-Disposition": "attachment; filename=speech.mp3"})
        
    except Exception as e:
        print(f"Error in TTS: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"TTS error: {str(e)}")
        
    finally:
        # Cleanup
        for path in [temp_wav, temp_mp3]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/ocr")
async def extract_text(file: UploadFile = File(...)):
    try:
        print(f"Received image for OCR: {file.filename}")
        
        # Read image content
        contents = await file.read()
        
        # Resize if necessary
        contents, width, height, _, _, _ = resize_and_pad_image(contents)
        
        # Encode to base64 for Ollama
        image_base64 = base64.b64encode(contents).decode('utf-8')
        
        # Prepare request for Ollama
        ollama_payload = {
            "model": "llava:7b-v1.6-mistral-q4_0",
            "prompt": "Read all the text in this image and list it line by line. Do not say anything else, only the text.",
            "images": [image_base64],
            "stream": False
        }
        
        print("Calling Ollama Vision API...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        
        if response.status_code != 200:
            print(f"Ollama Error: {response.status_code} - {response.text}")
            response.raise_for_status()
        
        result = response.json()
        extracted_text = result["response"]
        
        # Split into list of lines and clean up
        lines = [line.strip() for line in extracted_text.split('\n') if line.strip()]
        
        return {"text": lines}
        
    except Exception as e:
        print(f"Error in OCR: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"OCR error: {str(e)}")

@app.post("/invoice")
async def process_invoice(file: UploadFile = File(...), prompt: str | None = Form(None)):
    try:
        print(f"Processing invoice: {file.filename}")
        
        # Read image content
        contents = await file.read()
        
        # Check if it's a PDF
        if file.filename.lower().endswith('.pdf'):
            print(f"Converting PDF {file.filename} to image...")
            # Load the PDF from bytes
            pdf_document = fitz.open(stream=contents, filetype="pdf")
            if pdf_document.page_count > 0:
                # Use the first page
                page = pdf_document[0]
                # Render page to a pixmap (image)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Use 2x scale for better OCR
                # Convert pixmap to PNG bytes
                img_byte_arr = io.BytesIO(pix.tobytes("png"))
                contents = img_byte_arr.getvalue()
                print("PDF page 1 converted to PNG.")
            pdf_document.close()

        # Resize if necessary
        contents, width, height, _, _, _ = resize_and_pad_image(contents)

        # Encode to base64 for Ollama
        image_base64 = base64.b64encode(contents).decode('utf-8')
        
        # Determine the prompt
        if prompt:
            print(f"Using custom prompt: {prompt}")
        else:
            prompt = (
                "Extract the following details from this invoice image and return them in structured JSON format: "
                "Invoice Number, Date, Vendor Name, Total Amount, Tax, and a list of Line Items with Description, Quantity and Price. "
                "Return ONLY the JSON. No conversational filler."
            )
        
        ollama_payload = {
            "model": "qwen2.5vl:3b",
            "prompt": prompt,
            "images": [image_base64],
            "stream": False,
            "format": "json" # Ollama supports forcing JSON format
        }
        
        print("Calling Qwen2.5-VL for invoice extraction...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        
        if response.status_code != 200:
            print(f"Ollama Error: {response.status_code} - {response.text}")
            response.raise_for_status()
        
        result = response.json()
        extracted_json_str = result["response"]
        
        # Parse the JSON string from LLM
        import json
        try:
            invoice_data = json.loads(extracted_json_str)
        except json.JSONDecodeError:
            # Fallback in case there's some markdown or filler (though 'format': 'json' should prevent this)
            print("Failed to parse direct JSON, attempting to extract from code blocks...")
            if "```json" in extracted_json_str:
                content = extracted_json_str.split("```json")[1].split("```")[0].strip()
                invoice_data = json.loads(content)
            else:
                raise Exception("Could not parse JSON from model response")
        
        return invoice_data
        
    except Exception as e:
        print(f"Error in Invoice Processing: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Invoice processing error: {str(e)}")

@app.post("/detect")
async def detect_objects(file: UploadFile = File(...), prompt: str | None = Form(None)):
    """
    Detects objects in an image and returns their pixel coordinates.
    Expects prompt to be something like 'List all buttons' or 'Find the close button'.
    """
    try:
        print(f"Detecting objects in: {file.filename}")
        
        # Read image content
        contents = await file.read()
        
        # Resize and Pad to 16:9
        print("Resizing and padding image to 16:9 ratio...")
        contents, orig_w, orig_h, canvas_w, canvas_h, scale = resize_and_pad_image(contents, ratio=(16, 9))

        # Encode to base64 for Ollama
        image_base64 = base64.b64encode(contents).decode('utf-8')
        
        # Determine the prompt
        if not prompt:
            prompt = "Detect all visible UI elements like buttons, inputs, and windows."
            
        grounding_prompt = (
            f"{prompt}. Return the results as a JSON list of objects with 'label' and 'bbox' (normalized 0-1000). "
            "Example format: [{\"label\": \"button\", \"bbox\": [xmin, ymin, xmax, ymax]}]. "
            "Only return the JSON list."
        )
        
        ollama_payload = {
            "model": "qwen2.5vl:3b",
            "prompt": grounding_prompt,
            "images": [image_base64],
            "stream": False,
            # "format": "json"
        }
        
        print("Calling Qwen2.5-VL for detection...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        
        if response.status_code != 200:
            print(f"Ollama Error: {response.status_code} - {response.text}")
            response.raise_for_status()
        
        result = response.json()
        raw_response = result["response"]
        print(f"Ollama Raw Response: {raw_response}")
        
        # Parse the JSON string from LLM
        import json
        try:
            parsed_data = json.loads(raw_response)
        except json.JSONDecodeError:
            # Fallback for non-compliant model responses
            print("Failed to parse direct JSON, attempting to extract from code blocks...")
            if "```json" in raw_response:
                json_content = raw_response.split("```json")[1].split("```")[0].strip()
                parsed_data = json.loads(json_content)
            else:
                raise Exception(f"Could not parse JSON from model response: {raw_response}")
        
        # Extract list of elements from dictionary if necessary
        detected_elements = []
        if isinstance(parsed_data, list):
            detected_elements = parsed_data
        elif isinstance(parsed_data, dict):
            # Check for common keys or just take the first list found
            if "elements" in parsed_data:
                detected_elements = parsed_data["elements"]
            elif "objects" in parsed_data:
                detected_elements = parsed_data["objects"]
            else:
                # Search for any value that is a list
                for val in parsed_data.values():
                    if isinstance(val, list):
                        detected_elements.extend(val)
                        # If we found a list of dicts, it's likely our data
                        if val and isinstance(val[0], dict):
                            break

        # Map coordinates back to pixels
        final_elements = []
        for item in detected_elements:
            # Skip items that are not dictionaries (defensive coding)
            if not isinstance(item, dict):
                print(f"Skipping non-dict item: {item}")
                continue
                
            label = item.get("label") or item.get("name") or "unknown"
            bbox = item.get("bbox") or item.get("box") or item.get("coordinates")
            
            if bbox and isinstance(bbox, list) and len(bbox) == 4:
                # Check if coordinates are normalized (0-1000) or pixels
                # Most VLMs output normalized, but if any value > 1000, it's likely pixels
                is_normalized = all(0 <= v <= 1000 for v in bbox)
                
                xmin, ymin, xmax, ymax = bbox
                
                if is_normalized:
                    # Model saw a 16:9 canvas, map back to original pixels
                    pixel_coords = {
                        "xmin": int((xmin * canvas_w / 1000) / scale),
                        "ymin": int((ymin * canvas_h / 1000) / scale),
                        "xmax": int((xmax * canvas_w / 1000) / scale),
                        "ymax": int((ymax * canvas_h / 1000) / scale)
                    }
                else:
                    # Already pixels (rare in normalized grounding, but for safety)
                    pixel_coords = {
                        "xmin": int(xmin / scale),
                        "ymin": int(ymin / scale),
                        "xmax": int(xmax / scale),
                        "ymax": int(ymax / scale)
                    }
                
                final_elements.append({
                    "label": label,
                    "bbox_normalized": bbox if is_normalized else None,
                    "bbox_pixels": pixel_coords
                })
        
        return {"elements": final_elements, "image_size": [orig_w, orig_h]}
        
    except Exception as e:
        print(f"Error in Detection API: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Detection error: {str(e)}")

@app.post("/locate-keywords")
async def locate_keywords(file: UploadFile = File(...), keywords: str = Form(...)):
    """
    Finds specific keywords in an image and returns their pixel coordinates.
    """
    try:
        print(f"Locating keywords in: {file.filename}")
        keyword_list = [k.strip() for k in keywords.split(",") if k.strip()]
        print(f"Keywords: {keyword_list}")
        
        # Read image content
        contents = await file.read()
        
        # Resize and Pad to Square (1:1) for best grounding accuracy
        print("Resizing and padding image to square...")
        contents, orig_w, orig_h, canvas_w, canvas_h, scale = resize_and_pad_image(contents, ratio=(1, 1))

        # Encode to base64 for Ollama
        image_base64 = base64.b64encode(contents).decode('utf-8')
        
        grounding_prompt = (
            f"Find the bounding boxes for the following keywords in this image: {', '.join(keyword_list)}. "
            "Return the results as a JSON list of objects with 'keyword' and 'bbox' (normalized 0-1000) where bbox is [xmin, ymin, xmax, ymax]. "
            "Only return the JSON list."
        )
        
        ollama_payload = {
            "model": "qwen2.5vl:3b",
            "prompt": grounding_prompt,
            "images": [image_base64],
            "stream": False,
        }
        
        print("Calling Qwen2.5-VL for keyword localization...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        
        if response.status_code != 200:
            print(f"Ollama Error: {response.status_code} - {response.text}")
            response.raise_for_status()
        
        result = response.json()
        raw_response = result["response"]
        print(f"Ollama Raw Response: {raw_response}")
        
        # Parse the JSON string from LLM
        import json
        try:
            detected_items = json.loads(raw_response)
        except json.JSONDecodeError:
            print("Failed to parse direct JSON, attempting to extract from code blocks...")
            if "```json" in raw_response:
                json_content = raw_response.split("```json")[1].split("```")[0].strip()
                detected_items = json.loads(json_content)
            else:
                # Last resort: try regex for anything that looks like a list
                match = re.search(r'\[\s*\{.*\}\s*\]', raw_response, re.DOTALL)
                if match:
                    detected_items = json.loads(match.group())
                else:
                    raise Exception(f"Could not parse JSON from model response: {raw_response}")
        
        # Map coordinates back to pixels
        final_results = []
        if isinstance(detected_items, list):
            for item in detected_items:
                if not isinstance(item, dict): continue
                keyword = item.get("keyword") or item.get("label") or "unknown"
                bbox = item.get("bbox")
                
                if bbox and isinstance(bbox, list) and len(bbox) == 4:
                    xmin, ymin, xmax, ymax = bbox
                    
                    # Model saw a 16:9 canvas, map back to original pixels
                    pixel_coords = {
                        "xmin": int((xmin * canvas_w / 1000) / scale),
                        "ymin": int((ymin * canvas_h / 1000) / scale),
                        "xmax": int((xmax * canvas_w / 1000) / scale),
                        "ymax": int((ymax * canvas_h / 1000) / scale)
                    }
                    
                    final_results.append({
                        "keyword": keyword,
                        "bbox_normalized": bbox,
                        "bbox_pixels": pixel_coords
                    })
        
        return {"results": final_results, "image_size": {"width": orig_w, "height": orig_h}}
        
    except Exception as e:
        print(f"Error in Locate Keywords API: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Localization error: {str(e)}")


@app.post("/prompt")
async def create_prompt_text(request: PromptRequest):
    try:
        print(f"Creating prompt text for word: {request.word} (Translation: {request.translation})")
        
        # Format the prompt for Llama3 to generate a visual description
        llm_prompt = (
            f"Describe a simple, clear, educational illustration for the word '{request.word}' ({request.translation}). "
            f"The image should be cartoon-style, clean background. Output only the visual description in English, max 15 words."
        )
        
        ollama_payload = {
            "model": "llama3",
            "prompt": llm_prompt,
            "stream": False
        }
        
        print(f"Calling Ollama to generate visual description for '{request.word}'...")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        generated_description = result["response"].strip().strip('"').strip("'")
        
        # Apply the required formatting
        final_prompt = f"generate an illustration image for: {generated_description}, clear background."
        print(f"Final Prompt: {final_prompt}")
        
        return {
            "prompt": final_prompt, 
            "word": request.word, 
            "translation": request.translation
        }
        
    except Exception as e:
        print(f"Error creating prompt: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Prompt generation error: {str(e)}")


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
                "3. Extract the 'price' for this specific target product. The price may be a formatted string (e.g., '18.690.000₫') or a numeric value (e.g., 18690000). Prioritize the main price shown for the target product.\n"
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
            "4. Extract 'price' (numeric value preferred, e.g., '18690000').\n"
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

@app.get("/health")
async def health():
    return {"status": "healthy"}

if __name__ == "__main__":
    # Ensure port 8000 is available or restart
    uvicorn.run(app, host="0.0.0.0", port=8000)
