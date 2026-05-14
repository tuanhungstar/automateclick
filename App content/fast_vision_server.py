import io
import json
import base64
import uvicorn
import requests
import re
import logging
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from pydantic import BaseModel
from PIL import Image

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Vision Extraction API")

# Configuration - Ollama usually runs on 11434 by default
OLLAMA_BASE_URL = "http://localhost:11434"

def resize_and_pad_image(image_bytes, max_dim=1024, ratio=None, center=True):
    """
    Resizes image to fit max_dim. 
    If ratio is provided (e.g. (16, 9)), pads the image to that aspect ratio.
    Returns: (processed_bytes, original_w, original_h, canvas_w, canvas_h, scale_factor, offset_x, offset_y)
    """
    img = Image.open(io.BytesIO(image_bytes))
    if img.mode != "RGB":
        img = img.convert("RGB")
    orig_w, orig_h = img.size
    
    if ratio:
        target_ratio = ratio[0] / ratio[1]
        if target_ratio > (orig_w / orig_h): # Canvas is wider than image
            canvas_h = max_dim
            canvas_w = int(max_dim * target_ratio)
        else: # Canvas is taller than image
            canvas_w = max_dim
            canvas_h = int(max_dim / target_ratio)
            
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
        if max(orig_w, orig_h) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            new_w, new_h = img.size
            scale = new_w / orig_w
            return buf.getvalue(), orig_w, orig_h, new_w, new_h, scale, 0, 0
        return image_bytes, orig_w, orig_h, orig_w, orig_h, 1.0, 0, 0

@app.post("/extract-product-from-image")
async def extract_product_from_image(file: UploadFile = File(...), myproduct: str | None = Form(None), prompt: str | None = Form(None)):
    try:
        logger.info(f"Received request for image: {file.filename}")
        
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
        
        logger.info(f"Calling Ollama (qwen2.5vl:7b) for product: {target_product}")
        response = requests.post(f"{OLLAMA_BASE_URL}/api/generate", json=ollama_payload)
        response.raise_for_status()
        
        result = response.json()
        raw_res = result.get("response", "{}")
        
        try:
            extracted_data = json.loads(raw_res.strip())
            
            # Ensure all required keys exist
            expected_keys = ["product_name", "price", "status", "similarity"]
            for key in expected_keys:
                if key not in extracted_data:
                    extracted_data[key] = ""
            
            return extracted_data
        except Exception as parse_error:
            logger.error(f"JSON Parse Error: {parse_error}. Raw response: {raw_res}")
            match = re.search(r'\{.*\}', raw_res, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group())
                except: pass
            return {"error": "Failed to parse JSON", "raw": raw_res}
            
    except Exception as e:
        logger.error(f"Extraction error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
async def health():
    return {"status": "ok"}

if __name__ == "__main__":
    # Windows-friendly uvicorn invocation
    uvicorn.run(app, host="0.0.0.0", port=8000)
