"""Claude API interaction: upload a PDF, stream the response, parse the JSON."""

import base64
import json
import re
from pathlib import Path

import anthropic

from .constants import MODEL, MAX_TOKENS
from .prompts import SYSTEM_PROMPT, USER_PROMPT


def upload_pdf_and_get_response(pdf_path: str, api_key: str) -> dict:
    """Upload PDF to Claude and get structured JSON response."""
    client = anthropic.Anthropic(api_key=api_key, timeout=3600.0)

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
    print(f"[INFO] PDF file: {pdf_path.name} ({file_size_mb:.1f} MB)")

    # Use Files API for larger files, base64 for smaller ones
    if file_size_mb > 20:
        print("[INFO] Using Files API for large file upload...")
        with open(pdf_path, "rb") as f:
            file_upload = client.files.upload(
                file=(pdf_path.name, f, "application/pdf")
            )
        content = [
            {
                "type": "document",
                "source": {
                    "type": "file",
                    "file_id": file_upload.id,
                },
            },
            {
                "type": "text",
                "text": USER_PROMPT,
            },
        ]
        print(f"[INFO] File uploaded. ID: {file_upload.id}")
    else:
        print("[INFO] Using base64 inline upload...")
        with open(pdf_path, "rb") as f:
            pdf_data = base64.standard_b64encode(f.read()).decode("utf-8")
        content = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": pdf_data,
                },
            },
            {
                "type": "text",
                "text": USER_PROMPT,
            },
        ]

    print(f"[INFO] Sending request to {MODEL} (streaming enabled)...")
    print("[INFO] Receiving response from Claude...", end="", flush=True)

    response_text = ""
    with client.messages.stream(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": content,
            }
        ],
    ) as stream:
        for text in stream.text_stream:
            response_text += text
            print(".", end="", flush=True)

        message = stream.get_final_message()

    print()
    print(f"[INFO] Response received. Usage: {message.usage}")

    # Check if response was truncated
    if message.stop_reason == "max_tokens":
        print("[WARNING] Response was truncated due to max_tokens limit.")
        print("[WARNING] The output document may be incomplete.")
        print("[WARNING] Consider splitting the PDF into smaller parts.")

    # Parse JSON from response
    json_data = extract_json(response_text)
    return json_data


def extract_json(text: str) -> dict:
    """Extract JSON object from Claude's response text."""
    # Try direct parse first
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try to find JSON within markdown code blocks
    json_match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Try to find first { ... } block
    brace_start = text.find("{")
    if brace_start != -1:
        depth = 0
        for i in range(brace_start, len(text)):
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[brace_start : i + 1])
                    except json.JSONDecodeError:
                        break

    raise ValueError(
        "Could not parse JSON from Claude's response.\n"
        f"Response preview: {text[:500]}..."
    )
