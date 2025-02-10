import openai
import os
import pptx
from pptx import Presentation
from pptx.util import Inches
import json
from typing import List, Dict
import argparse
import hashlib
from datetime import datetime
from openai import AsyncOpenAI
import asyncio
import cProfile

# --- Constants and Configurations ---

LOG_FILE = "translation_log.txt"
SUPPORTED_EXTENSIONS = ('.pptx',)

# --- Helper Functions ---
def log_message(message: str, level: str = "INFO") -> None:
    """Logs messages to the console and a log file with timestamps."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}\n"
    print(log_entry, end="")
    with open(LOG_FILE, "a", encoding='utf-8') as logfile:
        logfile.write(log_entry)

def get_cache_filename(presentation_path: str, target_language: str) -> str:
    """Generates a unique cache filename based on the presentation and language."""
    base_name = os.path.splitext(os.path.basename(presentation_path))[0]
    path_hash = hashlib.sha256(presentation_path.encode('utf-8')).hexdigest()[:8]  # Short hash
    return f"translation_cache_{base_name}_{target_language}_{path_hash}.json"

def load_cache(cache_file: str) -> Dict:
    """Loads the translation cache from the specified file (if it exists)."""
    try:
        with open(cache_file, "r", encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_cache(cache: Dict, cache_file: str) -> None:
    """Saves the translation cache to the specified file."""
    with open(cache_file, "w", encoding='utf-8') as f:
        json.dump(cache, f, indent=4, ensure_ascii=False)

def generate_prompt_hash(prompt: str) -> str:
    """Generates a SHA-256 hash of the prompt for use as a cache key."""
    return hashlib.sha256(prompt.encode('utf-8')).hexdigest()

async def translate_text_with_openai(prompt: str, target_language: str, cache: Dict, max_retries: int = 3) -> str:
    """Translates text using the OpenAI API, with caching and retries."""
    prompt_hash = generate_prompt_hash(prompt)
    if prompt_hash in cache:
        log_message(f"Using cached translation for: {prompt[:50]}...", level="CACHE_HIT")
        return cache[prompt_hash]

    log_message(f"Translating: {prompt[:50]}... to {target_language}", level="API_CALL")

    schema_instruction = (
        "Return the translation as a JSON object exactly as follows: "
        "{\"translated\": \"<translated text>\"}"
    )

    system_instruction = (
        f"You are a helpful assistant that translates text to {target_language}. "
        "Maintain the original meaning as closely as possible. "
        f"Adjust the tone of the translation to be appropriate for professional presentations in the target language ({target_language}). "
        "The translated text should be approximately the same character length as the original text (within a 5% margin). "
        f"{schema_instruction}"
    )

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OpenAI API key not found. Please set the OPENAI_API_KEY environment variable.")

    client = AsyncOpenAI(api_key=api_key, timeout=30.0)

    for attempt in range(max_retries):
        try:
            response = await client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=4096,
                timeout=30,
                response_format={"type": "json_object"}
            )
            result = response.choices[0].message.content.strip()

            for parse_attempt in range(max_retries):
                try:
                    parsed = json.loads(result)
                    translated_text = parsed["translated"]
                    cache[prompt_hash] = translated_text
                    return translated_text  # Return immediately after saving to cache
                except (json.JSONDecodeError, KeyError) as e:
                    log_message(f"Error parsing response (attempt {parse_attempt+1}/{max_retries}): {str(e)}", level="ERROR")
                    if parse_attempt == max_retries - 1:
                        log_message(f"Failed to parse response after {max_retries} attempts. Returning original prompt.", level="ERROR")
                        return prompt
                except Exception as e:
                    log_message(f"Unexpected error during parsing: {e}", level="ERROR")
                    return prompt

        except Exception as e:
            log_message(f"OpenAI API Error (Attempt {attempt + 1}/{max_retries}): {str(e)}", level="ERROR")
            if attempt == max_retries - 1:
                return prompt
    return prompt


def extract_text_from_presentation(presentation_path: str) -> List[Dict]:
    """Extracts text and context from a PowerPoint presentation."""
    try:
        prs = Presentation(presentation_path)
        text_data = []

        for slide_number, slide in enumerate(prs.slides, start=1):
            for shape_index, shape in enumerate(slide.shapes):
                shape_id = f"slide{slide_number}_shape{shape_index}"
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.text.strip():
                                shape_type = "UNKNOWN"
                                if shape == slide.shapes.title:
                                    shape_type = "TITLE"
                                elif shape.has_table:
                                    shape_type = "TABLE"
                                else:
                                    shape_type = "BODY"

                                text_data.append({
                                    "slide_number": slide_number,
                                    "shape_type": shape_type,
                                    "text": run.text,
                                    "shape_id": shape_id,
                                })

                elif shape.has_table:
                    for row_idx, row in enumerate(shape.table.rows):
                        for col_idx, cell in enumerate(row.cells):
                            if cell.text.strip():
                                text_data.append({
                                    "slide_number": slide_number,
                                    "shape_type": "TABLE",
                                    "text": cell.text,
                                    "shape_id": f"{shape_id}_row{row_idx}_col{col_idx}"
                                })

        return text_data

    except Exception as e:
        log_message(f"Error extracting text from {presentation_path}: {e}", level="ERROR")
        return []


async def batch_translate_texts_with_openai(text_entries: List[Dict], target_language: str, cache: Dict, max_retries: int = 3, batch_size: int = 10) -> None:
    """Batch translates multiple texts using the OpenAI API with structured JSON output."""
    texts_to_translate = []
    for entry in text_entries:
        prompt_hash = generate_prompt_hash(entry["text"])
        if prompt_hash not in cache:
            texts_to_translate.append((prompt_hash, entry["text"]))

    if not texts_to_translate:
        log_message("All translations found in cache.", level="INFO")
        return

    total_batches = (len(texts_to_translate) + batch_size - 1) // batch_size
    log_message(f"Starting batch translation of {len(texts_to_translate)} texts in {total_batches} batches", level="INFO")

    tasks = []
    for batch_num, i in enumerate(range(0, len(texts_to_translate), batch_size), 1):
        batch = texts_to_translate[i:i + batch_size]
        payload = {hash_: text for hash_, text in batch}

        log_message(f"Processing batch {batch_num}/{total_batches} ({len(batch)} texts)", level="INFO")

        schema_instruction = (
            "Return the translations as a JSON object exactly as follows: \n"
            "{\"translations\": {\"<sha256 hash>\": \"<translated text>\"} }"
        )

        system_instruction = (
            f"You are a helpful assistant that translates multiple texts to {target_language}. "
            "Maintain the original meaning as closely as possible. "
            f"Adjust the tone of each translation to be appropriate for professional presentations in the target language ({target_language}). "
            "The translated text for each input should be approximately the same length as the original text (within a 10% margin). "
            f"{schema_instruction}"
        )

        prompt_data = {
            "texts": payload,
            "target_language": target_language,
            "instructions": "Translate each text, maintaining original meaning and formatting."
        }

        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OpenAI API key not found. Please set the OPENAI_API_KEY environment variable.")

        client = AsyncOpenAI(api_key=api_key, timeout=60.0)

        tasks.append(translate_batch(client, system_instruction, prompt_data, cache, max_retries, batch_num, total_batches))

    await asyncio.gather(*tasks)
    log_message("Batch translation completed", level="INFO")


async def translate_batch(client: AsyncOpenAI, system_instruction:str, prompt_data: Dict, cache: Dict, max_retries: int, batch_num: int, total_batches: int) -> None:
    """Translates a single batch (async).  This is now a separate function."""
    for attempt in range(max_retries):
        try:
            response = await client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": json.dumps(prompt_data)}
                ],
                temperature=0.2,
                max_tokens=4096,
                timeout=60,
                response_format={"type": "json_object"}
            )
            output = response.choices[0].message.content.strip()

            for parse_attempt in range(max_retries):
                try:
                    result = json.loads(output)
                    translations = result.get("translations", {})
                    for hash_, translated_text in translations.items():
                        cache[hash_] = translated_text
                    break  # Success, break out of parsing attempts
                except (json.JSONDecodeError, KeyError) as e:
                    log_message(f"Error parsing batch {batch_num} response (attempt {parse_attempt+1}/{max_retries}): {str(e)}", level="ERROR")
                    if parse_attempt == max_retries - 1:
                        log_message("Failed to parse batch response after all retries", level="ERROR")
                except Exception as e:
                    log_message(f"Unexpected error during parsing: {e}", level="ERROR")
            else:
                continue  # Failed all parsing attempts, retry API call
            break  # Success, break out of API call attempts

        except Exception as e:
            log_message(f"Batch {batch_num} translation error (Attempt {attempt + 1}/{max_retries}): {str(e)}", level="ERROR")
            if attempt == max_retries - 1:
                log_message(f"Max retries reached for batch {batch_num}", level="ERROR")


async def translate_presentation(presentation_path: str, target_language: str, output_path: str) -> None:
    """Translates a PowerPoint presentation and saves the translated version."""

    if not presentation_path.lower().endswith(SUPPORTED_EXTENSIONS):
        log_message(f"Unsupported file format: {presentation_path}.  Skipping.", level="WARNING")
        return

    if presentation_path == output_path:
        log_message("Input and output paths are the same.  This is not allowed.", level="ERROR")
        return

    try:
        text_data = extract_text_from_presentation(presentation_path)
        if not text_data:
            log_message(f"No text found to translate in {presentation_path}.", level="WARNING")
            return

        # Get the cache filename based on presentation path and language
        cache_file = get_cache_filename(presentation_path, target_language)
        cache = load_cache(cache_file)

        await batch_translate_texts_with_openai(text_data, target_language, cache)

        # Save the cache after batch translation
        save_cache(cache, cache_file)

        prs = Presentation(presentation_path)
        prs.save(output_path)

        translated_prs = Presentation(output_path)

        translated_text_data = []
        for text_entry in text_data:
            prompt_hash = generate_prompt_hash(text_entry["text"])
            translated_text = cache.get(prompt_hash, text_entry["text"])  # Fallback to original
            translated_text_entry = text_entry.copy()
            translated_text_entry["translated_text"] = translated_text
            translated_text_data.append(translated_text_entry)

        for slide_number, slide in enumerate(translated_prs.slides, start=1):
            for shape_index, shape in enumerate(slide.shapes):
                shape_id = f"slide{slide_number}_shape{shape_index}"

                translated_text_entry = next((entry for entry in translated_text_data if entry["shape_id"] == shape_id), None)

                if translated_text_entry:
                    if shape.has_text_frame:
                        try:
                            text_frame = shape.text_frame
                            for p_idx, paragraph in enumerate(text_frame.paragraphs):
                                for run_idx, run in enumerate(paragraph.runs):
                                    if run.text.strip():
                                        for entry in translated_text_data:
                                            if entry["shape_id"] == shape_id and entry["text"] == run.text:
                                                run.text = entry["translated_text"]
                                                break

                        except Exception as e:
                            log_message(f"Error modifying text in shape {shape_id} on slide {slide_number}: {e}", level="ERROR")
                            continue

                    elif shape.has_table:
                        try:
                            for row_idx, row in enumerate(shape.table.rows):
                                for col_idx, cell in enumerate(row.cells):
                                    cell_shape_id = f"{shape_id}_row{row_idx}_col{col_idx}"
                                    cell_translated_text_entry = next((entry for entry in translated_text_data if entry["shape_id"] == cell_shape_id), None)
                                    if cell_translated_text_entry:
                                        cell.text = cell_translated_text_entry["translated_text"]

                        except Exception as e:
                            log_message(f"Error modifying table {shape_id} on slide {slide_number}: {e}", level="ERROR")
                            continue

        translated_prs.save(output_path)
        log_message(f"Translated presentation saved to: {output_path}", level="SUCCESS")

    except Exception as e:
        log_message(f"Error during translation process: {e}", level="ERROR")


def main():
    """Main function to handle command-line arguments and process presentations."""

    parser = argparse.ArgumentParser(description="Translate PowerPoint presentations using the OpenAI API.")
    parser.add_argument("input_path", nargs='?', type=str, help="Path to the input PowerPoint file or directory.")
    parser.add_argument("-o", "--output", type=str, help="Output file or directory path.  If not specified, defaults to [original_filename]_translated.[ext]")
    parser.add_argument("-l", "--language", type=str, required=True, help="Target language code (e.g., es, fr, zh-CN)")
    parser.add_argument("-k", "--api_key", type=str, help="OpenAI API Key. If not provided, will check the OPENAI_API_KEY environment variable.")
    parser.add_argument("-p", "--profile", action="store_true", help="Enable profiling.")

    args = parser.parse_args()

    if args.api_key:
        os.environ["OPENAI_API_KEY"] = args.api_key
    elif not os.environ.get("OPENAI_API_KEY"):
        log_message("OpenAI API Key not found. Please set the OPENAI_API_KEY environment variable or use the -k option.", level="ERROR")
        return

    input_path = args.input_path
    output_path = args.output
    target_language = args.language

    if not input_path:
        log_message("No input file or directory specified.  Processing all .pptx files in the current directory.", level="INFO")
        for filename in os.listdir("."):
            if filename.lower().endswith(SUPPORTED_EXTENSIONS):
                default_output_path = filename.replace(".pptx", f"_translated_{target_language}.pptx")
                asyncio.run(translate_presentation(filename, target_language, default_output_path))
        return

    if os.path.isfile(input_path):
        if not output_path:
            base, ext = os.path.splitext(input_path)
            output_path = f"{base}_translated_{target_language}{ext}"
        if args.profile:
            cProfile.run(f"translate_presentation('{input_path}', '{target_language}', '{output_path}')", sort="tottime")
        else:
            asyncio.run(translate_presentation(input_path, target_language, output_path))
    elif os.path.isdir(input_path):
        if not output_path:
            output_path = input_path
        elif not os.path.isdir(output_path):
            log_message("If input is a directory, output must also be a directory (or not specified).", level="ERROR")
            return

        for filename in os.listdir(input_path):
            if filename.lower().endswith(SUPPORTED_EXTENSIONS):
                full_input_path = os.path.join(input_path, filename)
                base, ext = os.path.splitext(filename)
                full_output_path = os.path.join(output_path, f"{base}_translated_{target_language}{ext}")
                if args.profile:
                    cProfile.run(f"translate_presentation('{full_input_path}', '{target_language}', '{full_output_path}')", sort="tottime")
                else:
                    asyncio.run(translate_presentation(full_input_path, target_language, full_output_path))
    else:
        log_message(f"Invalid input path: {input_path}", level="ERROR")


if __name__ == "__main__":
    main()
