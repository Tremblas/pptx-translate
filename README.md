# PPTX Translate

## Overview
PPTX Translate is a Python command-line tool that leverages the OpenAI API to translate text in PowerPoint (.pptx) presentations to a specified target language. The tool extracts text from slides, translates it while preserving the original context and formatting, and generates a new presentation with the translated content.

## Features
- Extracts text from PowerPoint presentations including titles, body text, and tables.
- Uses asynchronous batch translation with the OpenAI API.
- Implements a caching mechanism to avoid redundant API calls.
- Provides detailed logging for actions and errors.
- Command-line interface for easy usage on individual files or directories.

## Prerequisites
- Python 3.7 or higher.
- Required Python packages:
  - python-pptx
  - openai
  - argparse
- An OpenAI API key should be available via the `OPENAI_API_KEY` environment variable or provided as a command-line argument.

## Installation
1. Clone the repository:
   ```
   git clone <repository-url>
   cd pptx-translate
   ```
2. Install the necessary dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
Run the script with the following basic command:
```
python script.py input_file.pptx -l <target_language>
```
Options:
- `input_file.pptx`: The path to the PowerPoint file to be translated.
- `-l, --language`: Target language code (e.g. es, fr, zh-CN).
- `-o, --output`: Output file path. If not specified, the output file will be named with `_translated_<language>` appended to the original filename.
- `-k, --api_key`: Specify the OpenAI API Key if it's not set in the environment.
- `-p, --profile`: Enable profiling for performance analysis.

Examples:
- Translate a single presentation to Spanish:
  ```
  python script.py presentation.pptx -l es
  ```
- Translate a presentation and specify an output file:
  ```
  python script.py presentation.pptx -l fr -o presentation_translated_fr.pptx
  ```
- Process all `.pptx` files in the current directory:
  Simply run:
  ```
  python script.py -l es
  ```

## Logging and Caching
- Logging is done to the console and to `translation_log.txt` for tracking translation progress, cache hits, and errors.
- Translations are cached in JSON files (e.g., `translation_cache_<filename>_<language>_<hash>.json`) to reduce API calls and accelerate subsequent translations.

## Contribution
Contributions are welcome! If you have improvements or bug fixes, feel free to fork the repository and create a pull request.

## License
Copyright 2025 Thiago Peres.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.