[![English](https://img.shields.io/badge/lang-English-blue.svg)](README.en.md)
[![Chinese](https://img.shields.io/badge/lang-简体中文-red.svg)](README.zh-CN.md)
[![Japanese](https://img.shields.io/badge/lang-日本語-green.svg)](README.md)

----

# Project Analysis: Office Files Quality Check Tool

## Usage Examples
1. **Simple Operation Interface**
    <img src="./picture/gui.png" width="600">
    A straightforward design allowing you to select the target folder and output destination. The lightweight Tkinter-based UI ensures anyone can start operating it without confusion.

2. **Real-time Text Extraction and Log Display**
    <img src="./picture/extract_and_log.png" width="600">
    During execution, the text extraction status from each file is displayed in real-time in the log window. You can instantly grasp whether processing has completed successfully and which files are currently being processed.

3. **Batch Hyperlink Extraction and Duplication Check**
    <img src="./picture/url_check.png" width="600">
    Comprehensively extracts and lists URLs within documents. A unique deduplication logic helps organize common references from large numbers of files, facilitating smooth batch verification in a browser.

4. **AI-powered Quality Check (LLM Integration)**
    [Prompt for AI analysis of extracted text available here](./prompt/QC-prompt.md)
    <img src="./picture/copilot_qc.png" width="600">
    By inputting the extracted Markdown-formatted text into an LLM like Copilot, you can receive advanced proofreading assistance, including suggestions for typos, inconsistent expressions, and logical contradictions.

## Project Overview

This project is a Python-based GUI application designed to extract content from Office documents (specifically `.docx`, `.xlsx`, `.xlsm`, and `.pptx` files). It leverages Tkinter for the user interface and dedicated Python libraries for document parsing.

**Core functionalities include:**

*   **Text Extraction:** Extracts plain text content from Office documents and saves it into Markdown (`.md`) files. These are organized into timestamped `PlainText_YYYYMMDDHHMMSS` directories within the user-specified output path.
*   **Hyperlink Extraction:** Extracts hyperlinks from Office documents and saves them into `.csv` files (structured as CSV with specific headers and extensions). These are organized into timestamped `HyperLinks_YYYYMMDDHHMMSS` directories.
*   **Batch URL Opening:** Provides a feature to open all extracted hyperlinks from the `.csv` files. This functionality includes robust global deduplication, ensuring each unique URL is opened only once across all processed documents. Before opening, it generates an informational HTML page summarizing the process and opens it in the default web browser.

**Key Technologies:**

*   **Language:** Python
*   **GUI Framework:** Tkinter
*   **Document Parsing Libraries:** `python-docx`, `openpyxl`, `python-pptx`
*   **Other Libraries:** `csv`, `os`, `sys`, `threading`, `webbrowser`, `datetime`

**Architecture:**

*   **GUI Layer:** `gui.py` - Handles user interactions, window management, event binding, and progress display.
*   **Core Logic Layer:** `core/` directory containing:
    *   `extractor.py`: Implements the logic for parsing Office files, extracting text, and hyperlinks.
    *   `url_opener.py`: Manages reading `.csv` files, performing URL deduplication, and opening URLs in the browser.

## Building and Running

**1. Install Dependencies:**

To set up the project environment and install all necessary Python packages, run the following command in your terminal:

```bash
pip install -r requirements.txt
```

**2. Run the Application:**

To launch the GUI application, execute the main script:

```bash
python gui.py
```

This will open the graphical user interface, allowing you to select input and output directories and initiate the extraction process.

## Development Conventions

*   **Code Organization:** The project is structured with a clear separation of concerns, placing core business logic in the `core/` directory and the user interface in `gui.py`.
*   **Comments and Documentation:** Code includes extensive comments, primarily in Japanese, explaining the purpose of functions, classes, and specific logic blocks. Docstrings are used for methods.
*   **Error Handling:** Robust error handling is implemented using `try-except` blocks for file operations, library imports, and data processing. User-facing errors are communicated through Tkinter's `messagebox` and redirected to the console/log window (`sys.stderr`).
*   **User Feedback:** The GUI provides visual feedback through a progress bar (`ttk.Progressbar`) and a dedicated log area that captures `stdout` and `stderr` output.
*   **Asynchronous Operations:** For long-running extraction tasks, a `threading.Thread` is used to prevent the GUI from freezing, ensuring a responsive user experience.
*   **File Naming Conventions:** Extracted files follow a consistent naming scheme:
    *   **Text:** `PlainText_[OriginalFileName]_[OriginalExtension].md` (e.g., `PlainText_Report_docx.md`)
    *   **Hyperlinks:** `Urls_[OriginalFileName]_[OriginalExtension].csv` (e.g., `Urls_Report_docx.csv`)
    *   Output directories are timestamped (e.g., `PlainText_20251227103000`).
*   **URL Deduplication:** The `url_opener.py` module implements both within-file and global deduplication of URLs to avoid redundant operations.
