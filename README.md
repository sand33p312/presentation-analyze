# AI Presentation Inconsistency Checker

This is a Python tool that uses the Google Gemini 2.5 Flash model to analyze PowerPoint (`.pptx`) presentations for factual and logical inconsistencies across slides.

## Features

-   **Automated Content Extraction:** The script automatically extracts all textual content and images from every slide in a `.pptx` file.
-   **Multimodal Analysis:** It leverages Gemini's multimodal capabilities to understand both text and visuals (like charts and diagrams in images).
-   **Inconsistency Detection:** The core feature is identifying three main types of issues:
    -   Conflicting numerical data.
    -   Contradictory textual claims.
    -   Timeline mismatches.
-   **Structured Terminal Output:** The findings are presented in a clear, easy-to-read report directly in the terminal, referencing the specific slide numbers for each issue.



## How to Run

1.  **Clone the repository:**
    ```bash
    git clone <your-repo-url>
    cd ai-ppt-checker
    ```
2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Set up your API Key:**
    Create a `.env` file in the root directory and add your Google API Key:
    ```
    GOOGLE_API_KEY="your-key-here"
    ```
4.  **Run the script:**
    ```bash
    python main.py /path/to/your/presentation.pptx
    ```