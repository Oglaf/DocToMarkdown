# Document to Markdown Converter for Azure DevOps

This is a graphical user interface (GUI) application designed to convert `.docx` documents into Azure DevOps-compliant Markdown files. It uses Pandoc for the core conversion and offers an optional AI post-processing step via Azure OpenAI to refine the output.

## Features

  - **DOCX to Markdown**: Converts `.docx` files to clean Markdown.
  - **Azure DevOps Compliant**: Automatically formats image links and filenames to work seamlessly with Azure DevOps Wikis.
  - **Image Extraction**: Extracts images from the document and places them in a `.attachments` folder at the wiki root.
  - **AI Post-Processing**: Optionally uses Azure OpenAI to refine the converted text based on a custom prompt.
  - **Persistent Configuration**: Saves your settings (paths, credentials) to a local, encrypted `config.ini` file so you don't have to enter them every time.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

1.  **Python 3.8+**: Make sure Python is installed and accessible from your command line. You can download it from [python.org](https://python.org/).
2.  **Pandoc**: This tool is required for the document conversion. You can download it from [pandoc.org](https://pandoc.org/installing.html).

## Setup

Follow these steps to get the application running.

### 1\. Create `requirements.txt`

Create a file named `requirements.txt` in the same folder as the script and add the following lines to it:

```text
cryptography
openai
```

### 2\. Create a Virtual Environment

It is highly recommended to use a Python virtual environment to avoid conflicts with other projects.

```bash
# Create the virtual environment
python -m venv .venv

# Activate it (on Windows)
.\.venv\Scripts\activate

# Activate it (on macOS/Linux)
source .venv/bin/activate
```

### 3\. Install Dependencies

Install all the required Python packages using the `requirements.txt` file you created.

```bash
pip install -r requirements.txt
```

## Usage

1.  **Launch the Application**: Run the `DocToMarkdown.py` script from your activated virtual environment.
    ```bash
    python DocToMarkdown.py
    ```
2.  **Configure Settings**:
      * **Pandoc Path**: The first time you run the app, click "Browse..." to locate your `pandoc.exe` file. This is usually in `C:\Program Files\Pandoc\pandoc.exe`.
      * **File and Folder Selection**:
          * **Document File**: Select the `.docx` file you want to convert.
          * **Output Folder**: Choose the folder where the `.md` file will be saved (e.g., a subfolder in your wiki).
          * **Wiki Root Folder**: Select the main root folder of your Azure DevOps Wiki. The `.attachments` folder will be created here.
      * **AI Post-Processing (Optional)**:
          * Check the "Post-process with AI" box to enable this feature.
          * Fill in your Azure OpenAI credentials (Endpoint, Key, and Deployment name).
          * Write a prompt to guide the AI in refining the text.
3.  **Save and Convert**:
      * Click **"Save Settings"** to encrypt and save your paths and credentials for future use.
      * Click **"Convert to Markdown"** to start the process. The log window will show the progress.