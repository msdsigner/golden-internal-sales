# Excel Automation with PDF and Google Drive

> [!NOTE]
> **🤖 AI-Assisted Project:** This project uses an AI methodology managed via [CLAUDE.md](file:///d:/Golden%20Auto/CLAUDE.md). If you are using an AI coding assistant, point it to that file first.

This automation tool cleans, filters Excel files, generates a PDF summary, and uploads the results to Google Drive.

## Setup

1.  **Install Python Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

2.  **Google Drive API Setup:**
    - Go to the [Google Cloud Console](https://console.cloud.google.com/).
    - Create a new project.
    - Enable the **Google Drive API**.
    - Go to "APIs & Services" > "Credentials".
    - Click "Create Credentials" > "OAuth client ID".
    - Set the "Application type" to "Desktop app".
    - Download the JSON file and rename it to `credentials.json`.
    - Place `credentials.json` in the root of this project (the same folder as `main.py`).

3.  **Input Data:**
    - Place any `.xlsx` files you want to process in the `input/` folder.

## Usage

Run the main script to start the automation:

```bash
python main.py
```

- **Cleaning Logic:** The script removes empty rows/cols, standardizes column names, and removes duplicates. You can customize this in `scripts/clean_data.py`.
- **Filtering Logic:** You can specify filtering criteria in `scripts/clean_data.py` (e.g., specific statuses or values).
- **Output:** Cleaned files and PDF reports are saved in the `output/` folder.
- **Upload:** The script will attempt to upload the results to Google Drive. On the first run, it will open your browser to request authorization.

## Configuration

In `main.py`, you can set a `FOLDER_ID` if you want to upload files to a specific Google Drive folder:

```python
FOLDER_ID = "your_google_drive_folder_id_here"
```
