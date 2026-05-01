# Facebook Marketplace Automator

A Python script that uses Google's Gemini 2.5 Flash-Lite API to process a folder of product images, automatically extract details, and generate a pre-filled Excel spreadsheet for bulk uploading to Facebook Marketplace.

## Features
* **Automated Data Extraction:** Identifies product name, estimates a fair price, determines condition, writes a bulleted description, and categorizes the item.
* **Cost Effective:** Uses the `gemini-2.5-flash-lite` model. Processing images usually costs fractions of a penny.
* **Resume Capability:** If the script halts (or you hit API quotas), simply run it again. It automatically reads the Excel file and picks up exactly where it left off.
* **Large File Support:** Automatically switches to Google's File Upload API for images over 4MB.

## Prerequisites
1. Python 3.10+ installed.
2. A free Google Gemini API Key from [Google AI Studio](https://aistudio.google.com/).

## Setup & Installation

1. Clone or download this repository.
2. Install the required Python libraries:
   ```bash
   pip install tqdm openpyxl google-genai Pillow
Set up your API key. You can do this in two ways:

* Option A (Recommended): Set it as an environment variable named GEMINI_API_KEY.

* Option B: Create a file named api_key.txt in the same folder as the script and paste your API key inside it (no quotes).

Ensure the Marketplace_Bulk_Upload_Template.xlsx file is in the root directory. Make sure it has no data below the header row.

## Usage
Place all your product images into the Input folder. (The script will create this folder the first time you run it).

Run the script:

```Bash
python FBMP_Automator.py
```
The script will create a new timestamped session folder inside the Output directory. This folder will contain your renamed images, a log file, and your completed Excel sheet.

## Known Limitations & Quirks
* Facebook Category Mapping: When uploading the generated Excel file to Facebook, the Marketplace interface occasionally fails to map the Category column correctly, even if the data is perfectly formatted. You may need to manually select the final category from Facebook's dropdown menu for some items before hitting publish.

* Free Tier Quotas: If you are using a free-tier API key without a billing account attached, Google limits requests to roughly 20 per day. The script will safely halt when it hits this limit and can be resumed 24 hours later.
