# Delivery and Return List PDF Generator

This repository contains a Google Apps Script that automates the generation and management of delivery and return list PDFs for couriers. The script processes data from a Google Sheet, generates PDF files, organizes them into a folder structure in Google Drive, and provides direct download links in the Google Sheet. Additionally, it merges multiple PDFs into a single file for each courier using the CloudConvert API.

## Features

- **Automated PDF Generation**: Generates delivery and return lists as PDFs from Google Sheets data.
- **Folder Organization**: Automatically creates and organizes folders in Google Drive based on the current date.
- **PDF Merging**: Merges multiple PDFs into a single file for each courier when necessary using CloudConvert.
- **Direct Download Links**: Inserts direct download links for the PDFs back into the Google Sheet.
- **Error Handling and Retries**: Includes delays and retries to handle Google Sheets calculations and API calls reliably.

## Setup Instructions

### Prerequisites

1. **Google Account**: You need a Google account to access Google Sheets and Google Drive.
2. **Google Apps Script**: Basic understanding of Google Apps Script.
3. **CloudConvert API Key**: You need to sign up for a CloudConvert account and obtain an API key to enable PDF merging.

### Step-by-Step Setup

1. **Create the Google Sheet**:
   - Create a Google Sheet with at least three sheets named: `Daily Deliveries`, `Delivery List`, and `Return List`.
   - Fill in the `Daily Deliveries` sheet with delivery and return IDs, order IDs, and carrier information starting from row 8.

2. **Deploy the Script**:
   - Open the Google Sheet.
   - Navigate to `Extensions > Apps Script`.
   - Copy and paste the provided script into the Apps Script editor.
   - Save the script.

3. **Set Up CloudConvert**:
   - Obtain a CloudConvert API key by creating an account on [CloudConvert](https://cloudconvert.com/).
   - Replace the placeholder API key in the script with your actual CloudConvert API key.

4. **Running the Script**:
   - Manually run the `generateDeliveryPDFs` function to generate the PDFs.
   - The script will create folders in Google Drive, generate PDFs, merge them if necessary, and insert download links in the Google Sheet.

## Script Functions

- **generateDeliveryPDFs()**: Main function to generate delivery and return list PDFs, organize them, and handle merging and linking.
- **getOrCreateFolder(folderName, parentFolder)**: Utility function to create or retrieve a folder in Google Drive.
- **exportSheetToPDF(sheet)**: Converts a Google Sheet into a PDF blob.
- **mergePDFsUsingCloudConvert(pdfUrls, dateFolder, orderIdsString)**: Merges multiple PDF files into one using the CloudConvert API.
- **createCloudConvertJob(apiKey, jobPayload)**: Creates a job in CloudConvert for processing files.
- **waitForCloudConvertJob(apiKey, jobId, exportTaskName)**: Polls CloudConvert for job completion and retrieves the download URL.
- **downloadAndSaveFile(downloadUrl, dateFolder, orderIdsString)**: Downloads and saves a merged PDF to Google Drive.
- **generateDirectDownloadUrl(fileId)**: Generates a direct download link for a Google Drive file.
- **setLinkInSheet(url, cellRef)**: Inserts a hyperlink in the Google Sheet.

## Usage

1. **Data Input**: Populate the `Daily Deliveries` sheet with delivery and return information.
2. **Run the Script**: Execute the `generateDeliveryPDFs` function from the Apps Script editor or set up a trigger to run it automatically.
3. **Retrieve PDFs**: Access the generated PDFs from Google Drive or directly from the links in the Google Sheet.
