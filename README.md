# 📧 Outlook Email to PDF Automator

Automated Python script for converting unread Outlook emails and their attachments into consolidated PDF documents with intelligent processing and organization.

## 🎯 Overview

This script automates the workflow of processing unread emails from a shared Outlook mailbox, converting email content and attachments into consolidated PDF files [file:12]. It handles multiple file formats, embedded content, and nested email attachments (.eml, .msg) to create comprehensive document packages for archival and compliance purposes.

## ✨ Features

- **Automatic Email Processing**: Monitors and processes unread emails from shared mailboxes
- **Comprehensive PDF Conversion**: Converts email body (HTML/text) to PDF format
- **Multi-Format Support**: Handles PDF, images (JPG, PNG, GIF, BMP), .eml, and .msg attachments
- **Embedded Content Extraction**: Extracts and processes embedded images using Content-ID (CID) references
- **Nested Email Processing**: Recursively processes forwarded emails (.eml and .msg attachments)
- **Smart File Consolidation**: Merges email content and all attachments into a single PDF document
- **Automatic Organization**: Creates structured folder hierarchy based on email subject and date
- **Inline Image Filtering**: Intelligently ignores signature images and decorative content
- **Detailed Logging**: Comprehensive console output for monitoring and troubleshooting
- **Auto-Read Marking**: Automatically marks processed emails as read

## 🚀 How It Works

The script follows a sequential workflow:

1. **Connection Phase**: Connects to Outlook via COM interface and accesses the configured shared mailbox
2. **Scanning Phase**: Retrieves all unread emails using MAPI filtering
3. **Folder Creation**: Creates organized output folders using email subject and timestamp
4. **Content Extraction**: Converts email HTML body to PDF, handling embedded images
5. **Attachment Processing**:
   - Saves all attachments with normalized filenames
   - Converts images to PDF format
   - Processes nested emails (.eml/.msg) recursively
   - Validates PDF integrity
   - Filters out inline signature images
6. **Consolidation Phase**: Merges email PDF and all attachment PDFs into a single document
7. **Cleanup**: Removes individual PDFs, keeping only the consolidated version
8. **Completion**: Marks email as read and logs processing summary

## 📋 Prerequisites

### System Requirements

- **Windows OS**: Required for COM interface with Outlook
- **Microsoft Outlook**: Must be installed and configured with access to the target mailbox
- **Python 3.8+**: With pip package manager

### Required Python Libraries

```bash
pip install pywin32           # Outlook COM automation
pip install PyPDF2            # PDF manipulation
pip install weasyprint        # HTML to PDF conversion
pip install beautifulsoup4    # HTML parsing
pip install Pillow            # Image processing
