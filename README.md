# Automated Loan Agreement Document Generator using Syncfusion DocIO

This repository demonstrates how to build a complete **Automated Loan Agreement Document Generator** using **Syncfusion DocIO**, **Syncfusion PDF**, and **Syncfusion Smart Data Extractor** in an ASP.NET Core MVC application. The sample showcases real-world financial document automation workflows including dynamic data merging, PDF conversion, and optional digital signature placement based on keyword detection.

## Overview

Financial institutions process hundreds of loan agreements daily, making manual document creation time-consuming and error-prone. This sample demonstrates how to:

- Automate loan agreement generation by merging JSON data into Word templates
- Handle simple and complex data structures with intelligent mail merge strategies
- Convert documents to professional PDF format
- Add secure digital signatures with automated placement
- Generate single documents or batch process multiple agreements
- Split documents by page breaks for individual distribution

The application leverages **Syncfusion DocIO** for Word document manipulation, **Syncfusion PDF** for document conversion and digital signing, and **Syncfusion Smart Data Extractor** for intelligent signature positioning.

## Key Features

### 📄 Dynamic Document Generation

- **Template-Based Approach**: Upload custom Word templates with mail merge fields for complete layout control
- **JSON Data Integration**: Merge loan details, borrower information, and payment schedules from JSON files
- **Smart Structure Detection**: Automatically detects JSON structure (flat, array, or nested) and applies appropriate merge strategy
- **No Code Changes Required**: Update templates and data without modifying application code

### 🔄 Intelligent Mail Merge Strategies

- **Simple/Flat Merge**: For basic key-value JSON data (e.g., borrower name, loan amount, interest rate)
- **Group Merge**: For JSON with arrays (e.g., multiple repayment schedules)
- **Nested Group Merge**: For hierarchical data structures (e.g., branches → loans → payments)
- **Automatic Detection**: System analyzes both template and JSON to choose optimal merge method

### 📑 Flexible Output Options

- **Single PDF Generation**: Create one merged document for individual loan agreements
- **Multiple PDF Generation**: Split documents by page breaks and bundle as ZIP for batch processing
- **Instant Download**: Generated documents delivered immediately via browser

### ✍️ Digital Signature Capability

- **Optional Feature**: Enable digital signatures via checkbox - completely optional
- **Custom or Default Signatures**: Upload your own signature image or use project defaults
- **Keyword-Based Placement**: Specify keywords (e.g., "Sign", "WITNESS", "AuthorizedSign") for automatic positioning
- **Keyword Detection**: Uses Syncfusion Smart Data Extractor to locate signature positions intelligently
- **Certificate Authentication**: Employs X.509 certificates (PDF.pfx) for document integrity and legal validity
- **Smart Fallback**: If keywords not found, document generated without signatures to prevent errors

### 🔍 Smart Data Extraction

- **Intelligent Text Processing**: AI-powered text extraction from PDF documents
- **Keyword Detection**: Automatically locates signature placeholder keywords with high accuracy
- **Precise Positioning**: Places signatures above detected keywords with proper spacing

## Project pre-requisites
Make sure that you have the compatible versions of Visual Studio and .NET SDK version in your machine before starting to work on this project.

## Syncfusion&reg; .NET Word Library
The Syncfusion&reg; DocIO is a [.NET Word library](https://www.syncfusion.com/document-processing/word-framework/net/word-library?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) allows you to add advanced Word document processing functionalities to any .NET application and does not require Microsoft Word application to be installed in the machine. It is a non-UI component that provides a full-fledged document instance model similar to the Microsoft Office COM libraries to iterate with the document elements explicitly and perform necessary manipulation. 

Take a moment to peruse the [documentation](https://help.syncfusion.com/file-formats/docio/getting-started?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), where you can find basic Word document processing options along with the features like [mail merge](https://help.syncfusion.com/file-formats/docio/working-with-mail-merge?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), [merge](https://help.syncfusion.com/file-formats/docio/word-document/merging-word-documents?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), and [compare documents](https://help.syncfusion.com/file-formats/docio/word-document/compare-word-documents?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), [find and replace](https://help.syncfusion.com/file-formats/docio/working-with-find-and-replace?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) text in the Word document, [protect](https://help.syncfusion.com/file-formats/docio/working-with-security?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) the Word documents, and most importantly, the [PDF](https://help.syncfusion.com/file-formats/docio/word-to-pdf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) and [Image](https://help.syncfusion.com/file-formats/docio/word-to-image?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) conversions with code examples.

Syncfusion&reg; .NET PDF library 
-------------------------------------

The Syncfusion&reg; [.NET PDF library](https://www.syncfusion.com/document-processing/pdf-framework/net) (Essential&reg; PDF) allows you to create, read and write PDF documents programatically in any .NET application. This high-performance and feature-rich .NET PDF framework works without Adobe dependencies. The creation of PDF follows the most popular PDF 1.7 (ISO 32000-1) and latest 2.0 (ISO 32000-2) specifications.

Compatible Microsoft Word Versions
----------------------------------

*   Microsoft Word 97-2003
*   Microsoft Word 2007
*   Microsoft Word 2010
*   Microsoft Word 2013
*   Microsoft Word 2016
*   Microsoft Word 2019
*   Microsoft 365

Supported File Formats
----------------------

*   Creates, reads, and edits popular text file formats like [DOC](https://help.syncfusion.com/file-formats/docio/word-file-formats?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples#doc-to-docx-and-docx-to-doc), DOT, [DOCM](https://help.syncfusion.com/file-formats/docio/word-file-formats?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples#macros), DOTM, [DOCX](https://help.syncfusion.com/file-formats/docio/word-file-formats?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples#doc-to-docx-and-docx-to-doc), [DOTX](https://help.syncfusion.com/file-formats/docio/word-file-formats?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples#templates), [HTML](https://help.syncfusion.com/file-formats/docio/html?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), [RTF](https://help.syncfusion.com/file-formats/docio/rtf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), [TXT](https://help.syncfusion.com/file-formats/docio/text?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), and [XML (WordML)](https://help.syncfusion.com/file-formats/docio/word-file-formats#word-processing-xml-xml?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples).
*   Converts Word documents also to [PDF](https://help.syncfusion.com/file-formats/docio/word-to-pdf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), [Image](https://help.syncfusion.com/file-formats/docio/word-to-image?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples), and [ODT](https://help.syncfusion.com/file-formats/docio/word-to-odt?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) files.

## How to Run the Examples

1. Clone this repository
2. Open the solution file in Visual Studio 2022
3. Register your Syncfusion license key in the sample
4. Restore NuGet packages
5. Build and run the project

## How to Use

### Basic Document Generation

1. **Upload Template**: Select a Word document (.docx) with merge fields like `«BorrowerName»`, `«LoanAmount»`
2. **Upload Data**: Select a JSON file with corresponding field values
3. **Select Output Type**: 
   - `Single` - One merged PDF
   - `Multiple` - Separate PDFs in ZIP archive
4. **Generate**: Click "Generate Document" to download

### Adding Digital Signatures

1. **Enable Signatures**: Check "Enable Digital Signature" checkbox
2. **Upload Signature Image** (optional): Choose a custom signature (PNG, JPG, etc.) or use default
3. **Specify Keywords** (optional): Enter comma-separated keywords where signatures should appear
   - Default: "Sign, WITNESS, AuthorizedSign, Signature"
4. **Generate**: Signatures will be automatically placed above detected keywords in the PDF

### Using Default Files

If you don't upload custom files, the application uses:
- `wwwroot/Data/Template.docx` - Default loan agreement template
- `wwwroot/Data/LoanAgreement.json` - Sample borrower data
- `wwwroot/Signature.png` - Default signature image (when enabled)

## Resources

- **Product page:** [Syncfusion&reg; Word Framework](https://www.syncfusion.com/document-processing/word-framework/net?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples)
- **Documentation:** [Getting Started with Word Library](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- **GitHub Examples:** [Syncfusion&reg; Word library examples](https://github.com/SyncfusionExamples/DocIO-Examples?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples)
- **Online demo:** [Syncfusion&reg; Word library - Online demos](https://ej2.syncfusion.com/aspnetcore/Word/SalesInvoice#/material3?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples)

## Support and feedback
- For any other queries, reach our [Syncfusion&reg; support team](https://support.syncfusion.com/agent/tickets/create?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples).
- Request new feature through [Syncfusion&reg; feedback portal](https://www.syncfusion.com/feedback/home?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples).

## License
This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/22.2.5/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-docio-video-examples).