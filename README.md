# Automated Loan Agreement Document Generator using Syncfusion DocIO

This repository demonstrates how to build a complete **Automated Loan Agreement Document Generator** using **Syncfusion DocIO**, **Syncfusion PDF**, and **Syncfusion Smart Data Extractor** in an ASP.NET Core MVC application. The sample showcases real-world financial document automation workflows including dynamic data merging, PDF conversion, and optional digital signature placement based on keyword detection.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── ErrorViewModel.cs
├── Views/
│   ├── Home/
│   │   ├── Index.cshtml
│   │   └── Privacy.cshtml
│   └── Shared/
├── wwwroot/
│   ├── Data/
│   │   ├── Template.docx
│   │   └── LoanAgreement.json
│   └── Signature.png
└── README.md
```
---

## 🚀 Getting Started

### Prerequisites

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Automated-Loan-Agreement
cd Automated-Loan-Agreement
```

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```
---

## 📋 How to Use

### Basic Workflow

1. **Upload Document Template (Optional)**
   - Upload a Word document (.docx, .doc, .rtf) with mail merge fields
   - Or use the default `Template.docx` provided
   - Drag & drop supported

2. **Upload Loan Data (Optional)**
   - Upload a JSON file with borrower and loan information
   - Or use the default `LoanAgreement.json` provided

3. **Select Output Type**
   - **Single PDF Document** – One merged document for all records
   - **Multiple PDF Files** – Separate PDFs per record (downloaded as ZIP)

4. **Enable Digital Signature (Optional)**
   - Check "Enable Digital Signature" to add signatures
   - Upload custom signature image or use default
   - Specify keywords for automatic signature placement (e.g., "Signatory, AuthorizedSign")

5. **Generate Document**
   - Click "Generate Document" button
   - Download automatically starts
   - File format: `GeneratedDocument.pdf` or `converted_pdfs.zip`

---

## 🎯 Use Cases

### Financial Scenarios

- **Loan Agreement Generation** – Create standardized loan contracts with borrower details and terms
- **Payment Schedule Reports** – Generate formatted repayment schedules with multiple installments
- **Batch Document Processing** – Process multiple loan agreements simultaneously for efficiency
- **Digital Document Signing** – Add secure digital signatures with automated keyword-based placement

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Simple Mail Merge](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/simple-mail-merge)
- [Mail Merge for Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-group)
- [Mail Merge for Nested Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-nested-groups)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)
- [Digital Signatures in PDF](https://help.syncfusion.com/file-formats/pdf/working-with-digitalsignature)

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can streamline financial document workflows in banking and lending environments.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific financial requirements.

You can modify the templates, data sources, processing logic, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---

## ⚠️ Limitations

- JSON property names must match Word template merge field names for successful data binding
- For repeating JSON arrays, the template group name must match the array element name (e.g., if `"Payments": [...]` contains payment objects, use `«GroupStart:Payments»` in template)
- Digital signatures require valid X.509 certificates for legal validity
- Signature keywords must exist in the document for automatic placement; documents without keywords will be generated without signatures

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---