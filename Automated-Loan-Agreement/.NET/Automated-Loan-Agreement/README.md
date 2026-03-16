# Syncfusion ASP.NET Core – Automated Loan Agreement Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Loan Agreement Document Generator** using **Syncfusion DocIO** and **Syncfusion PDF** libraries in an ASP.NET Core MVC application. The sample illustrates how financial teams can automate loan agreement generation by merging dynamic JSON data into Word templates and exporting them as PDF documents.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── ErrorViewModel.cs
├── Views/
│   ├── Home/
│   │   └── Index.cshtml
│   └── Shared/
├── wwwroot/
└── README.md
```

---

## ✨ Features

- Upload a **Word template** (.docx) with mail merge fields as placeholders.
- Upload a **JSON data file** containing loan applicant or agreement details.
- Automatically detect JSON structure and apply the appropriate mail merge strategy:
  - **Simple/Flat Merge** – for plain key-value JSON data.
  - **Group Merge** – for JSON containing arrays of records.
  - **Nested Group Merge** – for hierarchical JSON with nested arrays.
- Generate a **single merged PDF** for a single loan agreement.
- Generate **multiple PDFs** (split by page breaks) bundled as a **ZIP file** for batch loan agreements.
- Demonstrates how the Syncfusion DocIO library is used to automate document generation.

## 🚀 Getting Started

### Prerequisites

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

---

### 1. Clone the Repository

```bash
git clone https://github.com/your-repo/Automated-Loan-Agreement.git
```

---

### 2. Navigate to the Project Directory

```bash
cd Automated-Loan-Agreement
```

---

### 3. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

---

### 4. Add Syncfusion License Key

In your `Program.cs` or `Startup.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```
---

### 5. Run the Application

```bash
dotnet run
```

---

### 6. Open in Browser

Navigate to the localhost URL shown in the terminal output, for example:

```
https://localhost:5001
```

---

## 📋 How to Use

1. Open the application in your browser.
2. Upload a **Word template** file (`.docx`) that contains mail merge fields (e.g., `«BorrowerName»`, `«LoanAmount»`).
3. Upload a **JSON data file** with the corresponding field values.
4. Select the **Output Type**:
   - `single` – Downloads a single merged PDF.
   - `multiple` – Downloads a ZIP file containing separate PDFs split by page breaks.
5. Click **Generate Document** to download the result.

---

## 📄 JSON Data Format Examples

### Flat / Simple JSON (Single Loan Agreement)

```json
{
  "BorrowerName": "John Smith",
  "LoanAmount": "50000",
  "InterestRate": "7.5%",
  "LoanTerm": "5 Years",
  "StartDate": "01/04/2026"
}
```

### Group JSON (Multiple Records)

```json
{
  "BorrowerName": "John Smith",
  "LoanAmount": "50000",
  "Repayments": [
    { "Month": "April 2026", "Amount": "1000", "Balance": "49000" },
    { "Month": "May 2026", "Amount": "1000", "Balance": "48000" }
  ]
}
```

### Nested Group JSON (Hierarchical Records)

```json
{
  "Branches": [
    {
      "BranchName": "New York",
      "Loans": [
        {
          "BorrowerName": "Alice",
          "LoanAmount": "30000",
          "Repayments": [
            { "Month": "April 2026", "Amount": "600" }
          ]
        }
      ]
    }
  ]
}
```

---

## 🔗 Resources

- https://help.syncfusion.com/file-formats/docio/getting-started
- https://help.syncfusion.com/file-formats/docio/mail-merge/simple-mail-merge
- https://help.syncfusion.com/file-formats/docio/mail-merge/mail-merge-for-group
- https://help.syncfusion.com/file-formats/docio/mail-merge/mail-merge-for-nested-groups
- https://help.syncfusion.com/file-formats/pdf/getting-started

---

## ✅ Benefits

- Eliminates manual document creation for loan agreements.
- Supports both individual and batch document generation.
- Handles complex hierarchical data structures with nested mail merge.
- Produces professional, print-ready PDF documents.
- Easily customizable Word templates - no code changes needed for layout updates.

---

## 📣 Try It Out

Clone the repository, run the sample, and explore how **Syncfusion DocIO** can be used to build a complete **Automated Document Generation** system for real-world workflows.

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup here[https://www.syncfusion.com/account/downloads/studio]. If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads). 


