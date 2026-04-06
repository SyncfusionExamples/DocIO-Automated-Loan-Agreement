using Automated_Loan_Agreement.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using Syncfusion.SmartDataExtractor;
using System.Diagnostics;
using System.Dynamic;
using System.IO.Compression;

namespace Automated_Loan_Agreement.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult GenerateDocument(IFormFile file, IFormFile jsonFile, IFormFile signatureImage, string signatureKeywords, string outputType)
        {
            try
            {
                // Load Word document stream (uploaded or default)
                Stream wordStream = GetWordDocument(file);

                if (wordStream == null)
                    return View("Index");

                // Load JSON stream - only load default if Word is ALSO using default
                Stream jsonStream = GetJsonDocument(file, jsonFile);

                if (jsonStream == null)
                    return View("Index");

                bool isUsingDefaults = (file == null || file.Length == 0) && (jsonFile == null || jsonFile.Length == 0);

                return CreatePDF(wordStream, jsonStream, outputType, signatureImage, signatureKeywords, isUsingDefaults);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating document");
                ViewBag.Message = $"Error: {ex.Message}";
                return View("Index");
            }
        }
        /// <summary>
        /// Loads Word document, performs mail merge, and generates output PDF
        /// </summary>
        private IActionResult CreatePDF(Stream stream, Stream jsonStream, string type, IFormFile signatureImage, string signatureKeywords, bool isUsingDefaults)
        {
            // Validate type parameter to avoid NullReferenceException
            if (string.IsNullOrWhiteSpace(type))
            {
                ViewBag.Message = "Document type parameter is required.";
                return View("Index");
            }
            // Load the Word document with automatic format detection
            using (WordDocument document = new WordDocument(stream, FormatType.Automatic))
            {
                // Proceed with mail merge only if a valid JSON file is provided
                if (jsonStream != null && jsonStream.Length > 0)
                {
                    using (StreamReader reader = new StreamReader(jsonStream))
                    {
                        string jsonData = reader.ReadToEnd();

                        // Parse the JSON safely and handle malformed JSON
                        JObject jsonObject;
                        try
                        {
                            jsonObject = JObject.Parse(jsonData);
                        }
                        catch (JsonReaderException ex)
                        {
                            ViewBag.Message = $"Invalid JSON format: {ex.Message}";
                            return View("Index");
                        }

                        // Analyze the JSON structure to determine merge strategy
                        bool hasOnlySimpleFields = true;
                        bool hasGroups = false;

                        foreach (var property in jsonObject.Properties())
                        {
                            if (property.Value is JArray || property.Value is JObject)
                            {
                                // Found an array OR nested object — indicates group/nested merge needed
                                hasGroups = true;
                                hasOnlySimpleFields = false;
                            }
                        }
                        bool templateHasGroupFields = HasGroupMergeFields(document);
                        // SCENARIO 1: JSON contains only simple key-value pairs
                        // Use Execute() for flat mail merge
                        if (hasOnlySimpleFields || (hasGroups && !templateHasGroupFields))
                        {
                            // Get the array property
                            var arrayProperty = jsonObject.Properties()
                                .FirstOrDefault(p => p.Value is JArray);

                            if (arrayProperty != null && arrayProperty.Value is JArray jsonArray && jsonArray.Count > 0)
                            {
                                List<object> Data = ((JArray)arrayProperty.Value).ToObject<List<object>>();
                                document.MailMerge.Execute(Data);
                            }
                            else
                            {
                                // No array found - process as simple fields only
                                List<string> fieldNames = new List<string>();
                                List<string> fieldValues = new List<string>();
                                foreach (var property in jsonObject.Properties())
                                {
                                    if (!(property.Value is JArray) && !(property.Value is JObject))
                                    {
                                        fieldNames.Add(property.Name);
                                        fieldValues.Add(property.Value?.ToString() ?? string.Empty);
                                    }
                                }
                                if (fieldNames.Count > 0)
                                {
                                    document.MailMerge.Execute(fieldNames.ToArray(), fieldValues.ToArray());
                                }
                            }
                        }
                        // SCENARIO 2: JSON contains arrays or mixed content
                        // Use ExecuteGroup() or ExecuteNestedGroup() accordingly
                        else if ((hasGroups && templateHasGroupFields))
                        {
                            // Step 1: Extract and merge simple (non-array, non-object) fields first
                            List<string> simpleFieldNames = new List<string>();
                            List<string> simpleFieldValues = new List<string>();
                            foreach (var property in jsonObject.Properties())
                            {
                                if (!(property.Value is JArray) && !(property.Value is JObject))
                                {
                                    simpleFieldNames.Add(property.Name);
                                    simpleFieldValues.Add(property.Value?.ToString() ?? string.Empty);
                                }
                            }
                            // Execute simple fields merge only if any exist
                            if (simpleFieldNames.Count > 0)
                            {
                                document.MailMerge.Execute(simpleFieldNames.ToArray(), simpleFieldValues.ToArray());
                            }
                            // Step 2: Set each record to start on a new page (only for group merge)
                            document.MailMerge.StartAtNewPage = true;
                            // Step 3: Process each array/group property for group mail merge
                            foreach (var property in jsonObject.Properties())
                            {
                                string groupName = property.Name;
                                // Handle JObject (nested object) as single-record group
                                if (property.Value is JObject nestedObj)
                                {
                                    List<ExpandoObject> singleRecordList = new List<ExpandoObject>();
                                    dynamic record = new ExpandoObject();
                                    var recordDict = (IDictionary<string, object>)record;

                                    // Extract nested object's properties directly
                                    foreach (var nestedProp in nestedObj.Properties())
                                    {
                                        if (!(nestedProp.Value is JArray) && !(nestedProp.Value is JObject))
                                        {
                                            recordDict[nestedProp.Name] = nestedProp.Value?.ToString() ?? string.Empty;
                                        }
                                    }

                                    singleRecordList.Add(record);
                                    MailMergeDataTable dataTable = new MailMergeDataTable(groupName, singleRecordList);
                                    document.MailMerge.ExecuteGroup(dataTable);
                                }
                                else if (property.Value is JArray jsonArray && jsonArray.Count > 0)
                                {
                                    // Check whether the array contains nested groups (arrays within arrays)
                                    bool hasNestedGroups = CheckForNestedGroups(jsonArray);

                                    if (hasNestedGroups)
                                    {
                                        // Use ExecuteNestedGroup for hierarchical/nested data
                                        List<dynamic> parentDataList = ConvertToNestedDataList(jsonArray);
                                        MailMergeDataTable parentTable = new MailMergeDataTable(groupName, parentDataList);
                                        document.MailMerge.ExecuteNestedGroup(parentTable);
                                    }
                                    else
                                    {
                                        // Use ExecuteGroup for flat array data
                                        List<ExpandoObject> dataList = ConvertToFlatDataList(jsonArray);
                                        MailMergeDataTable dataTable = new MailMergeDataTable(groupName, dataList);
                                        document.MailMerge.ExecuteGroup(dataTable);
                                    }
                                }
                            }
                        }
                    }
                }

                // OUTPUT: Single merged PDF with signatures
                if (string.Equals(type, "single", StringComparison.OrdinalIgnoreCase))
                {
                    //Adding Signature
                    MemoryStream pdfStream = DataExtractor(SaveAsPDF(document), signatureImage, signatureKeywords, isUsingDefaults);
                    // Return the PDF as a downloadable file
                    return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
                }
                // OUTPUT: Split document by page breaks with signatures in each PDF
                else if (string.Equals(type, "multiple", StringComparison.OrdinalIgnoreCase))
                {
                    byte[] zipBytes = SplitByPageBreak(document, signatureImage, signatureKeywords, isUsingDefaults);

                    if (zipBytes != null && zipBytes.Length > 0)
                    {
                        // Return all split PDFs bundled inside a ZIP file
                        return File(zipBytes, "application/zip", "converted_pdfs.zip");
                    }
                    else
                    {
                        MemoryStream pdfStream = DataExtractor(SaveAsPDF(document), signatureImage, signatureKeywords, isUsingDefaults);
                        // Return the PDF as a downloadable file
                        return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
                    }
                }
                else
                {
                    // Handle any unrecognized type value
                    ViewBag.Message = $"Unknown document type '{type}'. Use 'single' or 'multiple'.";
                }

                return View("Index");
            }

        }
        /// <summary>
        /// Checks if the Word document has group merge regions
        /// </summary>

        private bool HasGroupMergeFields(WordDocument document)
        {
            string[] groupNames = document.MailMerge.GetMergeGroupNames();
            return groupNames != null && groupNames.Length > 0;

        }
        /// <summary>
        /// Extracts data from an input PDF stream and conditionally applies digital signatures
        /// at detected keyword locations using a provided signature image.
        /// </summary>

        private MemoryStream DataExtractor(Stream inputStream, IFormFile signatureImage, string signatureKeywordsInput, bool isUsingDefaults)
        {

            // Initialize the data extractor with table detection enabled
            // and form field detection explicitly disabled
            var extractor = new DataExtractor { EnableFormDetection = false, EnableTableDetection = true, ConfidenceThreshold = 0.6 };
            // Extract PDF data and load it into a PDF document
            PdfLoadedDocument document = extractor.ExtractDataAsPdfDocument(inputStream);

            // Get image path only from user-provided image
            string imagePath = GetSignatureImagePath(signatureImage, isUsingDefaults);

            //If no image provided, return PDF without signature processing
            if (string.IsNullOrEmpty(imagePath) || !System.IO.File.Exists(imagePath))
            {
                _logger.LogInformation("No signature image provided. Returning PDF without digital signatures.");
                var outMs = new MemoryStream();
                document.Save(outMs);
                document.Close(true);
                outMs.Position = 0;
                return outMs;
            }

            // ONLY process signatures if image exists
            string[] signatureKeywords = string.IsNullOrWhiteSpace(signatureKeywordsInput)
                ? new[] { "Signature", "WITNESS", "AuthorizedSign" }
                : signatureKeywordsInput.Split(',').Select(k => k.Trim()).ToArray();
            // Iterate through each page in the PDF document
            for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++)
            {
                PdfPageBase page = document.Pages[pageIndex];
                TextLineCollection textLines;
                // Extract text lines and words from the current page
                page.ExtractText(out textLines);
                // Loop through all extracted text lines
                foreach (TextLine line in textLines.TextLine)
                {
                    // Loop through each word within the line
                    foreach (TextWord word in line.WordCollection)
                    {
                        if (!signatureKeywords.Any(k => word.Text.Contains(k, StringComparison.Ordinal)))
                            continue;
                        // Determine the bounding rectangle of the keyword
                        RectangleF bounds = word.Bounds;
                        // Calculate signature placement relative to keyword position
                        float signatureX = bounds.X;
                        float signatureY = bounds.Y - bounds.Height - 5;
                        float signatureWidth = 80;
                        float signatureHeight = 20;

                        try
                        {
                            using System.IO.FileStream cert = new System.IO.FileStream(Path.GetFullPath("PDF.pfx"), System.IO.FileMode.Open, System.IO.FileAccess.Read);
                            PdfCertificate pdfCert = new PdfCertificate(cert, "syncfusion");
                            // Create a digital signature object for the document
                            PdfSignature signature = new PdfSignature(document, page, pdfCert, "Signature");
                            // Set signature position and siz
                            signature.Bounds = new RectangleF(signatureX, signatureY, signatureWidth, signatureHeight);

                            // Add appearance image
                            if (System.IO.File.Exists(imagePath))
                            {
                                using System.IO.FileStream img = new System.IO.FileStream(imagePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                                PdfBitmap signatureImageBitmap = new PdfBitmap(img);
                                // Draw the image into the signature appearance
                                signature.Appearance.Normal.Graphics.DrawImage(signatureImageBitmap, 0, 0, signatureWidth, signatureHeight);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Error adding signature to PDF");
                        }
                    }
                }
            }
            // Save the modified document to a memory stream
            var outputMs = new MemoryStream();
            document.Save(outputMs);
            document.Close(true);
            outputMs.Position = 0;
            return outputMs;
        }

        /// <summary>
        /// Gets signature image path only from user-provided image
        /// Returns null if no image provided - no fallback to local files
        /// </summary>
        private string GetSignatureImagePath(IFormFile signatureImage, bool isUsingDefaults)
        {
            if (signatureImage != null && signatureImage.Length > 0)
            {
                // Save user-provided image to temporary location
                return SaveTemporaryImage(signatureImage);
            }

            // If using default documents, use default local image
            if (isUsingDefaults)
            {
                string defaultImagePath = Path.GetFullPath("signature.png");
                if (System.IO.File.Exists(defaultImagePath))
                {
                    _logger.LogInformation("Using default signature image from local.");
                    return defaultImagePath;
                }
            }

            // No image provided and not using defaults - return null
            _logger.LogInformation("No signature image provided.");
            return null;
        }

        /// <summary>
        /// Saves user-provided signature image to temporary file
        /// </summary>
        private string SaveTemporaryImage(IFormFile imageFile)
        {
            try
            {
                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + System.IO.Path.GetExtension(imageFile.FileName));
                using (var stream = new System.IO.FileStream(tempPath, System.IO.FileMode.Create))
                {
                    imageFile.CopyTo(stream);
                }
                return tempPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving temporary signature image");
                return null;
            }
        }
        /// <summary>
        /// Retrieves a Word document stream from the uploaded file or a default template.
        /// </summary>

        private Stream GetWordDocument(IFormFile file)
        {
            // Case 1: Uploaded file exists and has content
            if (file != null && file.Length > 0)
            {
                string extension = Path.GetExtension(file.FileName).ToLower();
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".docm", ".xml", ".rtf" };
                // Validate the file extension
                if (supportedExtensions.Contains(extension))
                {
                    // Copy the uploaded file into an in-memory stream
                    MemoryStream stream = new MemoryStream();
                    file.CopyTo(stream);
                    // Reset stream position to the beginning for downstream reading
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please choose a Word format document to convert to PDF.";
                    return null;
                }
            }
            else
            {
                // Load default file from wwwroot\Data\
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Template.docx");
                return new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }
        /// <summary>
        /// Retrieves an Json stream based on the uploaded files.
        /// </summary>
        private Stream GetJsonDocument(IFormFile file, IFormFile jsonFile)
        {
            // If Word file was uploaded, JSON file must also be uploaded
            if (file != null && file.Length > 0)
            {
                // Ensure an Json file is also uploaded
                if (jsonFile != null && jsonFile.Length > 0)
                {
                    // Copy uploaded Json file into an in-memory stream.
                    MemoryStream stream = new MemoryStream();
                    jsonFile.CopyTo(stream);
                    // Reset stream position so it can be read from the beginning
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please upload a JSON data file along with the Word document.";
                    return null;
                }
            }
            else
            {
                // Both Word and JSON are defaults (no file uploaded)
                string defaultJsonPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "LoanAgreement.json");
                return new FileStream(defaultJsonPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }
        /// <summary>
        /// Splits document by page breaks using bookmarks and returns ZIP bytes
        /// Each section between page breaks becomes a separate PDF
        /// </summary>
        private byte[] SplitByPageBreak(WordDocument document, IFormFile signatureImage, string signatureKeywords, bool isUsingDefaults)
        {
            // Find all page breaks in the document
            List<Entity> entities = document.FindAllItemsByProperty(EntityType.Break, "BreakType", "PageBreak");
            if (entities == null || entities.Count == 0)
                return null;

            WSection section = document.Sections[0];
            WTextBody body = section.Body;
            int bookmarkIndex = 1;
            // Step 1: Insert a NEW paragraph at the very beginning with BookmarkStart
            WParagraph firstBookmarkPara = new WParagraph(document);
            firstBookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Insert(0, firstBookmarkPara);

            // Step 2: Iterate page break entities → insert bookmark paragraph directly after each
            foreach (Entity entity in entities)
            {
                WParagraph breakParagraph = entity.Owner as WParagraph;

                if (breakParagraph == null) continue;

                // Get the current index of this paragraph in the body
                int paraIndex = body.ChildEntities.IndexOf(breakParagraph);

                if (paraIndex < 0) continue;

                // Insert new paragraph right after the page break paragraph
                // Close current bookmark and open next bookmark in same paragraph
                WParagraph bookmarkPara = new WParagraph(document);
                bookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
                bookmarkIndex++;
                bookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
                body.ChildEntities.Insert(paraIndex + 1, bookmarkPara);
            }

            // Step 3: Insert a NEW paragraph at the very end with BookmarkEnd
            WParagraph lastBookmarkPara = new WParagraph(document);
            lastBookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Add(lastBookmarkPara);
            // Step 4: Create ZIP file and convert each bookmarked section to PDF
            var zipStream = new MemoryStream();
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 1; i <= bookmarkIndex; i++)
                {
                    try
                    {
                        // Navigate to each bookmark section
                        BookmarksNavigator navigator = new BookmarksNavigator(document);
                        navigator.MoveToBookmark($"Page_Bookmark_{i}", true, true);
                        WordDocumentPart documentPart = navigator.GetContent();

                        if (documentPart == null) continue;

                        // Extract content as new WordDocument.
                        using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                        using (DocIORenderer render = new DocIORenderer())
                        // Convert extracted document into PDF.
                        using (PdfDocument pdfDocument = render.ConvertToPDF(extractedDoc))
                        using (MemoryStream pdfStream = new MemoryStream())
                        {
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;

                            // Apply digital signatures to each PDF
                            MemoryStream signedPdfStream = DataExtractor(pdfStream, signatureImage, signatureKeywords, isUsingDefaults);

                            // Write PDF into ZIP entry
                            var entry = zip.CreateEntry($"Document_{i}.pdf", CompressionLevel.Fastest);
                            using (var entryStream = entry.Open())
                            {
                                signedPdfStream.CopyTo(entryStream);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue to next section if one fails
                        _logger.LogError(ex, $"Error converting bookmark {i} to PDF");
                    }
                }
            }
            return zipStream.ToArray();
        }
        /// <summary>
        /// Convert document into PDF and returns PDF stream
        /// </summary>
        private MemoryStream SaveAsPDF(WordDocument document)
        {
            using (DocIORenderer renderer = new DocIORenderer())
            {
                // Convert the Word document to PDF format
                using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                {
                    MemoryStream pdfStream = new MemoryStream();
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;

                    return pdfStream;
                }
            }
        }

        /// <summary>
        /// Checks if any item in the array contains a nested array (indicates nested groups)
        /// </summary>
        private bool CheckForNestedGroups(JArray jsonArray)
        {
            if (jsonArray.Count == 0) return false;

            var firstItem = jsonArray[0] as JObject;
            if (firstItem == null) return false;

            // If any property is an array, it's a nested group structure
            return firstItem.Properties().Any(p => p.Value is JArray);
        }

        /// <summary>
        /// Converts flat JSON array to list of ExpandoObjects (for simple group merge)
        /// Skips nested array properties
        /// </summary>
        private List<ExpandoObject> ConvertToFlatDataList(JArray jsonArray)
        {
            List<ExpandoObject> dataList = new List<ExpandoObject>();

            foreach (JObject item in jsonArray)
            {
                dynamic expando = new ExpandoObject();
                var expandoDict = expando as IDictionary<string, object>;

                foreach (var prop in item.Properties())
                {
                    if (!(prop.Value is JArray))
                    {
                        expandoDict[prop.Name] = prop.Value?.ToString() ?? string.Empty;
                    }
                }
                dataList.Add(expando);
            }
            return dataList;
        }

        /// <summary>
        /// Recursively converts nested JSON array to list of dynamic ExpandoObjects
        /// Handles any depth: Employees → Customers → Orders → etc.
        /// </summary>
        private List<dynamic> ConvertToNestedDataList(JArray jsonArray)
        {
            List<dynamic> dataList = new List<dynamic>();

            foreach (JObject item in jsonArray)
            {
                dynamic expando = new ExpandoObject();
                var expandoDict = expando as IDictionary<string, object>;

                foreach (var prop in item.Properties())
                {
                    if (prop.Value is JArray nestedArray)
                    {
                        // Recursive call: handles arrays at any depth
                        expandoDict[prop.Name] = ConvertToNestedDataList(nestedArray);
                    }
                    else if (prop.Value is JObject nestedObject)
                    {
                        // Convert nested object to ExpandoObject
                        dynamic nestedExpando = new ExpandoObject();
                        var nestedDict = nestedExpando as IDictionary<string, object>;

                        foreach (var nestedProp in nestedObject.Properties())
                        {
                            if (nestedProp.Value is JArray deepArray)
                            {
                                // Recursive call for arrays inside nested objects
                                nestedDict[nestedProp.Name] = ConvertToNestedDataList(deepArray);
                            }
                            else
                            {
                                nestedDict[nestedProp.Name] = nestedProp.Value?.ToString() ?? string.Empty;
                            }
                        }
                        expandoDict[prop.Name] = nestedExpando;
                    }
                    else
                    {
                        // Simple property value
                        expandoDict[prop.Name] = prop.Value?.ToString() ?? string.Empty;
                    }
                }
                dataList.Add(expando);
            }
            return dataList;
        }


        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
