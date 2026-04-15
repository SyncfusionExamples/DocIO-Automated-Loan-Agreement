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

        public IActionResult GenerateDocument(IFormFile file, IFormFile jsonFile, IFormFile signatureImage, string signatureKeywords, string outputType, bool enableDigitalSign)
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

                return CreatePDF(wordStream, jsonStream, outputType, signatureImage, signatureKeywords, enableDigitalSign);
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
        private IActionResult CreatePDF(Stream stream, Stream jsonStream, string type, IFormFile signatureImage, string signatureKeywords, bool enableDigitalSign)
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

                        // Parse as JToken to handle both JObject and JArray at root level
                        JToken jsonToken;
                        try
                        {
                            jsonToken = JToken.Parse(jsonData);
                        }
                        catch (JsonReaderException ex)
                        {
                            ViewBag.Message = $"Invalid JSON format: {ex.Message}";
                            return View("Index");
                        }

                        // Check if template has group merge fields
                        string[] templateGroups = HasGroupMergeFields(document);
                        bool templateHasGroupFields = templateGroups != null && templateGroups.Length > 0;

                        // Check if JSON contains groups (arrays or nested objects)
                        bool jsonHasGroups = JsonContainsGroups(jsonToken);

                        // Decide which mail merge method to use
                        // SCENARIO 1: JSON contains arrays or mixed content
                        // Use ExecuteGroup() or ExecuteNestedGroup() accordingly
                        if (templateHasGroupFields && jsonHasGroups)
                        {
                            ExecuteGroupMailMerge(document, jsonToken);
                        }
                        // SCENARIO 2: Simple fields OR groups without template group fields
                        else
                        {
                            ExecuteSimpleMailMerge(document, jsonToken);
                        }
                       
                    }
                }

                // OUTPUT: Single merged PDF with signatures
                if (string.Equals(type, "single", StringComparison.OrdinalIgnoreCase))
                {
                    //Adding Signature
                    MemoryStream pdfStream = DataExtractor(SaveAsPDF(document), signatureImage, signatureKeywords, enableDigitalSign);
                    // Return the PDF as a downloadable file
                    return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
                }
                // OUTPUT: Split document by page breaks with signatures in each PDF
                else if (string.Equals(type, "multiple", StringComparison.OrdinalIgnoreCase))
                {
                    byte[] zipBytes = SplitByPageBreak(document, signatureImage, signatureKeywords, enableDigitalSign);

                    if (zipBytes != null && zipBytes.Length > 0)
                    {
                        // Return all split PDFs bundled inside a ZIP file
                        return File(zipBytes, "application/zip", "converted_pdfs.zip");
                    }
                    else
                    {
                        MemoryStream pdfStream = DataExtractor(SaveAsPDF(document), signatureImage, signatureKeywords, enableDigitalSign);
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

        private string[] HasGroupMergeFields(WordDocument document)
        {
            string[] groupNames = document.MailMerge.GetMergeGroupNames();
            return groupNames;

        }
        /// <summary>
        /// Checks if JSON contains groups (arrays or nested objects)
        /// </summary>
        private bool JsonContainsGroups(JToken jsonToken)
        {
            if (jsonToken is JArray)
            {
                return true;
            }

            if (jsonToken is JObject rootObject)
            {
                foreach (var property in rootObject.Properties())
                {
                    if (property.Value is JArray || property.Value is JObject)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Executes a simple mail merge on a Word document using JSON data.
        /// Supports multiple JSON formats including:
        /// - Direct array of objects
        /// - Columns/Rows format
        /// - Wrapped array format
        /// - Single flat object
        /// </summary>
        private void ExecuteSimpleMailMerge(WordDocument document, JToken jsonToken)
        {
            // List to hold mail merge records
            List<ExpandoObject> records = new List<ExpandoObject>();

            // CASE 1: Root-level array
            if (jsonToken is JArray directArray)
            {
                records = ConvertDirectArrayToRecords(directArray);
            }
            // CASE 2: Root is an object
            else if (jsonToken is JObject parsedObject)
            {
                // CASE 2A: JSON in Columns/Rows format
                if (parsedObject.ContainsKey("Columns") && parsedObject.ContainsKey("Rows") &&
                    parsedObject["Columns"] is JArray columnsArray &&
                    parsedObject["Rows"] is JArray rowsArray)
                {
                    records = ConvertColumnsRowsToRecords(columnsArray, rowsArray);
                }
                // CASE 2B: Object that wraps an array property
                else
                {
                    var arrayProperty = parsedObject.Properties()
                        .FirstOrDefault(p => p.Value is JArray);

                    if (arrayProperty != null && arrayProperty.Value is JArray jsonArray && jsonArray.Count > 0)
                    {
                        records = ConvertWrappedArrayToRecords(jsonArray);
                    }
                    else
                    {
                        // CASE 2C: Single object (simple or nested JSON)
                        ExecuteSingleObjectMailMerge(document, parsedObject);
                        return;
                    }
                }
            }
            // Execute mail merge if records exist
            if (records.Count > 0)
            {
                document.MailMerge.Execute(records);
            }
        }

        /// <summary>
        /// Converts a direct JSON array to list of ExpandoObjects
        /// </summary>
        private List<ExpandoObject> ConvertDirectArrayToRecords(JArray directArray)
        {
            List<ExpandoObject> records = new List<ExpandoObject>();

            foreach (JObject item in directArray)
            {
                var flattenedFields = FlattenJson(item);
                if (flattenedFields.Count > 0)
                {
                    dynamic record = new ExpandoObject();
                    var recordDict = (IDictionary<string, object>)record;

                    foreach (var kvp in flattenedFields)
                    {
                        recordDict[kvp.Key] = kvp.Value;
                    }

                    records.Add(record);
                }
            }

            return records;
        }

        /// <summary>
        /// Converts Columns/Rows format JSON to list of ExpandoObjects
        /// </summary>
        private List<ExpandoObject> ConvertColumnsRowsToRecords(JArray columnsArray, JArray rowsArray)
        {
            List<ExpandoObject> records = new List<ExpandoObject>();
            List<string> columns = columnsArray.ToObject<List<string>>();

            foreach (JArray row in rowsArray)
            {
                if (row.Count == columns.Count)
                {
                    dynamic record = new ExpandoObject();
                    var recordDict = (IDictionary<string, object>)record;

                    for (int i = 0; i < columns.Count; i++)
                    {
                        recordDict[columns[i]] = row[i]?.ToString() ?? string.Empty;
                    }
                    records.Add(record);
                }
            }

            return records;
        }

        /// <summary>
        /// Converts wrapped array format JSON to list of ExpandoObjects
        /// Example: { "employees": [ { ... }, { ... } ] }
        /// </summary>
        private List<ExpandoObject> ConvertWrappedArrayToRecords(JArray jsonArray)
        {
            List<ExpandoObject> records = new List<ExpandoObject>();

            if (jsonArray.Count > 0 && jsonArray[0] is JObject)
            {
                foreach (JObject item in jsonArray)
                {
                    var flattenedFields = FlattenJson(item);
                    if (flattenedFields.Count > 0)
                    {
                        dynamic record = new ExpandoObject();
                        var recordDict = (IDictionary<string, object>)record;

                        foreach (var kvp in flattenedFields)
                        {
                            recordDict[kvp.Key] = kvp.Value;
                        }
                        records.Add(record);
                    }
                }
            }
            else
            {
                _logger.LogWarning("Array of primitives detected");
            }

            return records;
        }

        /// <summary>
        /// Executes mail merge for a single flat object
        /// Example: { "EmployeeId": "EMP001", "Name": "John Doe" }
        /// </summary>
        private void ExecuteSingleObjectMailMerge(WordDocument document, JObject parsedObject)
        {
            var flattenedFields = FlattenJson(parsedObject);
            if (flattenedFields.Count > 0)
            {
                document.MailMerge.Execute(flattenedFields.Keys.ToArray(), flattenedFields.Values.ToArray());
            }
        }
        /// <summary>
        /// Executes a mail merge operation on a Word document using data from a JSON token.
        /// Processes complex nested groups/arrays.
        /// </summary>
        private void ExecuteGroupMailMerge(WordDocument document, JToken jsonToken)
        {
            if (jsonToken is JObject groupJsonObject)
            {
                // Step 1: Extract simple fields from JSON
                Dictionary<string, string> simpleFields = ExtractSimpleFieldsFromJson(groupJsonObject);

                // Step 2: Process complex groups (nested objects and arrays)
                // Enable new page for each group execution
                document.MailMerge.StartAtNewPage = true;

                foreach (var property in groupJsonObject.Properties())
                {
                    string groupName = property.Name;
                    // Handle single nested object as a group
                    if (property.Value is JObject nestedObj)
                    {
                        List<ExpandoObject> singleRecordList = new List<ExpandoObject>();
                        dynamic record = new ExpandoObject();
                        var recordDict = (IDictionary<string, object>)record;
                        // Extract simple fields from nested object
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
                    // Handle array of records as a group
                    else if (property.Value is JArray jsonArray && jsonArray.Count > 0)
                    {
                        // Check if array contains nested groups (arrays or objects within objects)
                        bool hasNestedGroups = CheckForNestedGroups(jsonArray);

                        if (hasNestedGroups)
                        {
                            // Convert and execute as nested group mail merge
                            List<dynamic> parentDataList = ConvertToNestedDataList(jsonArray);
                            MailMergeDataTable parentTable = new MailMergeDataTable(groupName, parentDataList);
                            document.MailMerge.ExecuteNestedGroup(parentTable);
                        }
                        else
                        {
                            // Convert and execute as flat group mail merge
                            List<ExpandoObject> dataList = ConvertToFlatDataList(jsonArray);
                            MailMergeDataTable dataTable = new MailMergeDataTable(groupName, dataList);
                            document.MailMerge.ExecuteGroup(dataTable);
                        }
                    }
                }
                // Step 3: Execute remaining unmerged simple fields
                ExecuteFieldsDictionary(document, simpleFields, onlyUnmerged: true);
            }
        }
        /// <summary>
        /// Extracts simple fields (non-array, non-object properties) from JSON object
        /// </summary>
        private Dictionary<string, string> ExtractSimpleFieldsFromJson(JObject jsonObject)
        {
            Dictionary<string, string> simpleFields = new Dictionary<string, string>();

            foreach (var property in jsonObject.Properties())
            {
                if (!(property.Value is JArray) && !(property.Value is JObject))
                {
                    simpleFields[property.Name] = property.Value?.ToString() ?? string.Empty;
                }
            }

            return simpleFields;
        }
        /// <summary>
        /// Executes mail merge using a dictionary of field names and values.
        /// Can optionally check for unmerged fields only.
        /// </summary>
        private void ExecuteFieldsDictionary(WordDocument document, Dictionary<string, string> fields, bool onlyUnmerged = false)
        {
            if (fields == null || fields.Count == 0)
            {
                return;
            }
            Dictionary<string, string> fieldsToExecute = fields;
            // If onlyUnmerged is true, filter to only unmerged fields in the document
            if (onlyUnmerged)
            {
                string[] unmergedFields = document.MailMerge.GetMergeFieldNames();

                if (unmergedFields == null || unmergedFields.Length == 0)
                {
                    return;
                }

                fieldsToExecute = new Dictionary<string, string>();
                foreach (var unmergedField in unmergedFields)
                {
                    if (fields.ContainsKey(unmergedField))
                    {
                        fieldsToExecute[unmergedField] = fields[unmergedField];
                    }
                }
            }

            // Execute mail merge if we have fields to merge
            if (fieldsToExecute.Count > 0)
            {
                document.MailMerge.Execute(fieldsToExecute.Keys.ToArray(), fieldsToExecute.Values.ToArray());
            }
        }
        
        /// <summary>
        /// Flattens nested JSON objects into dot-notation key-value pairs
        /// </summary>
        private Dictionary<string, string> FlattenJson(JObject obj, string prefix = "")
        {
            var result = new Dictionary<string, string>();

            foreach (var property in obj.Properties())
            {
                string key = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";

                if (property.Value is JObject nestedObject)
                {
                    // Recursively flatten nested objects
                    var nested = FlattenJson(nestedObject, key);
                    foreach (var kvp in nested)
                    {
                        result[kvp.Key] = kvp.Value;
                    }
                }
                else if (property.Value is JArray)
                {
                    // Skip arrays - handled separately
                    continue;
                }
                else
                {
                    result[key] = property.Value?.ToString() ?? string.Empty;
                }
            }

            return result;
        }
        /// <summary>
        /// Extracts data from an input PDF stream and conditionally applies digital signatures
        /// at detected keyword locations using a provided signature image stream.
        /// </summary>
        private MemoryStream DataExtractor(Stream inputStream, IFormFile signatureImage, string signatureKeywordsInput, bool enableDigitalSign)
        {
            // Initialize the extractor with required detection settings
            var extractor = new DataExtractor { EnableFormDetection = false, EnableTableDetection = true, ConfidenceThreshold = 0.6 };
            // Extract PDF document from the input stream
            PdfLoadedDocument document = extractor.ExtractDataAsPdfDocument(inputStream);
            // If digital signing is disabled, return the extracted PDF as-is
            if (!enableDigitalSign)
            {
                _logger.LogInformation("Digital signature is not enabled. Returning PDF without signatures.");
                var outMs = new MemoryStream();
                document.Save(outMs);
                document.Close(true);
                outMs.Position = 0;
                return outMs;
            }
            // Get the signature image stream from the uploaded file
            Stream signatureStream = GetSignatureImageStream(signatureImage);
            // If signature image is not available, return PDF without applying signature
            if (signatureStream == null)
            {
                _logger.LogWarning("Digital signature enabled but no signature image found. Returning PDF without signatures.");
                var outMs = new MemoryStream();
                document.Save(outMs);
                document.Close(true);
                outMs.Position = 0;
                return outMs;
            }

            try
            {
                // Use default keywords if none are provided
                string[] signatureKeywords = string.IsNullOrWhiteSpace(signatureKeywordsInput)
                    ? new[] { "Sign", "WITNESS", "AuthorizedSign","Signature" }
                    : signatureKeywordsInput.Split(',').Select(k => k.Trim()).ToArray();
                // Iterate through each page in the document
                for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++)
                {
                    PdfPageBase page = document.Pages[pageIndex];
                    TextLineCollection textLines;
                    // Extract text lines from the page
                    page.ExtractText(out textLines);
                    // Iterate through each word on the page
                    foreach (TextLine line in textLines.TextLine)
                    {
                        foreach (TextWord word in line.WordCollection)
                        {
                            // Skip words that do not match any signature keyword
                            if (!signatureKeywords.Any(k => word.Text.Contains(k, StringComparison.Ordinal)))
                                continue;
                            // Calculate signature position above the keyword
                            RectangleF bounds = word.Bounds;
                            float signatureX = bounds.X;
                            float signatureY = bounds.Y - bounds.Height - 10;
                            float signatureWidth = 80;
                            float signatureHeight = 20;

                            try
                            {
                                // Load digital certificate
                                using System.IO.FileStream cert = new System.IO.FileStream(Path.GetFullPath("PDF.pfx"), System.IO.FileMode.Open, System.IO.FileAccess.Read);
                                PdfCertificate pdfCert = new PdfCertificate(cert, "syncfusion");
                                // Create and configure the PDF signature
                                PdfSignature signature = new PdfSignature(document, page, pdfCert, "Signature");
                                signature.Bounds = new RectangleF(signatureX, signatureY, signatureWidth, signatureHeight);

                                // Load signature image directly from stream
                                signatureStream.Position = 0; // Reset stream position for each use
                                PdfBitmap signatureImageBitmap = new PdfBitmap(signatureStream);
                                signature.Appearance.Normal.Graphics.DrawImage(signatureImageBitmap, 0, 0, signatureWidth, signatureHeight);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "Error adding signature to PDF");
                            }
                        }
                    }
                }

                var outputMs = new MemoryStream();
                document.Save(outputMs);
                document.Close(true);
                outputMs.Position = 0;
                return outputMs;
            }
            finally
            {
                // Clean up signature stream
                signatureStream?.Dispose();
            }
        }

        /// <summary>
        /// Gets signature image as a stream from user upload or default signature
        /// </summary>
        private Stream GetSignatureImageStream(IFormFile signatureImage)
        {
            // If user provided an image, return its stream
            if (signatureImage != null && signatureImage.Length > 0)
            {
                _logger.LogInformation("Using user-provided signature image stream.");
                var memoryStream = new MemoryStream();
                signatureImage.OpenReadStream().CopyTo(memoryStream);
                memoryStream.Position = 0;
                return memoryStream;
            }

            // No user image - use default signature from project
            string defaultImagePath = Path.Combine(_hostingEnvironment.ContentRootPath, "Signature.png");

            if (System.IO.File.Exists(defaultImagePath))
            {
                _logger.LogInformation($"Using default signature image from: {defaultImagePath}");
                try
                {
                    var fileStream = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read);
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    fileStream.Dispose();
                    memoryStream.Position = 0;
                    return memoryStream;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error loading default signature image: {defaultImagePath}");
                    return null;
                }
            }

            _logger.LogWarning($"Default signature image not found at: {defaultImagePath}");
            return null;
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
        private byte[] SplitByPageBreak(WordDocument document, IFormFile signatureImage, string signatureKeywords, bool enableDigitalSign)
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
                            MemoryStream signedPdfStream = DataExtractor(pdfStream, signatureImage, signatureKeywords, enableDigitalSign);

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
