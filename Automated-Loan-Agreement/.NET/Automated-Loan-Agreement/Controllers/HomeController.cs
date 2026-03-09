using Automated_Loan_Agreement.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Diagnostics;
using System.IO;
using System.Reflection.Metadata;

namespace Automated_Loan_Agreement.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        public ActionResult AutomatedLoanAgreement()
        {         
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<AgreementDetails> agreementList = GetAgreementDeatils();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("AgreementDetails", agreementList);
                    //Enable the boolean to remove empty paragraph and start each record in new page.
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);

                    //Instantiation of DocIORenderer for Word to PDF conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document
                        PdfDocument pdfDocument = render.ConvertToPDF(document);

                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save(stream);
                        stream.Position = 0;

                        //Download PDF document in the browser.
                        return File(stream, "application/pdf", "Sample.pdf");
                    }
                }
            }
        }
       
        /// <summary>
        /// Gets the agreement details to perform mail merge.
        /// </summary>
        public static List<AgreementDetails> GetAgreementDeatils()
        {
            List<AgreementDetails> details = new List<AgreementDetails>();
            details.Add(new AgreementDetails("LA-2026-10001", "March 15, 2026", "California", "Jennifer Anderson", "2458 Sunset Boulevard, Apt 3B", "Los Angeles", "California", "90028", "+1-213-555-0142", "jennifer.anderson@email.com", "Home Loan - Fixed Rate", "$450,000", "6.75", "360", "$2,918", "$2,250", "Residential Property", "3-bedroom house at 2458 Sunset Blvd, Los Angeles, CA", "$585,000", "April 15, 2026", "April 15, 2056", "2% of outstanding principal", "$75", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10006", "March 20, 2026", "Texas", "Robert Johnson", "1200 Commerce Street, Suite 400", "Dallas", "Texas", "75202", "+1-214-555-3456", "robert.johnson@startup.com", "Small Business Startup Loan", "$120,000", "10.25", "72", "$1,986", "$1,800", "Business Assets", "Office furniture, computers, and inventory", "$45,000", "April 20, 2026", "April 20, 2032", "3.5% of outstanding principal", "$85", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10007", "March 22, 2026", "New York", "Emily Davis", "350 Fifth Avenue, Apt 12C", "New York", "New York", "10118", "+1-212-555-6789", "emily.davis@email.com", "Personal Loan - Home Renovation", "$85,000", "8.50", "84", "$1,312", "$1,275", "Residential Property", "Apartment at 350 Fifth Avenue, New York, NY", "$620,000", "April 22, 2026", "April 22, 2033", "2% of outstanding principal", "$65", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10008", "March 25, 2026", "Florida", "Michael Thompson", "789 Ocean Drive, Unit 5", "Miami", "Florida", "33139", "+1-305-555-9012", "michael.thompson@email.com", "Vehicle Loan - New Car", "$55,000", "7.25", "60", "$1,097", "$825", "Motor Vehicle", "2026 Ford Explorer XLT", "$55,000", "April 25, 2026", "April 25, 2031", "No prepayment penalty", "$45", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10009", "March 28, 2026", "Illinois", "Sophia Martinez", "233 S Wacker Drive, Suite 800", "Chicago", "Illinois", "60606", "+1-312-555-3478", "sophia.martinez@business.com", "Business Expansion Loan", "$200,000", "9.00", "120", "$2,534", "$3,000", "Commercial Property", "Office space at 233 S Wacker Drive, Chicago, IL", "$350,000", "April 28, 2026", "April 28, 2036", "3% of outstanding principal", "$90", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10010", "March 30, 2026", "Washington", "Daniel Wilson", "1420 Harbor Ave SW, Unit 3", "Seattle", "Washington", "98116", "+1-206-555-7654", "daniel.wilson@email.com", "Investment Property Loan", "$520,000", "6.50", "360", "$3,288", "$3,900", "Investment Property", "Condo at 1420 Harbor Ave SW, Seattle, WA", "$680,000", "April 30, 2026", "April 30, 2056", "1.5% of outstanding principal", "$80", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));

            return details;
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
    /// <summary>
    /// Represents a class to maintain agreement details.
    /// </summary>
    public class AgreementDetails
    {
        public string AgreementNumber { get; set; }
        public string AgreementDate { get; set; }
        public string State { get; set; }
        public string BorrowerName { get; set; }
        public string BorrowerAddress { get; set; }
        public string BorrowerCity { get; set; }
        public string BorrowerState { get; set; }
        public string BorrowerZip { get; set; }
        public string BorrowerPhone { get; set; }
        public string BorrowerEmail { get; set; }
        public string LoanProduct { get; set; }
        public string LoanAmount { get; set; }
        public string InterestRate { get; set; }
        public string LoanTerm { get; set; }
        public string MonthlyPayment { get; set; }
        public string ProcessingFee { get; set; }
        public string CollateralType { get; set; }
        public string CollateralDescription { get; set; }
        public string CollateralValue { get; set; }
        public string RepaymentStartDate { get; set; }
        public string MaturityDate { get; set; }
        public string PrepaymentPenalty { get; set; }
        public string LatePaymentFee { get; set; }
        public string WitnessName { get; set; }
        public string WitnessDate { get; set; }
        public string BorrowerDate { get; set; }
        public string LenderDate { get; set; }

        public AgreementDetails(string agreementNumber, string agreementDate, string state, string borrowerName, string borrowerAddress, string borrowerCity, string borrowerState, string borrowerZip, string borrowerPhone, string borrowerEmail, string loanProduct, string loanAmount,
            string interestRate, string loanTerm, string monthlyPayment, string processingFee, string collateralType, string collateralDescription, string collateralValue, string repaymentStartDate, string maturityDate, string prepaymentPenalty, string latePaymentFee, string witnessDate, string borrowerDate, string lenderDate)
        {
            AgreementNumber = agreementNumber;
            AgreementDate = agreementDate;
            State = state;
            BorrowerName = borrowerName;
            BorrowerAddress = borrowerAddress;
            BorrowerCity = borrowerCity;
            BorrowerState = borrowerState;
            BorrowerZip = borrowerZip;
            BorrowerPhone = borrowerPhone;
            BorrowerEmail = borrowerEmail;
            LoanProduct = loanProduct;
            LoanAmount = loanAmount;
            InterestRate = interestRate;
            LoanTerm = loanTerm;
            MonthlyPayment = monthlyPayment;
            ProcessingFee = processingFee;
            CollateralType = collateralType;
            CollateralDescription = collateralDescription;
            CollateralValue = collateralValue;
            RepaymentStartDate = repaymentStartDate;
            MaturityDate = maturityDate;
            PrepaymentPenalty = prepaymentPenalty;
            LatePaymentFee = latePaymentFee;
            WitnessDate = witnessDate;
            BorrowerDate = borrowerDate;
            LenderDate = lenderDate;
        }
    }
}
