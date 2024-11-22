using ClosedXML;
using ClosedXML.Report;
using CMCS_MVC_App.Data;
using CMCS_MVC_App.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Authorization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using System.Text;
using iText.Kernel.Pdf;
using iText.Layout.Properties;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Font;

namespace CMCS_MVC_App.Controllers.HRController
{
    public class HRController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<IdentityUser> _userManager;

        public HRController(ApplicationDbContext context, UserManager<IdentityUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }


        public IActionResult DashboardForHR()
        {
            return View();
        }

        public IActionResult ReportView()
        {
            return View();
        }


        //Code Attribution for Viewing Report Functionality:
        //This is also code attribution for the viewing of the report as html
        //Author: Open AI
        //Chat Model: ChatGPT 4.o
        //Link: https://chatgpt.com/ (unable to send specific chat link because I attached images)
        //Date Accessed: 21 November 2024


        //Allows HR user to view the generated report as HTML in a new browser tab
       

        //Code Attribution for the Downloading and Generation of Approved Claims Report in XML form:
        //Author: Open AI
        //Chat Model: ChatGPT 4.o
        //Link: https://chatgpt.com/ (unable to send specific chat link because I attached images)
        //Date Accessed: 21 November 2024

        //Allows HR user to download the generated report as an (Open) XML File
        public IActionResult DownloadReport()
        {
            // Fetch only approved claims from the database
            var approvedClaims = _context.Claims
                .Where(c => c.Status == "Approved")
                .Select(c => new
                {
                    c.ClaimId,
                    c.UserId,
                    c.SubmissionDate,
                    c.HoursWorked,
                    c.HourlyRate,
                    c.PaymentAmount,
                    c.AdditionalNote,
                    c.DocumentName,
                    c.Status,
                    c.ApprovalDate
                })
                .ToList();

            // If no claims are found
            if (!approvedClaims.Any())
            {
                TempData["Message"] = "No approved claims available.";
                return RedirectToAction("DashboardForHR");  // Redirect to a page (e.g., HRView) to notify the user
            }

            // Cast to List<dynamic> to make it compatible with GenerateExcelReport
            var dynamicClaims = approvedClaims.Select(c => (dynamic)c).ToList();

            // Generate the Excel report
            var reportFile = GenerateExcelReport(dynamicClaims);

            // Return the Excel file as a downloadable file
            return File(reportFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ApprovedClaimsReport.xlsx");
        }

        //Generates the Report XML File itself
        private byte[] GenerateExcelReport(List<dynamic> approvedClaims)
        {
            // Create a new workbook and worksheet
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Approved Claims Report");

            // Add the "Approved Claims Report" title above the table
            worksheet.Cell(1, 1).Value = "Approved Claims Report";

            // Style the title
            var titleCell = worksheet.Cell(1, 1);
            titleCell.Style.Font.SetBold();  // Bold text for the title
            titleCell.Style.Font.SetFontSize(16);  // Set font size for the title
            titleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);  // Center-align the title
            titleCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);  // Vertically center the title
            titleCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(79, 129, 189));  // Set background color to the custom blue
            titleCell.Style.Font.SetFontColor(XLColor.White);  // White font color for the title
            worksheet.Row(1).Height = 30;  // Set height for the title row to make it more prominent

            // Merge cells for the title to span across all columns (from column 1 to 10)
            worksheet.Range(1, 1, 1, 10).Merge();

            // Set the header row (titles for the columns)
            worksheet.Cell(2, 1).Value = "ClaimId";
            worksheet.Cell(2, 2).Value = "UserId";
            worksheet.Cell(2, 3).Value = "SubmissionDate";
            worksheet.Cell(2, 4).Value = "HoursWorked";
            worksheet.Cell(2, 5).Value = "HourlyRate";
            worksheet.Cell(2, 6).Value = "PaymentAmount";
            worksheet.Cell(2, 7).Value = "AdditionalNote";
            worksheet.Cell(2, 8).Value = "DocumentName";
            worksheet.Cell(2, 9).Value = "Status";
            worksheet.Cell(2, 10).Value = "ApprovalDate";

            // Style the header row (with the custom blue and white text)
            var headerRange = worksheet.Range(2, 1, 2, 10);
            headerRange.Style.Fill.SetBackgroundColor(XLColor.FromArgb(79, 129, 189));  // Background color set to the custom blue
            headerRange.Style.Font.SetBold();  // Bold text for the header
            headerRange.Style.Font.SetFontColor(XLColor.White);  // White text
            headerRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);  // Center align the header text

            // Apply borders to all cells in the worksheet (including header and data rows)
            worksheet.RangeUsed().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);  // Thin black outer borders
            worksheet.RangeUsed().Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);  // Thin black inner borders for all cells

            // Fill the data rows (starting from row 3 because row 2 is for headers)
            int row = 3;  // Start from the 3rd row as row 2 is for column headers
            foreach (var claim in approvedClaims)
            {
                // Set the values in each cell
                worksheet.Cell(row, 1).Value = claim.ClaimId;
                worksheet.Cell(row, 2).Value = claim.UserId;
                worksheet.Cell(row, 3).Value = claim.SubmissionDate.ToString("yyyy-MM-dd");
                worksheet.Cell(row, 4).Value = $"{claim.HoursWorked} hrs";  // Add "hrs" for HoursWorked
                worksheet.Cell(row, 5).Value = claim.HourlyRate;
                worksheet.Cell(row, 6).Value = claim.PaymentAmount;
                worksheet.Cell(row, 7).Value = claim.AdditionalNote ?? string.Empty;
                worksheet.Cell(row, 8).Value = claim.DocumentName ?? string.Empty;
                worksheet.Cell(row, 9).Value = claim.Status;
                worksheet.Cell(row, 10).Value = claim.ApprovalDate?.ToString("yyyy-MM-dd");

                // Format columns for PaymentAmount and HourlyRate with currency (Rand symbol 'R')
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "R #,##0.00"; // Format for HourlyRate
                worksheet.Cell(row, 6).Style.NumberFormat.Format = "R #,##0.00"; // Format for PaymentAmount

                // Apply the blue background color to data rows and white text color
                var dataRowRange = worksheet.Range(row, 1, row, 10);
                dataRowRange.Style.Fill.SetBackgroundColor(XLColor.FromArgb(79, 129, 189));  // Set blue background for data rows
                dataRowRange.Style.Font.SetFontColor(XLColor.White);  // Set white font color for data rows

                // Apply borders to each individual cell value in the row (black borders around each cell)
                for (int col = 1; col <= 10; col++)
                {
                    worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);  // Black borders around individual cells
                }

                row++;
            }

            // Auto-size columns based on the longest value in the column
            worksheet.Columns().AdjustToContents();

            // Create a memory stream and save the workbook to it
            using (var memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                return memoryStream.ToArray();
            }
        }

    }
}
