using ClosedXML;
using ClosedXML.Report;
using CMCS_MVC_App.Data;
using CMCS_MVC_App.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
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
using CMCS_MVC_App.ViewModels;

namespace CMCS_MVC_App.Controllers.HRController
{
    [Authorize(Roles = "HR")]
    public class HRController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<IdentityUser> _userManager;

        public HRController(ApplicationDbContext context, UserManager<IdentityUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        //This is the dashboard for HR, containing the buttons
        //for actions that they are able to perform
        public IActionResult DashboardForHR()
        {
            return View();
        }

        //Code Attribution for original 'Listing All Users' Functionality:
        //Author: kudvenkat
        //Website 1: csharp-video-tutorials.blogspot.com
        //Website 1 Link: https://csharp-video-tutorials.blogspot.com/2019/07/list-all-users-from-aspnet-core.html
        //Also from:
        //YouTube Channel: @Csharp-video-tutorialsBlogspot on YouTube.com
        //YouTube video link: https://www.youtube.com/watch?v=OMX0UiLpMSA&list=PL6n9fhu94yhVkdrusLaQsfERmL_Jh4XmU&index=87
        //Date Accessed: 22 November 2024

        //Code Attribution for modifications to 'Listing All Users' functionality:
        //Author: Open AI
        //AI Model: ChatGPT 4.o
        //Link: https://chatgpt.com (unable to send specific chat link because I attached images in the chat)
        //Date Accessed: 22 November 2024


        [HttpGet]
        public IActionResult ManageLecturers()
        {
            // Fetch all users from the database
            var users = _userManager.Users.ToList();

            // Filter users who are in the "Lecturers" role
            var lecturers = users.Where(user => _userManager.IsInRoleAsync(user, "Lecturer").Result).ToList();

            return View(lecturers);
        }


        //Code Attribution for 'Editing Lecturer' Functionality:
        //Author: kudvenkat
        //Website 1: csharp-video-tutorials.blogspot.com
        //Website 1 Link: https://csharp-video-tutorials.blogspot.com/2019/07/edit-identity-user-in-aspnet-core.html
        //Also From:
        //YouTube Channel: kudvenkat (@Csharp-video-tutorialsBlogspot) on YouTube.com
        //YouTube video link: https://www.youtube.com/watch?v=QYlIfH8qyrU
        //Date Accessed: 22 November 2024

        //Allows HR Users to edit data of lecturers:
        //This is specifically the action for retrieving the data of a lecturer
        [HttpGet]
        public async Task<IActionResult> EditLecturer(string id)
        {
            var user = await _userManager.FindByIdAsync(id);

            if (user == null)
            {
                ViewBag.ErrorMessage = $"User with Id = {id} cannot be found";
                return View("NotFound");
            }

            var model = new EditLecturerViewModel
            {
                Id = user.Id,
                Email = user.Email,
                UserName = user.UserName,
                PhoneNumber = user.PhoneNumber
            };

            return View(model);
        }

        //Allows HR Users to edit data of lecturers
        //This is specifically the action for posting the updated data of a lecturer
        [HttpPost]
        public async Task<IActionResult> EditLecturer(EditLecturerViewModel model)
        {
            var user = await _userManager.FindByIdAsync(model.Id);

            if (user == null)
            {
                ViewBag.ErrorMessage = $"User with Id = {model.Id} cannot be found";
                return View("NotFound");
            }
            else
            {
                user.Email = model.Email;
                user.UserName = model.UserName;
                user.PhoneNumber = model.PhoneNumber;

                var result = await _userManager.UpdateAsync(user);

                if (result.Succeeded)
                {
                    return RedirectToAction("ManageLecturers");
                }

                foreach (var error in result.Errors)
                {
                    ModelState.AddModelError("", error.Description);
                }

                return View(model);
            }
        }


        //This the view that is displayed to the HR user
        //when they click on the 'Generate Report' button
        //on the HR dashboard
        public IActionResult ReportView()
        {
            return View();
        }


        

  


        //Code Attribution for the Downloading and Generation of Approved Claims Report in XML form:
        //Author: Open AI
        //Chat Model: ChatGPT 4.o
        //Link: https://chatgpt.com/ (unable to send specific chat link because I attached images in the chat)
        //Date Accessed: 21 November 2024

        //Allows HR user to download the generated report as an (Open) XML File
        public IActionResult DownloadReport()
        {
            //Fetch only approved claims from the database
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

            //If no claims are found
            if (!approvedClaims.Any())
            {
                TempData["Message"] = "No approved claims available.";

                //Redirect to a page (e.g., HRView) to notify the user
                return RedirectToAction("DashboardForHR");  
            }

            //Cast to List<dynamic> to make it compatible with GenerateExcelReport
            var dynamicClaims = approvedClaims.Select(c => (dynamic)c).ToList();

            //Calls the GenerateExcelReport method
            //to generate the report on approved claims
            var reportFile = GenerateExcelReport(dynamicClaims);

            //Return the Excel file as a downloadable file
            return File(reportFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ApprovedClaimsReport.xlsx");
        }


        //Generates the Report XML File itself
        private byte[] GenerateExcelReport(List<dynamic> approvedClaims)
        {
            //Create a new workbook and worksheet
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Approved Claims Report");

            //Add the "Approved Claims Report" title above the table
            worksheet.Cell(1, 1).Value = "Approved Claims Report";

            //Style the title
            var titleCell = worksheet.Cell(1, 1);
            //Bold text for the title
            titleCell.Style.Font.SetBold();
            //Set font size for the title
            titleCell.Style.Font.SetFontSize(16);
            //Center-align the title
            titleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //Vertically center the title
            titleCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            //Set background color to the custom blue
            titleCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(79, 129, 189));
            //White font color for the title
            titleCell.Style.Font.SetFontColor(XLColor.White);
            //Set height for the title row to make it more prominent
            worksheet.Row(1).Height = 30;  

            //Merge cells for the title to span across all columns (from column 1 to 10)
            worksheet.Range(1, 1, 1, 10).Merge();

            //Set the header row (titles for the columns)
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

            //Fill the data rows (starting from row 3 because row 2 is for column headers)
            int row = 3;
            foreach (var claim in approvedClaims)
            {
                // Set the values in each cell
                worksheet.Cell(row, 1).Value = claim.ClaimId;
                worksheet.Cell(row, 2).Value = claim.UserId;
                worksheet.Cell(row, 3).Value = claim.SubmissionDate.ToString("yyyy-MM-dd");
                //Add "hrs" for HoursWorked
                worksheet.Cell(row, 4).Value = $"{claim.HoursWorked} hrs";  
                worksheet.Cell(row, 5).Value = claim.HourlyRate;
                worksheet.Cell(row, 6).Value = claim.PaymentAmount;
                worksheet.Cell(row, 7).Value = claim.AdditionalNote ?? string.Empty;
                worksheet.Cell(row, 8).Value = claim.DocumentName ?? string.Empty;
                worksheet.Cell(row, 9).Value = claim.Status;
                worksheet.Cell(row, 10).Value = claim.ApprovalDate?.ToString("yyyy-MM-dd");

                //Format columns for PaymentAmount and HourlyRate with currency (Rand symbol 'R'):
                //Format for HourlyRate
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "R #,##0.00";
                //Format for PaymentAmount
                worksheet.Cell(row, 6).Style.NumberFormat.Format = "R #,##0.00"; 

                //Apply the blue background color to data rows and white text color
                var dataRowRange = worksheet.Range(row, 1, row, 10);
                //Set blue background for data rows
                dataRowRange.Style.Fill.SetBackgroundColor(XLColor.FromArgb(79, 129, 189));
                //Set white font color for data rows
                dataRowRange.Style.Font.SetFontColor(XLColor.White);  

                //Apply borders to each individual cell value in the row (black borders around each cell)
                for (int col = 1; col <= 10; col++)
                {
                    //Black borders around individual cells
                    worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);  
                }

                row++;
            }

            //Auto-size columns based on the longest value in the column
            worksheet.Columns().AdjustToContents();

            //Create a memory stream and save the workbook to it
            using (var memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                return memoryStream.ToArray();
            }
        }

    }
}
