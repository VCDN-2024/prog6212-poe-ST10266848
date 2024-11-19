using Microsoft.AspNetCore.Mvc;
using CMCS_MVC_App.Models;
using CMCS_MVC_App.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Authorization;
using System.Security.Claims;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;

namespace CMCS_MVC_App.Controllers.ClaimsController
{
    public class ClaimsController : Controller
    {

        private readonly ApplicationDbContext _context;
        private readonly UserManager<IdentityUser> _userManager;

        public ClaimsController(ApplicationDbContext context, UserManager<IdentityUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        //This essentially serves as the index view for Lecturers
        //to view the claims they have submitted throughout the year
        [Authorize(Roles = "Lecturer")]
        public async Task<IActionResult> LecturerClaim()
        {
            var user = await _userManager.GetUserAsync(User);

            if (await _userManager.IsInRoleAsync(user, "Lecturer"))
            {

                //Code Attribute for the below code segment (the 'Select' part to be specific):
                //Author: Open AI
                //Chat Model: ChatGPT 4.o
                //Link: https://chatgpt.com/share/671229d8-ae70-8002-9f58-60db6a1105d8
                //Date Accessed: 18 October 2024

                var claims = _context.Claims
                                            .Where(c => c.UserId == user.Id)
                                            .Include(c => c.User)
                                            .Select(c => new Models.Claim
                                             {
                                                ClaimId = c.ClaimId,
                                                UserId = c.UserId,
                                                User = c.User,
                                                HoursWorked = c.HoursWorked,
                                                HourlyRate = c.HourlyRate,
                                                PaymentAmount = c.PaymentAmount,
                                                SubmissionDate = c.SubmissionDate,
                                                AdditionalNote = c.AdditionalNote ?? string.Empty,
                                                Status = c.Status ?? "Pending",
                                                DocumentName = c.DocumentName ?? null,
                                                DocumentContent = c.DocumentContent ?? Array.Empty<byte>(),
                                                IsApprovedByPC = c.IsApprovedByPC,
                                                IsApprovedByAM = c.IsApprovedByAM,
                                                ApprovalDate = c.ApprovalDate,
                                                IsPaid = c.IsPaid,
                                                PaymentDate = c.PaymentDate
                                             })
                                            .ToList();

                return View(claims);
            }
            else
            {
                return Unauthorized();
            }
            
        }

        public IActionResult SubmitClaim()
        {
            ViewData["UserId"] = new SelectList(_context.Users, "Id", "UserName");
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SubmitClaim(Models.Claim claim, IFormFile DocumentContent)
        {
            if (!ModelState.IsValid)
            {

                var user = await _userManager.GetUserAsync(User);

                if (user != null)
                {
                    //Store the UserId in the Claim object
                    claim.UserId = user.Id; 
                }


                //Code Attribute for the below File Upload Implementation:
                //Author: Open AI
                //Chat Model: ChatGPT 4.o
                //Link: 
                //Date Accessed: 18 October 2024

                //Defining the maximum file size to be 1Mb
                const long maxFileSize = 1 * 1024 * 1024; 

                // Handle the file upload
                if (DocumentContent != null && DocumentContent.Length > 0)
                {
                    //Checks if the file size exceeds the max limit
                    if (DocumentContent.Length > maxFileSize)
                    {
                        ModelState.AddModelError("DocumentContent", "The file size exceeds the maximum limit of 1 MB.");
                        return View(claim);
                    }
                    else
                    {
                        // Get the document name
                        claim.DocumentName = DocumentContent.FileName;

                        // Convert the document to a byte array
                        using (var memoryStream = new MemoryStream())
                        {
                            await DocumentContent.CopyToAsync(memoryStream);
                            claim.DocumentContent = memoryStream.ToArray();
                        }
                    }   
                }
                else
                {
                    // If no document is uploaded,
                    // DocumentName and DocumentContent
                    // is explicitly set to null
                    claim.DocumentName = null;
                    claim.DocumentContent = null;
                }

                claim.HourlyRate = 450;
                claim.SubmissionDate = DateTime.Now;
                claim.Status = "Pending";
                claim.IsPaid = false;
                claim.PaymentAmount = 0;

                _context.Add(claim);
                await _context.SaveChangesAsync();

                return RedirectToAction(nameof(LecturerClaim));
            }

            ViewData["UserId"] = new SelectList(_context.Users, "Id", "UserName", claim.UserId);
            return View(claim);

        }


        public async Task<IActionResult> DownloadDocument(int id)
        {
            // Retrieve the claim with the specified ClaimId from the database
            var claim = await _context.Claims.FirstOrDefaultAsync(c => c.ClaimId == id);

            if(claim == null)
            {
                // Return a 404 not found response if the claim does not exist
                return NotFound("Claim not found.");
            }

            if (claim.DocumentContent == null || claim.DocumentName == null)
            {
                // Handle the case where the document is not available
                return BadRequest("No supporting document was uploaded for this claim.");
            }


            // Set the content type for PDF files
            var contentType = "application/pdf";

            // Return the file to be downloaded
            return File(claim.DocumentContent, contentType, claim.DocumentName);
        }


        //This essentially serves as the index view for Lecturers
        //to view the claims they have submitted throughout the year
        [Authorize(Roles = "Programme Coordinator, Academic Manager")]
        public async Task<IActionResult> PendingClaims()
        {
            var user = await _userManager.GetUserAsync(User);

            var claims = _context.Claims
                                            .Where(c => c.UserId == user.Id)
                                            .Include(c => c.User)
                                            .Select(c => new Models.Claim
                                            {
                                                ClaimId = c.ClaimId,
                                                UserId = c.UserId,
                                                User = c.User,
                                                HoursWorked = c.HoursWorked,
                                                HourlyRate = c.HourlyRate,
                                                PaymentAmount = c.PaymentAmount,
                                                SubmissionDate = c.SubmissionDate,
                                                AdditionalNote = c.AdditionalNote ?? string.Empty,
                                                Status = c.Status ?? "Pending",
                                                DocumentName = c.DocumentName ?? null,
                                                DocumentContent = c.DocumentContent ?? Array.Empty<byte>(),
                                                IsApprovedByPC = c.IsApprovedByPC,
                                                IsApprovedByAM = c.IsApprovedByAM,
                                                ApprovalDate = c.ApprovalDate,
                                                IsPaid = c.IsPaid,
                                                PaymentDate = c.PaymentDate
                                            })
                                            .ToList();
            return View(claims);
        }

        //[Authorize(Roles = "Programme Coordinator, Academic Manager")]
        //public async Task<IActionResult> PendingClaims()
        //{

        //}


    }
}
