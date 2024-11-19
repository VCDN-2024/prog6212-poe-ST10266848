using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;

namespace CMCS_MVC_App.Models
{
    public class Claim
    {
        //Serves as the Primary Key for Claim
        [Key]
        public int ClaimId { get; set; }

        //Serves as the Foreign Key for User
        [Required]
        public string UserId { get; set; } = null!;

        public IdentityUser User { get; set; } = null!;

        //Stores the number of hours that the lecturer has worked or claims to have worked
        [Required(ErrorMessage = "Hours Worked Is Required")]
        [Range(1, 160, ErrorMessage = "Can Only Claim For Anywhere Between 1 hour (minimum) and 150 hours (maximum) Per Month!")]
        public int HoursWorked { get; set; }

        //Stores the hourly rate of all lecturers
        //Here, it is assumed that all lecturers have the same hourly rate
        //It is also pre-defined as to prevent the lecturer from selecting
        //a false value
        public double HourlyRate { get; set; }

        //The result of the calculation involving both of the above fields
        //will be stored here
        public double? PaymentAmount { get; set; }

        //Stores date when claim was submitted by lecturer
        public DateTime SubmissionDate { get; set; }

        //Stores any additional notes that the lecturer may have,
        //regarding their monthly claim.
        //Not marked as Required as it is optional.
        [StringLength(750, ErrorMessage = "Max Character Limit Reached!")]
        public string? AdditionalNote { get; set; }

        //The status of the claim would be stored here
        // There are 3 types of claim statuses (Pending, Approved, Rejected)
        public string Status { get; set; } = null!;

        //Stores the name of the pdf document, for displaying purposes
        public string? DocumentName { get; set; }

        //Stores the actual content of the pdf document in an appropriate format
        //that can be stored in a local sql database
        public byte[]? DocumentContent { get; set; }

        //Checks whether the claim was approved by the Programme Coordinator
        //Programme Coordinator is abreviated as PC here
        public bool IsApprovedByPC { get; set; } = false;

        //Checks whether the claim was approved by the Academic Manager
        //Academic Manager is abreviated as PC here
        public bool IsApprovedByAM { get; set; } = false;

        //Stores the date when the claim was approved
        public DateTime? ApprovalDate { get; set; }

        //Checks whether the payment amount calculated for the claim
        //has been paid out to the lecturer
        public bool IsPaid { get; set; } = false;

        //Stores the date when a payment is made out to Lecturers
        public DateTime? PaymentDate { get; set; }

    }
}
