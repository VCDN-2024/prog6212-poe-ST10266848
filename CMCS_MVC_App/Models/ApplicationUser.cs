using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;

namespace CMCS_MVC_App.Models
{
    public class ApplicationUser : IdentityUser
    {

        //Code Attribution for the implementation of AspNetCore Identity, Roles, Authorization
        //,ApplicationUser (adding additional fields to AspNetUsers table):
        //Author: Digital TechJoint
        //Source: Two YouTube Videos
        //Video 1 Link: https://www.youtube.com/watch?v=ghzvSROMo_M
        //Video 2 Link: https://www.youtube.com/watch?v=qvsWwwq2ynE
        //Date Accessed: 15 October 2024


        //Attributes for each of the below properties are contained with Register.cshtml.cs
        public string FirstName { get; set; } = null!;
        public string LastName { get; set; } = null!;
        public string BankName { get; set; } = null!;
        public string AccountHolderName { get; set; } = null!;
        public string AccountNumber { get; set; }


    }
}
