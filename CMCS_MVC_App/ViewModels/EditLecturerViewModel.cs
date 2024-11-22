using System.ComponentModel.DataAnnotations;

namespace CMCS_MVC_App.ViewModels
{
    public class EditLecturerViewModel
    {

        public string Id { get; set; }

        [Required]
        public string UserName { get; set; }

        [Required]
        [EmailAddress]
        public string Email { get; set; }

        [Phone]
        [Display(Name = "Phone number")]
        public string PhoneNumber { get; set; }

        
    }
}
