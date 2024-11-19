using CMCS_MVC_App.Models;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;

namespace CMCS_MVC_App.Data
{
    public class ApplicationDbContext : IdentityDbContext
    {
        public ApplicationDbContext(DbContextOptions options) : base(options)
        {  
        }

        public DbSet<Claim> Claims { get; set; }
        public DbSet<ApplicationUser> ApplicationUser { get; set; }
        
    }
}
