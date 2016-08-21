using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace MelavaWebsite.Models
{
    public class PersonDetails
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Display(Name = "Candidate Id")]
        public int Id { get; set; }

        [Required(ErrorMessage = "Please enter Name.")]
        [Display(Name = "Full Name")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Please enter Birth Name or enter '-'.")]
        [Display(Name = "Birth Name")]
        public string BirthName { get; set; }

        [Required(ErrorMessage = "Please enter First Gotra.")]
        [Display(Name = "First Gotra")]
        public string FirstGotra { get; set; }

        [Required(ErrorMessage = "Please enter Second Gotra.")]
        [Display(Name = "Second Gotra")]
        public string SecondGotra { get; set; }

        [Display(Name = "Date of Birth")]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime DateOfBirth { get; set; }

        [Display(Name = "Birth Place")]
        public string BirthPlace { get; set; }

        [Required(ErrorMessage = "Please enter Education Details.")]
        [Display(Name = "Education")]
        public string Education { get; set; }

        [Required(ErrorMessage = "Please enter Height.")]
        [Display(Name = "Height")]
        public decimal Height { get; set; }

        [Required(ErrorMessage = "Please select Gender.")]
        [Display(Name = "Gender")]
        public string Gender { get; set; }

        [Required(ErrorMessage = "Please enter Occupation.")]
        [Display(Name = "Occupation")]
        public string Occupation { get; set; }

        [Required(ErrorMessage = "Please enter Address for communication.")]
        [Display(Name = "Address")]
        public string Address { get; set; }

        [Required(ErrorMessage = "Please enter Contact Number.")]
        [Display(Name = "Contact Number")]
        public string ContactNumber { get; set; }

        [Required(ErrorMessage = "Please enter age.")]
        [Display(Name = "Age")]
        public int Age { get; set; }

        [Required(ErrorMessage = "Please enter Email Address.")]
        [DataType(DataType.EmailAddress, ErrorMessage = "Please enter valid email address.")]
        [Display(Name = "Email")]
        public string Email { get; set; }

        [Display(Name = "Anubandh Id")]
        public long AnubandhId { get; set; }
        public DateTime CreatedDate { get; set; }
    }

    public class DbPersonDetails : DbContext
    {
        public DbSet<PersonDetails> Persons { get; set; }
    }

    public enum Gender
    {
        Male,
        Female
    }
}