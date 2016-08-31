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

        [Required(ErrorMessage = "Please enter Full Name.")]
        [Display(Name = "Full Name")]
        public string Name { get; set; }

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

        [Required(ErrorMessage = "Please select Height.")]
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

        [Required(ErrorMessage = "Please enter Age.")]
        [Display(Name = "Age")]
        public int Age { get; set; }

        //[DataType(DataType.EmailAddress, ErrorMessage = "Please enter valid E-mail Address.")]
        [RegularExpression(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$", ErrorMessage = "Please enter a valid e-mail adress")]
        [Display(Name = "Email")]
        public string Email { get; set; }

        [Display(Name = "Anubandh Id")]
        public int? AnubandhId { get; set; }

        public DateTime CreatedDate { get; set; }

        [Display(Name = "Photo")]
        public string ImagePath { get; set; }


    }

    public class DbPersonDetails : DbContext
    {
        public DbSet<PersonDetails> Persons { get; set; }

        public DbSet<UserDetails> Users { get; set; }
    }

    public enum Gender
    {
        Male,
        Female
    }
}