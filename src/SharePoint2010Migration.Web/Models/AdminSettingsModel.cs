using System.ComponentModel.DataAnnotations;

namespace SharePoint2010Migration.Web.Models
{
    public class AdminSettingsModel
    {
        [Required]
        public string SharePointClientDllPath { get; set; }

        [Required]
        public string SharePointRuntimeDllPath { get; set; }

        [Required]
        public string SiteUrl { get; set; }

        [Required]
        public string LibraryTitle { get; set; }

        [Required]
        public string Domain { get; set; }

        [Required]
        public string UserName { get; set; }

        [Required]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        public bool IsClientDllFound { get; set; }

        public bool IsRuntimeDllFound { get; set; }
    }
}
