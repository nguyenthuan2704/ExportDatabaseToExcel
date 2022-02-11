using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace ExportDatabaseToExcel.Models
{
    [Table("Charts")]
    public class Chart
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string Description { get; set; }
        [Required]
        public string CodeUd { get; set; }
        [Required]
        public string Type { get; set; }
    }
}