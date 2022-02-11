using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace ExportDatabaseToExcel.Models
{
    public class ExportDatabaseToExcelDBContext : DbContext
    {
        public ExportDatabaseToExcelDBContext(): base("name=StrConnect") { }
        public DbSet<Chart> Charts { get; set; }
    }
}