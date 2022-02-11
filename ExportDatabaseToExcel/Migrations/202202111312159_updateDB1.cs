namespace ExportDatabaseToExcel.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class updateDB1 : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Charts",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Description = c.String(nullable: false),
                        CodeUd = c.String(nullable: false),
                        Type = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Charts");
        }
    }
}
