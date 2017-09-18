namespace ContosoO365DocSync.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddDocumentIdToCatalog : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.DestinationCatalogs", "DocumentId", c => c.String());
            AddColumn("dbo.SourceCatalogs", "DocumentId", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.SourceCatalogs", "DocumentId");
            DropColumn("dbo.DestinationCatalogs", "DocumentId");
        }
    }
}
