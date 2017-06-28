namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class UpdateDestinationPointNameLength : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.DestinationCatalogs", "Name", c => c.String());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.DestinationCatalogs", "Name", c => c.String(maxLength: 255));
        }
    }
}