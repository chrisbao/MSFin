namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddRangeIdToDestination : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.DestinationPoints", "RangeId", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.DestinationPoints", "RangeId");
        }
    }
}