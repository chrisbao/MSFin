namespace SmartLink.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class UpdateRangeIdLength : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.DestinationPoints", "RangeId", c => c.String(maxLength: 255));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.DestinationPoints", "RangeId", c => c.String());
        }
    }
}
