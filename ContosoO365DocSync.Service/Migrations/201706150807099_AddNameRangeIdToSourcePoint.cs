namespace ContosoO365DocSync.Service.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class AddNameRangeIdToSourcePoint : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.SourcePoints", "NamePosition", c => c.String());
            AlterColumn("dbo.SourcePoints", "NameRangeId", c => c.String(maxLength: 255));
        }

        public override void Down()
        {
            AlterColumn("dbo.SourcePoints", "NamePosition", c => c.String());
            AlterColumn("dbo.SourcePoints", "NameRangeId", c => c.String());
        }
    }
}
