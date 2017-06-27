/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace SmartLink.Service
{
    using SmartLink.Entity;
    using System.Data.Entity;

    public class SmartlinkDbContext : System.Data.Entity.DbContext
    {
        // Your context has been configured to use a 'dbContext' connection string from your application's 
        // configuration file (App.config or Web.config). By default, this connection string targets the 
        // 'SmartLink.Service.dbContext' database on your LocalDb instance. 
        // 
        // If you wish to target a different database and/or database provider, modify the 'dbContext' 
        // connection string in the application configuration file.
        public SmartlinkDbContext()
            : base("name=DefaultConnection")
        {
            //Database.SetInitializer<SmartlinkDbContext>(new SmartlinkDbContextInitializer());
            this.Configuration.LazyLoadingEnabled = false;
        }

        public virtual DbSet<SourceCatalog> SourceCatalogs { get; set; }
        public virtual DbSet<SourcePoint> SourcePoints { get; set; }
        public virtual DbSet<SourcePointGroup> SourcePointGroups { get; set; }
        public virtual DbSet<PublishedHistory> PublishedHistories { get; set; }
        public virtual DbSet<DestinationPoint> DestinationPoints { get; set; }
        public virtual DbSet<DestinationCatalog> DestinationCatalogs { get; set; }
        public virtual DbSet<CustomFormat> CustomFormats { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
        }
    }
}