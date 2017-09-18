using System.Linq;

using ContosoO365DocSync.Service;

namespace ContosoO365DocSync.Console
{
    public class DbTest
    {
        public void TestInsertSourceCategory()
        {
            var dbContext = new ContosoO365DocSyncDbContext();
            dbContext.SourceCatalogs.Add(new Entity.SourceCatalog() { Name = "First One" });
            dbContext.SaveChanges();

            dbContext.SourceCatalogs.ToList().ForEach(o =>
            {
                System.Console.WriteLine("SourceCatalog Id:{0}\tName:{1}", o.Id, o.Name);
            });
        }
        
    }
}
