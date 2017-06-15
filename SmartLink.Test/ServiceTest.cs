using Autofac;
using Autofac.Extras.Moq;
using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.WindowsAzure.Storage.Table;
using SmartLink.Common;
using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;

namespace SmartLink.Service.Tests
{
    [TestClass()]
    public class ServiceTest
    {
        static private IContainer Container { get; set; }
        static private AutoMock MockContainer { get; set; }

        [ClassInitialize]
        static public void Initial(TestContext testContext)
        {
            var builder = new ContainerBuilder();

            builder.RegisterType<SourceService>().As<ISourceService>().InstancePerDependency();
            builder.RegisterType<DestinationService>().As<IDestinationService>().InstancePerDependency();
            builder.RegisterType<SmartlinkDbContext>().AsSelf().InstancePerDependency();
            builder.RegisterType<ConfigService>().As<IConfigService>().SingleInstance();
            builder.RegisterType<AzureStorageService>().As<IAzureStorageService>().SingleInstance();
            builder.RegisterType<LogService>().As<ILogService>().SingleInstance();
            builder.RegisterType<MailService>().As<IMailService>().SingleInstance();
            builder.RegisterType<UserProfileService>().As<IUserProfileService>().InstancePerDependency();
            builder.RegisterType<DocumentService>().As<IDocumentService>().InstancePerDependency();
            var mapperConfiguration = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new ServiceMappingProfile());
                //This list is keep on going...

            });
            var mapper = mapperConfiguration.CreateMapper();
            builder.RegisterInstance(mapper).As<IMapper>().SingleInstance();
            Container = builder.Build();

            Database.SetInitializer(new MigrateDatabaseToLatestVersion<SmartlinkDbContext, SmartLink.Service.Migrations.Configuration>());
            //InitDatabase();

            MockContainer = AutoMock.GetLoose();
            MockContainer.Mock<IUserProfileService>().Setup(x => x.GetCurrentUser()).Returns(new UserProfile() { Email = "UT@canviz.com", Username = "UT" });
            MockContainer.Mock<IConfigService>();
            //MockContainer.Mock<ILogService>().Setup(x => x.WriteLog(It.IsAny<LogEntity>()));
        }

        //[ClassCleanup]
        static public void CleanUp()
        {
            using (var dbContext = new SmartlinkDbContext())
            {
                dbContext.DestinationPoints.RemoveRange(dbContext.DestinationPoints.Where(o => o.Catalog.Name.StartsWith("https://cand3.onmicrosoft.com/test/DestinationCatalog")));
                dbContext.DestinationCatalogs.RemoveRange(dbContext.DestinationCatalogs.Where(o => o.Name.StartsWith("https://cand3.onmicrosoft.com/test/DestinationCatalog")));
                dbContext.SourcePoints.RemoveRange(dbContext.SourcePoints.Where(o => o.Name.StartsWith("SourcePoint")));
                dbContext.SourceCatalogs.RemoveRange(dbContext.SourceCatalogs.Where(o => o.Name.StartsWith("https://cand3.onmicrosoft.com/test/SourceCatalog")));
                dbContext.SaveChanges();
            }
        }

        static private void InitDatabase()
        {
            var dbContext = new SmartlinkDbContext();

            var SourceCatalog = new SourceCatalog() { Name = "https://cand3.onmicrosoft.com/test/SourceCatalog1.xlsx" };
            dbContext.SourceCatalogs.Add(SourceCatalog);

            var sourcePoints = new SourcePoint[500];
            var destinationCatalog = new DestinationCatalog[500];
            for (int i = 0; i < 500; i++)
            {
                sourcePoints[i] = new SourcePoint()
                {
                    Name = $"SourcePoint{i}",
                    Catalog = SourceCatalog,
                    RangeId = $"Range{i}",
                    Creator = $"Creator{i}",
                    Position = $"Position{i}",
                    Value = $"Value{i}",
                    Status = SourcePointStatus.Created,
                    Created = DateTime.UtcNow
                };
                dbContext.SourcePoints.Add(sourcePoints[i]);

                destinationCatalog[i] = new DestinationCatalog() { Name = $"https://cand3.onmicrosoft.com/test/DestinationCatalog{i}.docx" };
                dbContext.DestinationCatalogs.Add(destinationCatalog[i]);
                destinationCatalog[i].DestinationPoints.Add(new DestinationPoint() { Catalog = destinationCatalog[i], RangeId = $"Range{i}", Creator = $"Creator{i}", Created = DateTime.UtcNow, ReferencedSourcePoint = sourcePoints[i] });
            }
            dbContext.SaveChanges();
        }

        //[TestMethod]
        public void Publish500SourcePointTest()
        {
            using (var dbContext = new SmartlinkDbContext())
            {
                var azureService = Container.Resolve<IAzureStorageService>();
                var sourceService = new SourceService(
                    dbContext,
                    Container.Resolve<IMapper>(),
                    azureService,
                    MockContainer.Create<ILogService>(),
                    MockContainer.Create<IUserProfileService>());
                var sourcePoints = Container.Resolve<SmartlinkDbContext>().SourcePoints.Where(o => o.Name.StartsWith("SourcePoint")).Take(500).ToArray();

                var publishSourcePointForm = new PublishSourcePointForm[500];
                for (int i = 0; i < 500; i++)
                {
                    publishSourcePointForm[i] = new PublishSourcePointForm() { SourcePointId = sourcePoints[i].Id, CurrentValue = sourcePoints[i].Value, Position = sourcePoints[i].Position };
                }
                var tasks = sourceService.PublishSourcePointListAsync(publishSourcePointForm).Result;

                var query = new TableQuery<PublishStatusEntity>() { FilterString = $"PartitionKey eq '{tasks.BatchId}'" };
                var value = azureService.GetTable(Constant.PUBLISH_TABLE_NAME).ExecuteQuery(query).ToArray();
                Assert.AreEqual(value.Count(), tasks.SourcePoints.Count());
                var except = tasks.SourcePoints.Select(o => o.Id).Except(value.Select(x => Guid.Parse(x.SourcePointId))).ToArray();
                Assert.IsTrue(except.Length == 0);
            }
        }
        [TestMethod()]
        public void UpdateSdtBlockTest()
        {
            var destinationPoints = (new List<DestinationPoint>()); //4148 tag
            destinationPoints.Add(new DestinationPoint() { Id = Guid.Parse("bf21e129-09f7-4809-b8f6-794859cb33a2"),RangeId= "6df20cc4 - 1e0b - 2cdc - 699e-3f4e8d902670" });

            var value = "6%";
            using (var stream = new FileStream("DK - Press Release FY17Q2_4.docx", FileMode.Open))
            {
                var result = new DocumentUpdateResult();
                DocumentService.UpdateStream(destinationPoints, value, stream, result);
                stream.Close();
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("DK - Press Release FY17Q2_4.docx", true))
            {
                var items = wordDoc.MainDocumentPart.Document.Descendants<SdtBlock>().Where(
                    o =>
                    {
                        var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                        if (tagedItem != null)
                        {
                            return tagedItem.Val == destinationPoints[0].RangeId;
                        }
                        return false;
                    });
                Assert.IsTrue(items.Count() == 1);

                Assert.AreEqual(value, items.First().InnerText);
            }
        }
    }
}

