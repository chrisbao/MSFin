/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Autofac;
using AutoMapper;
using ContosoO365DocSync.Service;
using ContosoO365DocSync.Web.Mappings;

namespace ContosoO365DocSync.Web
{
    public class AutofacBootstrap
    {
        internal static void Init(ContainerBuilder builder)
        {
            builder.RegisterType<SourceService>().As<ISourceService>().InstancePerRequest();
            builder.RegisterType<DestinationService>().As<IDestinationService>().InstancePerRequest();
            builder.RegisterType<ContosoO365DocSyncDbContext>().AsSelf().InstancePerRequest();
            builder.RegisterType<ConfigService>().As<IConfigService>().SingleInstance();
            builder.RegisterType<AzureStorageService>().As<IAzureStorageService>().SingleInstance();
            builder.RegisterType<LogService>().As<ILogService>().SingleInstance();
            builder.RegisterType<MailService>().As<IMailService>().SingleInstance();
            builder.RegisterType<UserProfileService>().As<IUserProfileService>().InstancePerRequest();

            var mapperConfiguration = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new MappingProfile());
                cfg.AddProfile(new ServiceMappingProfile());
                // This list is keep on going...
            });
            var mapper = mapperConfiguration.CreateMapper();
            builder.RegisterInstance(mapper).As<IMapper>().SingleInstance();
        }
    }
}