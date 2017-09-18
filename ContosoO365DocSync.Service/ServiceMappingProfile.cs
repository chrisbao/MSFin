/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using AutoMapper;
using ContosoO365DocSync.Entity;

namespace ContosoO365DocSync.Service
{
    public class ServiceMappingProfile : Profile
    {
        public ServiceMappingProfile()
        {
            CreateMap<SourcePoint, PublishedHistory>()
                .ForMember(dest => dest.Id, opt => opt.Ignore())
                .ForMember(dest => dest.PublishedUser, opt => opt.MapFrom(source => source.Creator))
                .ForMember(dest => dest.PublishedDate, opt => opt.MapFrom(source => source.Created));
            CreateMap<PublishStatusEntity, PublishStatusItem>()
                .ForMember(dest => dest.Status, opt => opt.MapFrom(source => source.Status.ToString()));
        }

        public override string ProfileName
        {
            get
            {
                return "DomainModelMappings";
            }
        }
    }
}