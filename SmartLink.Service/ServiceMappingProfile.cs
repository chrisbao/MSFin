﻿using AutoMapper;
using SmartLink.Entity;

namespace SmartLink.Service
{
    public class ServiceMappingProfile : Profile
    {
        public override string ProfileName
        {
            get
            {
                return "DomainModelMappings";
            }
        }

        public ServiceMappingProfile()
        {
            CreateMap<SourcePoint, PublishedHistory>()
                .ForMember(dest => dest.Id, opt => opt.Ignore())
                .ForMember(dest => dest.PublishedUser, opt => opt.MapFrom(source => source.Creator))
                .ForMember(dest => dest.PublishedDate, opt => opt.MapFrom(source => source.Created));
            CreateMap<PublishStatusEntity, PublishStatusItem>()
                .ForMember(dest => dest.Status, opt => opt.MapFrom(source => source.Status.ToString()));
        }
    }
}
