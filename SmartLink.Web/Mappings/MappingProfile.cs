using AutoMapper;
using SmartLink.Entity;
using SmartLink.Web.ViewModel;
using System;
using System.Linq;

namespace SmartLink.Web.Mappings
{
    public class MappingProfile : Profile
    {
        public override string ProfileName
        {
            get
            {
                return "DomainViewModelMappings";
            }
        }
        /// <summary>
        /// mapping the view model to entity
        /// </summary>
        public MappingProfile()
        {
            CreateMap<SourcePointForm, SourcePoint>()
                .ForMember(dest => dest.Groups, opt => opt.MapFrom(source => source.GroupIds.Select(o => new SourcePointGroup() { Id = o })));
            CreateMap<DestinationPointForm, DestinationPoint>()
                .ForMember(dest => dest.ReferencedSourcePoint, opt => opt.MapFrom(source => new SourcePoint() { Id = Guid.Parse(source.SourcePointId) }))
                .ForMember(dest => dest.CustomFormats, opt => opt.MapFrom(source => source.CustomFormatIds.Select(o => new CustomFormat() { Id = o })));
        }
    }
}