using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface IDestinationService
    {
        Task<DestinationPoint> AddDestinationPoint(string fileName, DestinationPoint destinationPoint);
        Task<DestinationCatalog> GetDestinationCatalog(string fileName);
        Task<IEnumerable<DestinationPoint>> GetDestinationPointBySourcePoint(Guid sourcePointId);
        Task DeleteDestinationPoint(Guid destinationPointId);
        Task DeleteSelectedDestinationPoint(IEnumerable<Guid> seletedDestinationPointIds);

        Task<IEnumerable<CustomFormat>> GetCustomFormats();
    }
}
