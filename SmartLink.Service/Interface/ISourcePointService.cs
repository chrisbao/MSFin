using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface ISourceService
    {
        Task<SourcePoint> AddSourcePointAsync(string fileName, SourcePoint sourcePoint);
        Task<SourcePoint> EditSourcePointAsync(int[] groupIds, SourcePoint sourcePoint);
        Task<SourceCatalog> GetSourceCatalogAsync(string fileName);
        Task<int> DeleteSourcePointAsync(Guid sourcePointId);
        Task DeleteSelectedSourcePointAsync(IEnumerable<Guid> selectedSourcePointIds);

        Task<PublishSourcePointResult> PublishSourcePointListAsync(IEnumerable<PublishSourcePointForm> publishSourcePointForms);
        Task<IEnumerable<SourcePointGroup>> GetAllSourcePointGroupAsync();
        IEnumerable<PublishStatusEntity> GetPublishStatus(string batchId);
        Task<IEnumerable<SourceCatalog>> GetAllSourceCatalogAsync();
        Task<PublishedHistory> GetPublishHistoryByIdAsync(Guid publishHistoryId);
    }
}
