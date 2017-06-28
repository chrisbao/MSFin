/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using SmartLink.Entity;
using System;
using System.Collections.Generic;
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