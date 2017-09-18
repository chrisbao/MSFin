/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using ContosoO365DocSync.Entity;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ContosoO365DocSync.Service
{
    public interface IDestinationService
    {
        Task<DestinationPoint> AddDestinationPointAsync(string fileName, string documentId, DestinationPoint destinationPoint);

        Task<DestinationCatalog> GetDestinationCatalogAsync(string fileName, string documentId);

        Task<IEnumerable<DestinationPoint>> GetDestinationPointBySourcePointAsync(Guid sourcePointId);

        Task DeleteDestinationPointAsync(Guid destinationPointId);

        Task DeleteSelectedDestinationPointAsync(IEnumerable<Guid> seletedDestinationPointIds);

        Task<IEnumerable<CustomFormat>> GetCustomFormatsAsync();

        Task<DestinationPoint> UpdateDestinationPointCustomFormatAsync(DestinationPoint destinationPoint);
    }
}