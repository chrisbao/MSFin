/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using SmartLink.Entity;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface IDocumentService
    {
        Task<DocumentUpdateResult> UpdateBookmarkValueAsync(string documentId, IEnumerable<DestinationPoint> destinationPoints, string value);
        Task<DocumentCheckResult> GetDocumentUrlByIdAsync(DocumentCheckResult result);
    }
}