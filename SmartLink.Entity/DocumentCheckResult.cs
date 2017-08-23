/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace SmartLink.Entity
{
    public class DocumentCheckResult
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }

        public string DocumentId { get; set; }

        public string DocumentUrl { get; set; }

        public bool IsUpdated { get; set; }

        public DocumentTypes DocumentType { get; set; }
    }

    public enum DocumentTypes
    {
        SourcePoint = 1,
        DestinationPoint = 2
    }
}
