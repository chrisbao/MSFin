﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace ContosoO365DocSync.Entity
{
    public enum PublishStatus
    {
        InProgess,
        Completed,
        Error
    }
    public class PublishStatusItem
    {
        public string PublishBatchId { get; set; }

        public string SourcePointId { get; set; }

        public string Status { get; set; }

        public string ErrorSummary { get; set; }

        public string ErrorDetail { get; set; }
    }
}