/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;
using System.Collections.Generic;

namespace SmartLink.Entity
{
    public class PublishSourcePointResult
    {
        public Guid BatchId { get; set; }
        public ICollection<SourcePoint> SourcePoints { get; set; }
    }
}
