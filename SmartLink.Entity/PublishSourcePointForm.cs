/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;

namespace SmartLink.Entity
{
    public class PublishSourcePointForm
    {
        public Guid SourcePointId { get; set; }
        public string CurrentValue { get; set; }
        public string Position { get; set; }
    }
}