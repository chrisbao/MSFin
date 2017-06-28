/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Collections.Generic;

namespace SmartLink.Entity
{
    public class DocumentUpdateResult
    {
        public bool IsSuccess { get; set; }

        public List<string> Message { get; set; }

        public DocumentUpdateResult()
        {
            Message = new List<string>();
        }
    }
}