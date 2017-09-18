/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace ContosoO365DocSync.Service
{
    public class CheckDocumentEntity : TableEntity
    {
        public CheckDocumentEntity()
        {
            this.PartitionKey = Guid.NewGuid().ToString();
            this.RowKey = Guid.NewGuid().ToString();
        }

        public string Comments { get; set; }
    }
}
